using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.IO;
using System.Timers;
using System.Threading;
using System.Reflection;
using System.Net.Mail;
using ImageMagick;
using System.Drawing;

namespace FaxPDFWinService
{
    public partial class FaxPDFService : ServiceBase
    {
        string WatchPath1 = ConfigurationManager.AppSettings["WatchPath1"];

        public FaxPDFService()
        {
            InitializeComponent();
            fileSystemWatcher.Changed += new FileSystemEventHandler(OnChanged);
            fileSystemWatcher.Created += new FileSystemEventHandler(OnCreated);
            fileSystemWatcher.Deleted += new FileSystemEventHandler(OnChanged);
            fileSystemWatcher.Renamed += new RenamedEventHandler(OnRenamed);
        }

        private void writeLog(string FullPath, string FileName, string EventType)
        {
            StreamWriter SW;
            string Logs = ConfigurationManager.AppSettings["Logs"];
            if (Directory.Exists(Logs))
            {
                Logs = System.IO.Path.Combine(Logs, "FaxPDFlog_" + DateTime.Now.ToString("yyyyMMdd") + ".txt");
                if (!File.Exists(Logs))
                {
                    SW = File.CreateText(Logs);
                    SW.Close();
                }
            }
            using (SW = File.AppendText(Logs))
            {
                SW.Write("\r\n");
                if ((EventType == "Created" || EventType == "Changed" || EventType == "Renamed" || EventType == "Converted" || EventType == "Mailed" ))
                {
                    SW.WriteLine(DateTime.Now.ToString("dd-MM-yyyy H:mm:ss") + ": File " + EventType + " with Name: " + FileName + " at this location: " + FullPath);
                }
                else if ((EventType == "Stopped" || EventType == "Started"))
                {
                    SW.WriteLine("Service " + EventType + " at " + DateTime.Now.ToString("dd -MM-yyyy H:mm:ss"));
                }
                else if (EventType == "Error")
                {
                    SW.WriteLine("ERROR: " + DateTime.Now.ToString("dd-MM-yyyy H:mm:ss") + FullPath + FileName);
                }
                SW.Close();
            }
        }

        private static bool GetIdleFile(string path)
        {
            var fileIdle = false;
            const int MaximumAttemptsAllowed = 30;
            var attemptsMade = 0;

            while (!fileIdle && attemptsMade <= MaximumAttemptsAllowed)
            {
                try
                {
                    using (File.Open(path, FileMode.Open, FileAccess.ReadWrite))
                    {
                        fileIdle = true;
                    }
                }
                catch
                {
                    attemptsMade++;
                    Thread.Sleep(100);
                }
            }

            return fileIdle;
        }

        private void convertPDF(string FullPath)
        {
            string FullPdfPath = FullPath.Replace("TIF", "pdf");
            using (MagickImage image = new MagickImage(FullPath))
            {
                image.Write(FullPdfPath);
            }
        }

        private static void SendMail(string FullPath, string FileName)
        {
            int SMTPPort = Convert.ToInt32(ConfigurationManager.AppSettings["SMTPPort"]);
            string SMTPHost = ConfigurationManager.AppSettings["SMTPHost"];
            string SMTPTo = ConfigurationManager.AppSettings["SMTPTo"];
            string SMTPFrom = ConfigurationManager.AppSettings["SMTPFrom"];
            string SMTPSubj = ConfigurationManager.AppSettings["SMTPSubj"];
            string SMTPBody = ConfigurationManager.AppSettings["SMTPBody"];
            bool SMTPUserCreds = Convert.ToBoolean(ConfigurationManager.AppSettings["SMTPUserCreds"]);
            string SMTPUser = ConfigurationManager.AppSettings["SMTPUser"];
            string SMTPPass = ConfigurationManager.AppSettings["SMTPPass"];


            System.Net.Mail.Attachment attachment;

            SmtpClient client = new SmtpClient();
            client.Port = SMTPPort;
            client.Host = SMTPHost;
            client.EnableSsl = false;
            client.Timeout = 40000;
            client.DeliveryMethod = SmtpDeliveryMethod.Network;
            client.UseDefaultCredentials = true;

            client.Credentials = new System.Net.NetworkCredential(SMTPUser, SMTPPass);

            MailMessage mm = new MailMessage(SMTPFrom, SMTPTo, SMTPSubj, SMTPBody);
            attachment = new System.Net.Mail.Attachment(FullPath);
            mm.Attachments.Add(attachment);
            mm.BodyEncoding = UTF8Encoding.UTF8;
            mm.DeliveryNotificationOptions = DeliveryNotificationOptions.OnFailure;

            client.Send(mm);
            attachment.Dispose();
        }

        private void OnChanged(object source, FileSystemEventArgs e)
        {
            try
            {
                writeLog(e.FullPath, e.Name, "Changed");
            }
            catch (Exception ex)
            {
                writeLog(ex.Message, ex.InnerException.Message, "Error");
            }
        }

        private void OnCreated(object sender, FileSystemEventArgs e)
        {
            try
            {
                writeLog(e.FullPath, e.Name, "Created");

                if (e.Name.IndexOf(".tif", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    if (GetIdleFile(e.FullPath))
                    {
                        convertPDF(e.FullPath);
                        writeLog(e.FullPath, e.Name, "Converted");
                    }  
                }

                if (e.Name.IndexOf(".pdf", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    if (GetIdleFile(e.FullPath))
                    {
                        SendMail(e.FullPath, e.Name);
                        writeLog(e.FullPath, e.Name, "Mailed");
                    }
                }
            }

            catch (Exception ex)
            {
                writeLog(ex.Message, ex.InnerException.Message, "Error");
            }
        }

        private void OnRenamed(object source, RenamedEventArgs e)
        {
            try
            {
                writeLog(e.FullPath, e.Name, "Renamed");
            }
            catch (Exception ex)
            {
                writeLog(ex.Message, ex.InnerException.Message, "Error");
            }
        }

        protected override void OnStart(string[] args)
        {
            try
            {
                writeLog("", "", "Started");

                fileSystemWatcher.Path = WatchPath1;
            }
            catch (Exception ex)
            {
                writeLog(ex.Message, ex.InnerException.Message, "Error");
            }
        }

        protected override void OnStop()
        {
            try
            {
                writeLog("", "", "Stopped");
            }
            catch (Exception ex)
            {

                writeLog(ex.Message, ex.InnerException.Message, "Error");
            }
        }
    }
}
