using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Net;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace FileDownloader
{
    public partial class Form1 : Form
    {
        private WebClient wc = new WebClient();

        private List<string> RunSheetFailures = new List<string>();

        private bool AutoRunFromCommandLineParameter = false;

        private string[] cmdArgs;


        EventWaitHandle waitHandle = new EventWaitHandle(true, EventResetMode.AutoReset, "RBA_DOWNLOADER_SHARED_BY_ALL_PROCESSES");
        public Form1(string[] args)
        {
            InitializeComponent();

            if (args?.Length > 0)
            {
                //Kick off the job as its being run from a Scheduled Task
                AutoRunFromCommandLineParameter = true;
                cmdArgs = args;
                textBox1.Text = args[0];
                textBox2.Text = args[1];
                LogActivity("Auto: " + DateTime.Now.ToString() + "\t\tFileUrl: " + args[0] + "\t\tDirectory: " + args[1]);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {            
            try
            {
                if (AutoRunFromCommandLineParameter)
                {
                    if (cmdArgs.Length < 1)
                    {
                        MessageBox.Show("Please provide 2 command line arguments, eg: FileDownloader.exe UrlOfFileToDownload DownloadDirectory");
                    }
                    else
                    {
                        DownloadFile(cmdArgs[0], cmdArgs[1]);                        
                        ReportFailures(Application.StartupPath, ConfigurationManager.AppSettings["PeopleToEmail"].ToString());
                        Environment.Exit(0);
                    }
                }
            }
            catch (Exception ex)
            {
                LogActivity("Exception " + DateTime.Now.ToString() + "\t\tError Message: " + ex.Message);
                Email email = new Email();
                email.SendEmail(ConfigurationManager.AppSettings.Get("SupportEmail").ToString(), "", "", "Exception in File Downloader!!", "Message: " + ex.Message + "<BR>" + ex.StackTrace, null);
            }
        }

        private void LogActivity(string logMessage)
        {
            //Using a EventWaitHandle allows multiple processes to use this file, one at a time
            if (waitHandle.WaitOne())
            {
                try
                {
                    using (FileStream fs = File.Open("Log.txt", FileMode.Append, FileAccess.Write, FileShare.ReadWrite))
                    {
                        byte[] bytes = Encoding.ASCII.GetBytes(logMessage + Environment.NewLine);
                        fs.Write(bytes, 0, bytes.Length);
                        fs.Close();
                    }
                }
                finally
                {
                    waitHandle.Set();
                }
            }
        }
        
        private void BtnGo_Click(object sender, EventArgs e)
        {
            MessageBox.Show(DownloadFile(textBox1.Text, textBox2.Text));
        }

        private string DownloadFile(string url, string directory)
        {
            url = url.Replace("\"", "");
            directory = directory.Replace("\"", "");
            byte[] byteArray = null;
            bool success = false;
            try
            {
                byteArray = TestForHTTP404(url);
                
                if (byteArray != null)
                {                    
                    success = Download(url, directory, byteArray);
                }
            }
            catch (Exception ex)
            {
                AddReportFailure("FF0000", "Error in processing: " + ex.Message, url, directory, "");
                return ex.Message;
            }

            return success ? "Successfully Downloaded" : "Problem downloading file";
        }

        private byte[] TestForHTTP404(string url)
        {
            try
            {
                return wc.DownloadData(url);
            }
            catch (WebException webEx)
            {
                //A 404, keep trying
            }
            catch (Exception ex)
            {
                //Something else unknown, meh?
            }
            return null;
        }

        private bool Download(string url, string directory, byte[] byteArray)
        {
            string fileName = string.Empty;
            try
            {
                if (byteArray == null) byteArray = wc.DownloadData(url);
                
                ////Try to extract the filename from the Content-Disposition header
                //if (!String.IsNullOrEmpty(wc.ResponseHeaders["Content-Disposition"]))
                //{
                //    fileName = wc.ResponseHeaders["Content-Disposition"].Substring(wc.ResponseHeaders["Content-Disposition"].IndexOf("filename=") + 10).Replace("\"", "");
                //}

                //Resort to naming the file based on the KeyWords found in the link
                if (string.IsNullOrEmpty(fileName))
                {
                    fileName = Path.GetFileName(url);
                }
                                
                string destFileName = Path.Combine(directory, fileName);
                if (!Directory.Exists(directory)) Directory.CreateDirectory(directory);
                File.WriteAllBytes(destFileName, byteArray);
                
                AddReportFailure("00AA00", "Updated  Successfully!", fileName, directory, url);
                fileName = string.Empty;
                return true;
            }
            catch (WebException webEx)
            {
                AddReportFailure("0000FF", "Exception: " + ((HttpWebResponse)webEx.Response).StatusCode, fileName, directory, url);
                return false;
            }
            catch (Exception Ex)
            {
                AddReportFailure("FF0000", "Error: " + Ex.Message, fileName, directory, url);
                return false;
            }            
        }

        private void AddReportFailure(string colour, string col1, string col2, string col3, string col4)
        {
            RunSheetFailures.Add("<tr><td style=\"color:#" + colour + "\">" + col1 + "</td><td>" + col2 + " </td><td>" + col3 + " </td><td>" + col4.Replace(" ", "%20") + "</td></tr>");
        }

        internal void ReportFailures(string startupPath, string emailTo)
        {
            if (RunSheetFailures.Count > 0)
            {
                StringBuilder sb = new StringBuilder();
                sb.Append("Running from: ");
                sb.Append(startupPath);
                sb.Append("<BR><BR>");
                sb.AppendLine("File Downloader Report:");
                sb.Append("<Table>");

                sb.Append("<tr><td width=100><b>Status</b></td><td width=100><b>File Name</b></td><td width=100><b>Directory</b></td><td width=300><b>URL</b></td></tr>");

                foreach (string line in RunSheetFailures)
                {
                    sb.AppendLine(line);
                }

                sb.Append("</Table>");

                Email email = new Email();
                email.SendEmail(emailTo, "", "", "File Downloader Report", sb.ToString(), null);
            }
        }

    }
}
