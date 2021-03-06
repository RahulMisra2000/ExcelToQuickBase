﻿using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

namespace ReadExcelProject
{
    /// <summary>
    /// Requires EEPlus Nuget Package 
    /// Install this application in a folder where you have both read (Because we read Excel) and write permission (we are writing out log.txt)
    /// </summary>
    class Program
    {        

        // Order of precedence -- Values in instance variable - Values from Settings -- Values from command-line arguments
        private string _baseUrl = "https://octo.quickbase.com/db/";
        private string _username = "rahulmisra2000@gmail.com";
        private string _password = "";      
        private string _appToken = "bwy3haxdh8fsi7v6p6k7dyexycn";
        private string _authTicket;

        private string _fileName = "1.xlsx";
        private string _fqFileName;
        private string _tableId = "bkr5uxfek";

        private DataTable dt = null;

        private Stopwatch _stopWatch = new Stopwatch();            
        private List<string> _msgList = new List<string>();
        private const int FAIL = 9000;
        private const int PASS = 9001;
        
        private string _fromEmail = "ExcelToQuickBase@gmail.com"; // In gmail enable allow less secure apps
        private string _fromPwd = "";        
        private string _toEmail = "rahulmisra2000@gmail.com";

        private enum AppFeatures
        {
            None = 0,
            Fiddler = 1,
            SendEmail = 2,
            EventViewer = 4,
            ConsoleOutput = 8,
            WriteDTtoScreen = 16,
            HttpPostPayload = 32,
            WriteToLogFile = 64
        }

        private AppFeatures _enabledFeatures;
        
        static void Main(string[] args)
        {            
            new Program().Run(args);                
        }
        
        private void Run(string[] args)
        {
            BeginHouseKeeping(args);            
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
            ProcessWorkFlow();
            EndHouseKeeping();                  
        }

        private void WriteToConsole(List<string> ls, bool wait=false)
        {
            if ((_enabledFeatures & AppFeatures.ConsoleOutput) == AppFeatures.ConsoleOutput)
            {
                foreach (var s in ls) { Console.WriteLine(DateTime.Now + ": " + s); }

                if (wait)
                {
                    Console.WriteLine(DateTime.Now + " : Hit <Enter> to continue");
                    Console.ReadLine();
                }
            }
        }

        private void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            _msgList.Add(e.ExceptionObject.ToString());
            LogToEventViewer(FAIL);
            WriteToLogFile();
            WriteToConsole(new List<string> {"Severe Error: Check Event Viewer. Hit <Enter> to continue..."},true);            
            Environment.Exit(1);
        }

        private void LogToEventViewer(int id)
        {
            if ((_enabledFeatures & AppFeatures.EventViewer) == AppFeatures.EventViewer)
            {
                using (EventLog eventLog = new EventLog("Application"))
                {
                    eventLog.Source = "Application";
                    eventLog.WriteEntry(DateTime.Now + " : " + String.Join("\r\n",_msgList), 
                                        id == FAIL ? EventLogEntryType.Error : EventLogEntryType.Information, 
                                        id, 
                                        1);
                }
            }
        }

        private void WriteToLogFile()
        {
            if ((_enabledFeatures & AppFeatures.WriteToLogFile) == AppFeatures.WriteToLogFile)
            {
                try {
                    using (StreamWriter sw = File.AppendText("Log.txt"))
                    {
                        sw.WriteLine(DateTime.Now + " : " + String.Join("\r\n", _msgList));
                    }
                }
                catch (Exception e)
                {
                    _msgList.Add("ERROR: Cannot write log. Please give user Write permission to the application folder " + e.Message);
                }
            }
        }

        private void ProcessWorkFlow()
        {                                    
            if (!SuccessfullAuthenticated()) { _msgList.Add(string.Format("Failed Authentication for {0}",_username)); return; }
            if (!SuccessfullyHydratedDataTable()) { _msgList.Add(string.Format("Failed To Read Excel {0}",_fqFileName)); return; }
            WriteDataTableToScreen();
            if (!SuccessfullyWrittenToQuickBase()) { _msgList.Add(string.Format("Failed To Write to QuickBase TableID {0}",_tableId)); return; }
            // All OK
            _stopWatch.Stop();
            TimeSpan ts = _stopWatch.Elapsed;
            _msgList.Add(String.Format("Successfully wrote {0} Records In {1:00}:{2:00}:{3:00}.{4:00}",
                                            dt.Rows.Count, ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10));
        }

        private void BeginHouseKeeping(string[] args)
        {
            // sp: (S)pecial (P)rocess
            // This command-line argument is only used by the "Custom Action" called by the Setup Project and is invoked immediately after a successful installation
            if (!string.IsNullOrEmpty(args.SingleOrDefault(arg => arg.StartsWith("-sp:")))) { EncryptSensitiveSettings(args); Environment.Exit(1); }  // Fast Exit


            _enabledFeatures =  
                                  AppFeatures.SendEmail
                                | AppFeatures.EventViewer 
                                | AppFeatures.ConsoleOutput 
                             // | AppFeatures.WriteToLogFile
                                ;
            EncryptSensitiveSettings(args); // Insurance : just in case post-installation custom action did not encrypt the sensitive stuff in .config
            LoadFrom2SettingsFiles();
            ProcessCommandLineArguments(args);

            WriteToConsole(new List<string> {"Run started ..." });
            _stopWatch.Start();
        }

        /// <summary>
        /// Encrypt sensitive settings if they are not already encrypted
        /// </summary>
        private void EncryptSensitiveSettings(string[] args)
        {
            string tmp;
            string installationFolder = "";

            string baseAssemblyName = Assembly.GetExecutingAssembly().GetName().Name + ".exe";

            if (!string.IsNullOrEmpty(tmp = args.SingleOrDefault(arg => arg.StartsWith("-if:")))) { installationFolder = tmp.Replace("-if:", ""); }
            
            string assemblyLocation = Path.Combine(installationFolder, baseAssemblyName);
            
            try {
                Configuration config = ConfigurationManager.OpenExeConfiguration(assemblyLocation);
                var section = config.GetSectionGroup("applicationSettings").Sections["ReadExcelProject.Properties.Sensitive"];

                if (!section.SectionInformation.IsProtected)
                {
                    section.SectionInformation.ProtectSection("DataProtectionConfigurationProvider");

                }
                //else
                //{
                //    // For decrypting
                //    section.SectionInformation.UnprotectSection();
                //}
                config.Save();
            }
            catch (Exception e)
            {
                _msgList.Add(e.Message);                
            }

           
        }

        /// <summary>
        /// For Overriding hard coded settings in source code from settings file
        /// </summary>
        private void LoadFrom2SettingsFiles()
        {
            string tmp;

            // Sensitive.settings
            if (!string.IsNullOrEmpty(tmp = Properties.Sensitive.Default.Password))     { _password = tmp; }
            if (!string.IsNullOrEmpty(tmp = Properties.Sensitive.Default.FromPassword)) { _fromPwd = tmp; }

            // Settings.settings
            if (!string.IsNullOrEmpty(tmp = Properties.Settings.Default.AppToken))      { _appToken = tmp; }
            if (!string.IsNullOrEmpty(tmp = Properties.Settings.Default.TableId))       { _tableId = tmp; }
            if (!string.IsNullOrEmpty(tmp = Properties.Settings.Default.FileName))      { _fileName = tmp; }
            if (!string.IsNullOrEmpty(tmp = Properties.Settings.Default.UserName))      { _username = tmp; }            
            if (!string.IsNullOrEmpty(tmp = Properties.Settings.Default.BaseUrl))       { _baseUrl = tmp; }
            if (!string.IsNullOrEmpty(tmp = Properties.Settings.Default.FromEmail))     { _fromEmail = tmp; }            
            if (!string.IsNullOrEmpty(tmp = Properties.Settings.Default.ToEmail))       { _toEmail = tmp; }            
        }

        /// <summary>
        /// For Overriding settings and features from command-line arguments
        /// </summary>
        /// <param name="args"></param>
        private void ProcessCommandLineArguments(string[] args)
        {
            string tmp;
            // Override Settings
            if (!string.IsNullOrEmpty(tmp = args.SingleOrDefault(arg => arg.StartsWith("-b:")))) { _baseUrl = tmp.Replace("-b:", ""); }
            if (!string.IsNullOrEmpty(tmp = args.SingleOrDefault(arg => arg.StartsWith("-u:")))) { _username = tmp.Replace("-u:", ""); }
            if (!string.IsNullOrEmpty(tmp = args.SingleOrDefault(arg => arg.StartsWith("-p:")))) { _password = tmp.Replace("-p:", ""); }
            if (!string.IsNullOrEmpty(tmp = args.SingleOrDefault(arg => arg.StartsWith("-f:")))) { _fileName = tmp.Replace("-f:", ""); }            
            if (!string.IsNullOrEmpty(tmp = args.SingleOrDefault(arg => arg.StartsWith("-t:")))) { _tableId = tmp.Replace("-t:", ""); }
            if (!string.IsNullOrEmpty(tmp = args.SingleOrDefault(arg => arg.StartsWith("-e:")))) { _toEmail = tmp.Replace("-e:", ""); }
            _fqFileName = Path.GetDirectoryName(AppDomain.CurrentDomain.BaseDirectory) + "\\" + _fileName;

            // Disable Features
            if (!string.IsNullOrEmpty(tmp = args.SingleOrDefault(arg => arg.StartsWith("-x:c")))) { 
               if (_enabledFeatures.HasFlag(AppFeatures.ConsoleOutput))
                {
                    _enabledFeatures = _enabledFeatures & ~AppFeatures.ConsoleOutput;
                }
            }
            if (!string.IsNullOrEmpty(tmp = args.SingleOrDefault(arg => arg.StartsWith("-x:l"))))
            {  
                if (_enabledFeatures.HasFlag(AppFeatures.WriteToLogFile))
                {
                    _enabledFeatures = _enabledFeatures & ~AppFeatures.WriteToLogFile;
                }
            }
            if (!string.IsNullOrEmpty(tmp = args.SingleOrDefault(arg => arg.StartsWith("-x:m"))))
            {
                if (_enabledFeatures.HasFlag(AppFeatures.SendEmail))
                {
                    _enabledFeatures = _enabledFeatures & ~AppFeatures.SendEmail;
                }
            }
            if (!string.IsNullOrEmpty(tmp = args.SingleOrDefault(arg => arg.StartsWith("-x:e"))))
            {
                if (_enabledFeatures.HasFlag(AppFeatures.EventViewer))
                {
                    _enabledFeatures = _enabledFeatures & ~AppFeatures.EventViewer;
                }
            }
        }

        private void EndHouseKeeping()
        {            
            WriteToLogFile();
            SendEmail();            
            LogToEventViewer(_msgList.Where(m=>m.ToLower().Contains("failed")).Count()>0 ? FAIL : PASS);
            WriteToConsole(_msgList,true);
        }

        /// <summary>
        /// Remember to turn on "Access to less Secure apps in gmail
        /// </summary>
        private void SendEmail()
        {
            if ((_enabledFeatures & AppFeatures.SendEmail) == AppFeatures.SendEmail)
            {
                var client = new SmtpClient("smtp.gmail.com", 587)
                {
                    Credentials = new NetworkCredential(_fromEmail, _fromPwd),
                    EnableSsl = true
                };
                client.Send(_fromEmail, _toEmail, "Notification: Excel to QuickBase", DateTime.Now + " : " + String.Join("\r\n",_msgList));
            }
        }

        private bool SuccessfullAuthenticated()
        {
            string result = CallApiParams(_baseUrl + "main", "API_Authenticate", null, null, "username", _username, "password", _password, "hours", 24);
            return !ParseResult(result, "ticket", out _authTicket);            
        }

        /// <summary>
        /// The order and number of excel columns must match the order and number of fields in the quickbase table (ignore the 6 built-in ones)
        /// If not then, you will need to use the clist
        /// https://www.quickbase.com/api-guide/index.html#importfromcsv.html%3FTocPath%3DQuickBase%2520API%2520Call%2520Reference%7C_____51
        /// </summary>
        /// <returns></returns>
        private bool SuccessfullyWrittenToQuickBase()
        {
            StringBuilder sb = new StringBuilder();
            foreach (DataRow r in dt.Rows)
            {
                sb.Append(String.Join(",", r.ItemArray) + "\r\n");
            }
            string csv_values = @"<![CDATA[" +
                                    sb.ToString() +
                                    @"]]>";
            if ((_enabledFeatures & AppFeatures.HttpPostPayload) == AppFeatures.HttpPostPayload) { WriteToConsole( new List<string> { csv_values }); }
            
            //string clist = "6.7.8.9.10.11.12.13.14.15.16.17.18.19.20.21.22.23.24.25.26.27.28.29.30.31.32.33.34.35.36.37";

            string result = CallApiParams(_baseUrl + _tableId, "API_ImportFromCSV", _authTicket, _appToken,
                        "records_csv", csv_values,
              //         "clist", clist,
                        "skipfirst", 0);

            string errcode;
            ParseResult(result, "errcode", out errcode);
            return (Convert.ToInt32(errcode) == 0);
        }
        
        private string CallApiParams(string url, string quickbase_action, string ticket, string appToken, params object[] name_value)
        {
            string values = "";
            for (int i = 0; i + 1 < name_value.Length; i += 2)
            {
                values += @"<" + name_value[i].ToString() + ">" + name_value[i + 1].ToString() + "</" + name_value[i].ToString() + ">";
            }
            return CallApi(url, quickbase_action, ticket, appToken, values);
        }

        private string CallApi(string url, string quickbase_action, string ticket, string appToken, string values)
        {
            WriteToConsole(new List<string> { string.Format("Step - {0} @ {1}", quickbase_action, url) });
            HttpWebRequest request = (HttpWebRequest)HttpWebRequest.Create(url);
            request.Method = "POST";
            request.ContentType = "application/xml";
            request.Headers.Add("QUICKBASE-ACTION", quickbase_action);            
            if ((_enabledFeatures & AppFeatures.Fiddler) == AppFeatures.Fiddler) { request.Proxy = new WebProxy("127.0.0.1", 8888); }
            using (TextWriter writer = new StreamWriter(request.GetRequestStream()))
            {
                string temp = @"<qdbapi><udata>optional data</udata>" +
                                       (ticket != null ? @"<ticket>" + ticket + @"</ticket>" : "") +
                                       (appToken != null ? @"<apptoken>" + appToken + @"</apptoken>" : "") +
                                       (values != null ? values : "") +
                                 @"</qdbapi>";

                writer.Write(temp);
            }
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            using (TextReader reader = new StreamReader(response.GetResponseStream()))
                return reader.ReadToEnd();
        }

        private bool ParseResult(string result, string tag, out string data)
        {
            Match match = Regex.Match(result, @"<" + tag + @">([^\s<>]+)</" + tag + @">", RegexOptions.IgnoreCase);
            data = match.Success ? match.Groups[1].Value : "";
            return !match.Success;
        }

        private void WriteDataTableToScreen()
        {
            if ((_enabledFeatures & AppFeatures.WriteDTtoScreen) == AppFeatures.WriteDTtoScreen)
            {
                StringBuilder sb = new StringBuilder();
                foreach (DataRow r in dt.Rows)
                {
                    sb.Append(String.Join(",", r.ItemArray) + "\r\n");
                }
                string csv_values = @"<![CDATA[" +
                                        sb.ToString() +
                                        @"]]>";
                WriteToConsole(new List<string> { csv_values });                
            }
        }

        private bool SuccessfullyHydratedDataTable()
        {
            try
            {
                using (ExcelPackage pck = new ExcelPackage())
                {
                    using (var stream = File.OpenRead(_fqFileName))
                    {
                        pck.Load(stream);
                    }

                    ExcelWorksheet ws = pck.Workbook.Worksheets.First();
                    dt = WorksheetToDataTable(ws);
                }                
            }
            catch (Exception ex)
            {
                _msgList.Add(ex.Message);
                return false;
            }
            return true;
        }

        private DataTable WorksheetToDataTable(ExcelWorksheet ws, bool hasHeader = true)
        {
            DataTable dt = new DataTable(ws.Name);
            int totalCols = ws.Dimension.End.Column;
            int totalRows = ws.Dimension.End.Row;
            int startRow = hasHeader ? 2 : 1;
            ExcelRange wsRow;
            DataRow dr;
            int unq = 0;
            foreach (var firstRowCell in ws.Cells[1, 1, 1, totalCols])
            {
              dt.Columns.Add(hasHeader ? firstRowCell.Text + unq++ : string.Format("Column {0}", firstRowCell.Start.Column));
            }

            for (int rowNum = startRow; rowNum <= totalRows; rowNum++)
            {
                if (rowNum % 100 == 0) {
                    WriteToConsole(new List<string> { string.Format("Step - {0} of {1} Records From Excel Processed", rowNum, totalRows) });
                }
                wsRow = ws.Cells[rowNum, 1, rowNum, totalCols];
                dr = dt.NewRow();
                foreach (var cell in wsRow)
                {
                    dr[cell.Start.Column - 1] = Regex.Matches(cell.Text, @"[A-Za-z,]").Count > 0 
                        ? "\"" + cell.Text + "\"" 
                        : cell.Text;                    
                }

                dt.Rows.Add(dr);
            }
            WriteToConsole(new List<string> { string.Format("Step - {0} of {1} Records From Excel Processed", totalRows, totalRows) });
            return dt;
        }
    }
}

