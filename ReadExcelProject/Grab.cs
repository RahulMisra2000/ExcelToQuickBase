//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;

//namespace ReadExcelProject
//{
//    <<<<<<QuickBaseAPI Class:>>>>>>

//using System;
//using System.Data;
//using System.IO;
//using System.Net;
//using System.Text.RegularExpressions;
//using System.Xml;
//using System.Diagnostics;

//namespace sbo.AppTools
//    {
//        /// <summary>
//              /// Summary description for QuickBaseAPI.
//              /// </summary>
//        public class QuickBaseAPI
//        {
//            public QuickBaseAPI(string base_url, string appToken, bool writeToTrace)
//            {
//                if (!writeToTrace)
//                    writer = new StringWriter();
//                this.base_url = base_url;
//                this.appToken = appToken;
//            }

//            private string base_url, appToken, ticket;

//            private StringWriter writer;

//            private void Write(string value) { if (writer != null) writer.Write(value); else Trace.Write(value); }
//            private void WriteLine(string value) { if (writer != null) writer.WriteLine(value); else Trace.WriteLine(value); }
//            public string Result { get { return writer.ToString(); } }

//            private string CallApi(string url, string quickbase_action, string ticket, string appToken, string values)
//            {
//                HttpWebRequest request = (HttpWebRequest)HttpWebRequest.Create(url);
//                request.Method = "POST";
//                request.ContentType = "application/xml";
//                request.Headers.Add("QUICKBASE-ACTION", quickbase_action);
//                using (TextWriter writer = new StreamWriter(request.GetRequestStream()))
//                    writer.Write(@"
//<qdbapi>
//   <udata>optional data</udata>" + (ticket != null ? @"
//   <ticket>" + ticket + @"</ticket>" : "") + (appToken != null ? @"
//   <apptoken>" + appToken + @"</apptoken>" : "") + (values != null ? values : "") + @"
//</qdbapi>
//                              ");
//                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
//                using (TextReader reader = new StreamReader(response.GetResponseStream()))
//                    return reader.ReadToEnd();
//            }
//            private string CallApi(string url, string quickbase_action, string ticket, string appToken)
//            {
//                return CallApi(url, quickbase_action, ticket, appToken, null);
//            }
//            private string CallApiParams(string url, string quickbase_action, string ticket, string appToken,
//            params object[] name_value)
//            {
//                string values = "";
//                for (int i = 0; i + 1 < name_value.Length; i += 2)
//                    values += @"
//   <" + name_value[i].ToString() + ">" + name_value[i + 1].ToString() + "</" + name_value[i].ToString() + ">";
//                return CallApi(url, quickbase_action, ticket, appToken, values);
//            }

//            private bool ParseResult(string result, string tag, out string data)
//            {
//                Match match = Regex.Match(result, @"<" + tag + @">([^\s<>]+)</" + tag + @">", RegexOptions.IgnoreCase);
//                data = match.Success ? match.Groups[1].Value : "";
//                return !match.Success;
//            }

//            private bool CheckForError(string result)
//            {
//                string errcode;
//                ParseResult(result, "errcode", out errcode);
//                return (Convert.ToInt32(errcode) != 0);
//            }

//            private string FormatCsvData(object data)
//            {
//                return "\"" + data.ToString().Replace("\"", "\"\"") + "\"";
//            }

//            private string GetSchema(string url, string ticket, string appToken, out DataTable dt_xml)
//            {
//                dt_xml = null;

//                //get schema
//                string result = CallApiParams(url, "API_GetSchema", ticket, appToken);
//                if (CheckForError(result))
//                    return result;

//                XmlDocument xml = new XmlDocument();
//                xml.LoadXml(result);
//                XmlNodeList fields = xml.DocumentElement.SelectNodes("//fields/field");
//                dt_xml = new DataTable();
//                dt_xml.Rows.Add(dt_xml.NewRow());
//                foreach (XmlNode field in fields)
//                {
//                    string name = field.SelectSingleNode("label").InnerText;
//                    string id = field.Attributes["id"].Value;
//                    dt_xml.Columns.Add(name);
//                    dt_xml.Rows[0][name] = id;
//                }

//                return result;
//            }

//            private string LoadTable(string base_url, string dbid, string ticket, string appToken, string sql,
//            out int recordsLoaded, bool debug)
//            {
//                recordsLoaded = 0;

//                string result;
//                DataTable dt_xml = null;
//                if (!debug)
//                {
//                    //purge table
//                    result = CallApiParams(base_url + dbid, "API_PurgeRecords", ticket, appToken);
//                    if (CheckForError(result))
//                        return result;

//                    //get schema
//                    result = GetSchema(base_url + dbid, ticket, appToken, out dt_xml);
//                    if (CheckForError(result))
//                        return result;
//                }

//                //get data
//                DataTable dt;
//                Db_.LoadTable(Db_.CreateConnection(), null, sql, dt = new DataTable());

//                if (!debug)
//                {
//                    bool new_fields = false;

//                    //create fields
//                    foreach (DataColumn col in dt.Columns)
//                        if (!dt_xml.Columns.Contains(col.ColumnName))
//                        {
//                            result = CallApiParams(base_url + dbid, "API_AddField", ticket, appToken,
//                            "label", col.ColumnName,
//                            "type",
//                            col.DataType == typeof(DateTime) ? "datetime" :
//                            col.DataType == typeof(decimal) ||
//                            col.DataType == typeof(int) ||
//                            col.DataType == typeof(float) ||
//                            col.DataType == typeof(double) ? "float" :
//                            "text");
//                            if (CheckForError(result))
//                                return result;

//                            new_fields = true;
//                        }

//                    if (new_fields)
//                    {
//                        //get schema again (after creating new fields)
//                        result = GetSchema(base_url + dbid, ticket, appToken, out dt_xml);
//                        if (CheckForError(result))
//                            return result;
//                    }
//                }

//                //prepare data
//                string[] data = new string[dt.Rows.Count + 1];
//                string[] field = new string[dt.Columns.Count];
//                string[] field_id = new string[dt.Columns.Count];
//                for (int j = 0; j < dt.Columns.Count; j++)
//                {
//                    string col = dt.Columns[j].ColumnName;
//                    field[j] = FormatCsvData(col);
//                    field_id[j] = (dt_xml.Columns.Contains(col) ? dt_xml.DefaultView[0][col].ToString() : "0");
//                }
//                string clist = string.Join(".", field_id);
//                data[0] = string.Join(",", field);
//                Regex regex = new Regex(@"[;,\s]*<br>[;,\s]*", RegexOptions.IgnoreCase);
//                for (int i = 0; i < dt.Rows.Count; i++)
//                {
//                    field = new string[dt.Columns.Count];
//                    for (int j = 0; j < dt.Columns.Count; j++)
//                        field[j] = regex.Replace(FormatCsvData(dt.Rows[i][dt.Columns[j]])
//                        .Replace("\r", "").Replace("\n", ""), ", ");
//                    data[i + 1] = string.Join(",", field);
//                }

//                //load data
//                const int CHUNK_SIZE = 500;
//                for (int i = 0; i < data.Length / CHUNK_SIZE + (data.Length % CHUNK_SIZE > 0 ? 1 : 0); i++)
//                {
//                    int len = (i + 1) * CHUNK_SIZE > data.Length ? data.Length % CHUNK_SIZE : CHUNK_SIZE;
//                    int dim = i > 0 ? len + 1 : len;
//                    string[] tmp = new string[dim];
//                    Array.Copy(data, i * CHUNK_SIZE, tmp, i > 0 ? 1 : 0, len);
//                    if (i > 0)
//                        tmp[0] = data[0];
//                    string csv_values = @"
//      <![CDATA[
//" + string.Join("\r\n", tmp) + @"
//      ]]>
//   ";
//                    if (!debug)
//                    {
//                        result = CallApiParams(base_url + dbid, "API_ImportFromCSV", ticket, appToken,
//                        "records_csv", csv_values,
//                        "clist", clist,
//                        "skipfirst", 1);
//                        if (CheckForError(result))
//                            return csv_values + "\r\n\r\n" + result;
//                        else
//                            recordsLoaded += (i == 0 ? len - 1 : len);
//                    }
//                    else
//                    {
//                        using (StreamWriter writer = File.AppendText(@"c:\temp\CBE_to_QB.txt"))
//                            writer.Write(csv_values);
//                        recordsLoaded += (i == 0 ? len - 1 : len);
//                    }
//                }

//                return null;
//            }

//            private bool Output(string result, int recordsLoaded, string tableName)
//            {
//                if (recordsLoaded > 0)
//                    WriteLine(tableName + ":\t" + recordsLoaded + " record(s) loaded");
//                if (result != null)
//                {
//                    WriteLine(result);
//                    return true;
//                }
//                return false;
//            }

//            private bool failed = false;
//            public bool Failed { get { return failed; } }

//            public QuickBaseAPI Authenticate(string username, string password)
//            {
//                //authenticate
//                string result = CallApiParams(base_url + "main", "API_Authenticate", null, null,
//"username", username,
//"password", password,
//"hours", 24);
//                if (failed = ParseResult(result, "ticket", out ticket))
//                    WriteLine(result);
//                return this;
//            }

//            public QuickBaseAPI LoadTable(string table_dbid, string table_description, string sql, bool debug)
//            {
//                int recordsLoaded;
//                string result = LoadTable(base_url, table_dbid, ticket, appToken, sql, out recordsLoaded, debug);
//                failed = Output(result, recordsLoaded, table_description);
//                return this;
//            }

//            public QuickBaseAPI LoadTable_Applications() { return LoadTable_Applications("bd7gvuwnq", false); }
//            public QuickBaseAPI LoadTable_Applications(string table_dbid, bool debug)
//            {
//                return LoadTable(table_dbid, "Applications", @"
//SELECT 
//      r.LsdbeApplicationID, 
//      r.CompanyID, 
//      ApplicationType, 
//      LsdbeApplication_CreatedOn AS CreatedOn, 
//      LsdbeApplication_SubmittedOn AS SubmittedOn, 
//      CertCategories, 
//      r.Issued_LSDBE_Number AS CBE_Number, 
//      ApplicationStatus, (
//            SELECT TOP 1 ChangeDate
//            FROM  dbo.vwLsdbeApplication_StatusHistory_fix
//            WHERE LsdbeApplicationID = r.LsdbeApplicationID
//            ORDER BY ChangeDate DESC, LsdbeApplication_StatusHistoryID DESC 
//      ) AS AppStatusDate, 
//--    StatusDate, 
//      ApprovedAsOf AS ApprovalDate, 
//      ExpiredAsOf AS ExpirationDate, 
//      CompanyApprovedAsOf AS CompanyApprovalDate, 
//      CompanyExpiredAsOf AS CompanyExpirationDate, 
//      CompanyName AS BusinessName, 
//      FEIN_Number, 
//      CompanyStatus, 
//      RevenueAverage, 
//      BusinessStructure, 
//      BusinessServices, 
//      IndustryGroup, 
//      Address AS BusinessAddress, 
//      GIS_Quadrant, 
//      GIS_Ward, 
//      BusinessZip1 AS BusinessZip, 
//      ContactPhone, 
//      ContactEmail, 
//      SpecName AS Specialist, 
//      CompanyOwners, 
//      FullName AS PrincipalOwner, 
//      Gender AS PrincipalOwner_Gender, 
//      Race AS PrincipalOwner_Race, 
//      Citizenship AS PrincipalOwner_Citizenship, 
//      LGBT AS PrincipalOwner_IsLGBT, 
//      Veteran AS PrincipalOwner_IsVeteran, 
//--    lsdbeapplication_statusid, 
//      SITE_VISIT_COUNT AS SiteVisitCount, (
//            SELECT      The45BusinessDayBenchmarkCount 
//            FROM  dbo.tblLsdbeApplication 
//            WHERE LsdbeApplicationID = r.LsdbeApplicationID 
//      ) AS The45BusinessDayBenchmarkCount 
//FROM dbo.view_userReports r LEFT OUTER JOIN 
//      ( dbo.vwActiveApp v INNER JOIN 
//            dbo.vwActiveCompany c ON v.CompanyID = c.CompanyID AND 
//                  v.ApprovedAsOf >= c.CompanyApprovedAsOf AND v.ExpiredAsOf <= c.CompanyExpiredAsOf ) ON 
//            v.LsdbeApplicationID = r.LsdbeApplicationID ", debug);
//            }

//            public QuickBaseAPI LoadTable_NIGP_Codes() { return LoadTable_NIGP_Codes("bd6uhfyy5", false); }
//            public QuickBaseAPI LoadTable_NIGP_Codes(string table_dbid, bool debug)
//            {
//                return LoadTable(table_dbid, "NIGP Codes", @"
//SELECT     LO.LsdbeApplicationID, SR.Class + '-' + SR.Item + '-' + SR.Subitem AS NIGP, RTRIM(SR.Description) AS Description
//FROM         dbo.tblLsdbe_NIGPCode LO INNER JOIN
//                      dbo.NIGP7digit SR ON LO.NIGPCodeID = SR.ID ", debug);
//            }

//            public QuickBaseAPI LoadTable_Trade_Divisions() { return LoadTable_Trade_Divisions("bd6uhgmdh", false); }
//            public QuickBaseAPI LoadTable_Trade_Divisions(string table_dbid, bool debug)
//            {
//                return LoadTable(table_dbid, "Trade Divisions", @"
//SELECT     app.LsdbeApplicationID, t.TradeName AS Trade, td.TradeDivisionName AS TradeDivision, 
//                      tsd.SubDivisionCode + ' - ' + tsd.SubDivisionName AS SubDivision
//FROM         dbo.TradeDivision td INNER JOIN
//                      dbo.Trade t ON td.TradeID = t.TradeID INNER JOIN
//                      dbo.TradeSubDivision tsd ON td.TradeDivisionID = tsd.TradeDivisionID INNER JOIN
//                      dbo.tblTradeDivisions app ON LEFT(tsd.SubDivisionCode, 5) = app.SubDivisionCode ", debug);
//            }

//            public QuickBaseAPI LoadTable_Status_History() { return LoadTable_Status_History("bd7gvve5p", false); }
//            public QuickBaseAPI LoadTable_Status_History(string table_dbid, bool debug)
//            {
//                return LoadTable(table_dbid, "Status History", @"
//SELECT      LsdbeApplicationID, ApplicationStatus, ChangeDate, ChangeComment, FullName AS UserName, 
//      LsdbeApplication_StatusHistoryID AS StatusHistoryID 
//FROM  dbo.viewLsdbeApplicationHistory ", debug);
//            }

//            public QuickBaseAPI LoadTable_Status_Periods() { return LoadTable_Status_Periods("bd8t7ux3y", false); }
//            public QuickBaseAPI LoadTable_Status_Periods(string table_dbid, bool debug)
//            {
//                return LoadTable(table_dbid, "Status Periods", @"
//SELECT      LsdbeApplicationID, ApplicationStatus, StartDate, EndDate, 
//      LsdbeApplication_StatusHistoryID AS StatusHistoryID 
//FROM  dbo.vwLsdbeApplication_StatusPeriod h INNER JOIN
//      dbo.GetVwLsdbeApplication_StatusRef(0, 1) s ON h.LsdbeApplication_StatusID = s.LsdbeApplication_StatusID ", debug);
//            }

//            public QuickBaseAPI LoadTable_Specialist_History() { return LoadTable_Specialist_History("bd8uwy4vx", false); }
//            public QuickBaseAPI LoadTable_Specialist_History(string table_dbid, bool debug)
//            {
//                return LoadTable(table_dbid, "Specialist History", @"
//SELECT DISTINCT LsdbeApplicationID, LastName + ', ' + FirstName AS Specialist, SpecialistWasAssignedOn AS AssignedOn, 
//      DATEDIFF(second, '1/1/1970', SpecialistWasAssignedOn) AS AssignedOnID 
//FROM  dbo.tblLsdbeApplication_SpecialistHistory INNER JOIN 
//      dbo.tblSystemUser ON SystemUserID = SpecialistSystemUserID 
//ORDER BY AssignedOnID ", debug);
//            }

//            public QuickBaseAPI LoadTable_Companies() { return LoadTable_Companies("bd7gvudiw", false); }
//            public QuickBaseAPI LoadTable_Companies(string table_dbid, bool debug)
//            {
//                return LoadTable(table_dbid, "Companies", @"
//SELECT     CompanyID, CompanyName, Email AS UserEmail
//FROM        dbo.tblCompany c LEFT OUTER JOIN dbo.tblSystemUser u ON u.SystemUserID = c.SystemUserID ", debug);
//            }
//        }
//    }


////<<<<<< scheduled task code >>>>>>>

////                  //*******DMITRIY: Offload data to QuickBase
////                  try
////                  {
////                        Trace.WriteLine("");
////                        Trace.WriteLine("****************************************************");
////                        Trace.WriteLine("*** Initiating data load to QuickBase " + DateTime.Now);
////                        Trace.WriteLine("");
////                        //*****PUT LOGIC HERE

////                        QuickBaseAPI qb = new QuickBaseAPI("https://octo.quickbase.com/db/",
//// "dnedmsd8hjfwmcdysqtwbar8xsw", // the "DSLBD testbed" application token
////                              true);

////                        //authenticate
////                        if (qb.Authenticate(Config.QuickBaseUsername, Config.QuickBasePassword).Failed)
////                              completedWithExceptions = true;
////                        else
////                        {
////                              //load Applications table
////                              completedWithExceptions |= qb.LoadTable_Applications().Failed;
////                              //load NIGP Codes table
////                              //completedWithExceptions |= qb.LoadTable_NIGP_Codes().Failed;
////                              //load Trade Divisions table
////                              //completedWithExceptions |= qb.LoadTable_Trade_Divisions().Failed;
////                              //load Application Status History table
////                              completedWithExceptions |= qb.LoadTable_Status_History().Failed;
////                              //load Application Status Periods table
////                              completedWithExceptions |= qb.LoadTable_Status_Periods().Failed;
////                              //load Application Specialist History table
////                              completedWithExceptions |= qb.LoadTable_Specialist_History().Failed;
////                              //load Companies table
////                              completedWithExceptions |= qb.LoadTable_Companies().Failed;
////                        }

//////log
////Trace.WriteLine("");
////                        Trace.WriteLine("*** Ended loading data to QuickBase at "  + DateTime.Now);
////                        Trace.WriteLine("****************************************************");
                  
////                  }
////                  catch (Exception exRem)
////                  {
////                        Trace.WriteLine("An exception has occured while loading data to QuickBase:"
////                              + Environment.NewLine + AVK.Errors.Explain(exRem, Environment.NewLine));
////                        completedWithExceptions = true;
////                  }


////<<<<<<<<< the log file example >>>>>>>>>>>

////****************************************************
////*** Initiating data load to QuickBase 5/15/2012 2:00:46 AM

////Applications:      9264 record(s) loaded
////Status History:   38527 record(s) loaded
////Status Periods:  38527 record(s) loaded
////Specialist History:             13656 record(s) loaded
////Companies:        8310 record(s) loaded

////*** Ended loading data to QuickBase at 5/15/2012 2:02:57 AM
////****************************************************

//}
