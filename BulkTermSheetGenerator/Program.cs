using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using log4net;
using System.IO;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.Data.SqlClient;
using RestSharp;
using System.Net;

namespace BulkTermSheetGenerator
{
    class Program
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public static string folderPath = ConfigurationManager.AppSettings.Get("filePath");
        public static string outputPath = ConfigurationManager.AppSettings.Get("outputPath");
        static void Main(string[] args)
        {
            List<string> fileList = new List<string>();
            Program app = new Program();
            fileList = app.GetAllFiles(folderPath, "xlsx");

            DataTable quoteData = app.readDataFile(fileList[0].ToString());
            var quoteID="";
            var approvedScenarioName = "";
            string termsheetStatus = "";
            SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["CRMConnectionString"].ToString());
            foreach (DataRow dtRow in quoteData.Rows)
            {
                quoteID = dtRow["Quote ID"].ToString().Trim();
                string selectScenario = string.Format("select pt_name,statuscodename from FilteredPT_designwinquotepart where statecode=0 and statuscode in (3,10) and amd_quoteid = '{0}';", quoteID);
                DataTable Approvedscenarios = app.GetDataFromDB(con, selectScenario);
                foreach (DataRow scen in Approvedscenarios.Rows)
                {
                    string scenarioName = scen["pt_name"].ToString().Trim();
                    string status = scen["statuscodename"].ToString().Trim();
                    if (quoteID != "" && quoteID.StartsWith("Q") && scenarioName != "")
                    {
                        var client = new RestClient(String.Format("http://{0}:81/api/termsheets/?quoteid={1}&scenarioname={2}", "atlcrmbeprdv01", quoteID, scenarioName));
                        var request = new RestRequest(Method.GET);
                        // var r2 = new RestRequest(Method.POST);
                        request.AddHeader("postman-token", "fff26f08-6c83-9d03-c45e-afa4a56eaa6e");
                        request.AddHeader("cache-control", "no-cache");
                        request.AddHeader("authorization", "Basic Y3JtYWRtaW46QHRobG9uMDk=");
                        //IRestResponse TermsheetResponse = client.Execute(request);
                        //var fileBytes = client.DownloadData(new RestRequest("#", Method.GET));
                        var fileBytes = client.DownloadData(request);
                        
                        File.WriteAllBytes(Path.Combine(outputPath, quoteID + "_" + scenarioName+"_"+status+".pdf"), fileBytes);
                    }

                }
                

            }


        }
        public List<String> GetAllFiles(String directory, string type)
        {
            if (Directory.Exists(directory))
            {
                return Directory.GetFiles(directory, "*." + type.Trim(), SearchOption.AllDirectories).ToList();
            }
            return null;
        }
        private DataTable readDataFile(string filePath)
        {

            DataTable ExcelTable = new DataTable();
            List<string[]> resultsList = null;

            string strConn;
            if (filePath.Substring(filePath.LastIndexOf('.')).ToLower() == ".xlsx")
                strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties=\"Excel 12.0;HDR=YES;IMEX=1\"";
            else
                strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";Extended Properties=\"Excel 8.0;HDR=YES;IMEX=1\"";
            OleDbConnection conn = new OleDbConnection(strConn);
            conn.Open();
            DataTable schemaTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
            foreach (DataRow schemaRow in schemaTable.Rows)
            {
                string sheetName = schemaRow["TABLE_NAME"].ToString();
                if (!sheetName.EndsWith("_"))
                {
                    //string query = "SELECT  * FROM [" + sheetName + "] WHERE TaskNumber = " + "\"" + parameter + "\"";
                    string query = "SELECT  * FROM [" + sheetName + "]";
                    OleDbDataAdapter oleAdap = new OleDbDataAdapter(query, conn);
                    ExcelTable.Locale = CultureInfo.CurrentCulture;
                    oleAdap.Fill(ExcelTable);
                    oleAdap.Dispose();
                }
            }
            conn.Close();
            return ExcelTable;
        }

        private DataTable GetDataFromDB(SqlConnection con, string sqlQuery)
        {
            
            SqlDataAdapter da;
            DataTable dt = null;
            try
            {

                string qry = sqlQuery;
                con.Open();
                da = new SqlDataAdapter(qry, con);
                da.SelectCommand.CommandTimeout = 180;
                dt = new DataTable();
                da.Fill(dt);
                con.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                
            }
            finally
            {
                con.Close();
            }
            return dt;
        }
    }
}
