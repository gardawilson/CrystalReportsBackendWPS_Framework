using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web.Http;

namespace CrystalReportsBackendWPS_Framework.Controllers
{
    public class CrystalReportController : ApiController
    {
        [HttpGet]
        [Route("api/crystalreport/wps/export-pdf")]
        public IHttpActionResult ExportPDF(string reportName)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(reportName))
                    return BadRequest("Parameter 'reportName' wajib diisi.");

                string networkPath = @"\\192.168.10.100\ReportForms\WPS";
                string metadataFilePath = System.Web.Hosting.HostingEnvironment.MapPath($"~/Metadata/{reportName}.json");

                if (!File.Exists(metadataFilePath))
                    return InternalServerError(new FileNotFoundException("File metadata.json tidak ditemukan di folder Metadata."));

                dynamic meta = JsonConvert.DeserializeObject(File.ReadAllText(metadataFilePath));
                string rptFile = meta.rptFile;
                string outputFile = meta.outputFile;
                string datasetName = meta.dataset;
                string storedProcedure = meta.storedProcedure;
                List<string> sqlParams = ((IEnumerable<dynamic>)meta.sqlParameters).Select(p => (string)p).ToList();

                var queryParams = Request.GetQueryNameValuePairs().ToDictionary(kv => kv.Key, kv => kv.Value, StringComparer.OrdinalIgnoreCase);

                foreach (var param in sqlParams)
                {
                    if (!queryParams.ContainsKey(param))
                        return BadRequest($"Parameter '{param}' wajib dikirim.");
                }

                string username = queryParams.ContainsKey("Username") ? queryParams["Username"] : "Unknown";

                string rptPath = Path.Combine(networkPath, rptFile);
                if (!File.Exists(rptPath))
                    return InternalServerError(new FileNotFoundException($"File .rpt tidak ditemukan di: {rptPath}"));

                string logDir = @"C:\Temp";
                Directory.CreateDirectory(logDir);
                string logFile = Path.Combine(logDir, "crystal_debug_log.txt");

                File.AppendAllText(logFile, $"\n[{DateTime.Now}] Menjalankan report: {reportName} ({rptFile})\n");
                File.AppendAllText(logFile, $"Stored Procedure: {storedProcedure}\n");

                DataTable dt = new DataTable(datasetName);
                using (SqlConnection conn = new SqlConnection("Data Source=192.168.10.100;Initial Catalog=WPS;User ID=sa;Password=Utama1234"))
                {
                    using (SqlCommand cmd = new SqlCommand($"dbo.{storedProcedure}", conn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;

                        foreach (var param in sqlParams)
                        {
                            string value = queryParams[param];
                            object finalValue = DateTime.TryParse(value, out DateTime dateVal) ? (object)dateVal : value;
                            cmd.Parameters.AddWithValue("@" + param, string.IsNullOrEmpty(value) ? DBNull.Value : finalValue);
                            File.AppendAllText(logFile, $" - @{param} = {finalValue}\n");
                        }

                        SqlDataAdapter da = new SqlDataAdapter(cmd);
                        da.Fill(dt);

                        // Konversi kolom InOut jadi Boolean agar sesuai dengan report
                        if (dt.Columns.Contains("InOut"))
                        {
                            dt.Columns["InOut"].ColumnName = "InOut_Old"; // hindari konflik
                            dt.Columns.Add("InOut", typeof(bool));
                            foreach (DataRow row in dt.Rows)
                            {
                                int val = row["InOut_Old"] != DBNull.Value ? Convert.ToInt32(row["InOut_Old"]) : 0;
                                row["InOut"] = (val == 1);
                            }
                            dt.Columns.Remove("InOut_Old");
                        }


                    }
                }

                if (dt.Rows.Count == 0)
                    return BadRequest("Data tidak ditemukan.");

                ReportDocument rptDoc = new ReportDocument();
                rptDoc.Load(rptPath);
                rptDoc.SetDataSource(dt);

                // === JALANKAN SUBREPORT JIKA ADA ===
                if (meta.storedProcedureSubreports != null)
                {
                    foreach (var sub in meta.storedProcedureSubreports)
                    {
                        string subreportName = sub.Name; // Nama subreport dalam RPT (bisa .rpt atau tanpa)
                        string subSP = sub.Value.storedProcedure;
                        string subDataset = sub.Value.dataset;
                        List<string> subParams = ((IEnumerable<dynamic>)sub.Value.sqlParameters).Select(p => (string)p).ToList();

                        DataTable dtSub = new DataTable(subDataset);

                        using (SqlConnection connSub = new SqlConnection("Data Source=192.168.10.100;Initial Catalog=WPS;User ID=sa;Password=Utama1234"))
                        {
                            using (SqlCommand cmdSub = new SqlCommand(subSP, connSub))
                            {
                                cmdSub.CommandType = CommandType.StoredProcedure;
                                foreach (var p in subParams)
                                {
                                    string val = queryParams.ContainsKey(p) ? queryParams[p] : null;
                                    object finalVal = DateTime.TryParse(val, out var dtVal) ? (object)dtVal : val;
                                    cmdSub.Parameters.AddWithValue("@" + p, string.IsNullOrEmpty(val) ? DBNull.Value : finalVal);
                                    File.AppendAllText(logFile, $"[SUBREPORT PARAM] {subreportName} -> @{p} = {finalVal}\n");
                                }

                                SqlDataAdapter daSub = new SqlDataAdapter(cmdSub);
                                daSub.Fill(dtSub);
                            }
                        }

                        try
                        {
                            rptDoc.Subreports[subreportName].SetDataSource(dtSub);
                        }
                        catch (Exception exSub)
                        {
                            File.AppendAllText(logFile, $"[SUBREPORT ERROR] {subreportName} ❌ {exSub.Message}\n");
                        }
                    }
                }


                // ==== SET DAN LOG PARAMETER CRYSTAL REPORT ====
                for (int i = 0; i < rptDoc.ParameterFields.Count; i++)
                {
                    var paramField = rptDoc.ParameterFields[i];
                    string paramName = paramField.Name;

                    string paramKey = queryParams.Keys.FirstOrDefault(k =>
                        k.Equals(paramName, StringComparison.OrdinalIgnoreCase) ||
                        ("Txt" + k).Equals(paramName, StringComparison.OrdinalIgnoreCase)
                    );

                    if (!string.IsNullOrEmpty(paramKey))
                    {
                        string rawValue = queryParams[paramKey];
                        object value = (paramField.ParameterValueType == ParameterValueKind.DateParameter ||
                                        paramField.ParameterValueType == ParameterValueKind.DateTimeParameter)
                            ? (DateTime.TryParse(rawValue, out var dtVal) ? (object)dtVal : rawValue)
                            : rawValue;

                        rptDoc.SetParameterValue(paramName, value);
                        File.AppendAllText(logFile, $"[SET PARAM] {paramName} ← dari query '{paramKey}' = {value}\n");
                    }
                    else if (paramName.Equals("Username", StringComparison.OrdinalIgnoreCase))
                    {
                        rptDoc.SetParameterValue("Username", username);
                        File.AppendAllText(logFile, $"[SET PARAM] {paramName} = {username}\n");
                    }
                    else
                    {
                        File.AppendAllText(logFile, $"[MISSING PARAM] {paramName} ❌ Tidak ditemukan di query string\n");
                    }
                }

                Stream pdfStream = rptDoc.ExportToStream(ExportFormatType.PortableDocFormat);
                pdfStream.Seek(0, SeekOrigin.Begin);

                HttpResponseMessage result = new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new StreamContent(pdfStream)
                };
                result.Content.Headers.ContentType = new MediaTypeHeaderValue("application/pdf");
                result.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("inline") { FileName = outputFile };

                return ResponseMessage(result);
            }
            catch (Exception ex)
            {
                string errLog = @"C:\Temp\crystal_error_log.txt";
                File.AppendAllText(errLog, $"[{DateTime.Now}] {ex}\n");
                return InternalServerError(new Exception("Terjadi kesalahan saat memproses report. Lihat crystal_error_log.txt"));
            }
        }
    }
}
