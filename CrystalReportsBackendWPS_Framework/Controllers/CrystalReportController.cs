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
        private const string ConnString = "Data Source=192.168.10.100;Initial Catalog=WPS;User ID=sa;Password=Utama1234";
        private const string DbServer = "192.168.10.100";
        private const string DbName = "WPS";
        private const string DbUser = "sa";
        private const string DbPass = "Utama1234";

        private const string NetworkReportRoot = @"\\192.168.10.100\ReportForms\WPS";
        private const string LogDir = @"C:\Temp";

        // ==== Metadata models ====
        private class SubSource   // satu tabel dalam subreport
        {
            public string storedProcedure { get; set; }
            public string dataset { get; set; }
            public List<string> sqlParameters { get; set; } = new List<string>();
            public bool noFill { get; set; } = false; // << tambahkan ini
        }

        private class SubMeta     // kumpulan tabel (bisa 1 atau banyak)
        {
            // kompatibel lama (single):
            public string storedProcedure { get; set; }
            public string dataset { get; set; }
            public List<string> sqlParameters { get; set; } = new List<string>();

            // baru (multi):
            public List<SubSource> sources { get; set; } = new List<SubSource>();
        }

        private class ReportMeta
        {
            public string dataset { get; set; }
            public string outputFile { get; set; }
            public string rptFile { get; set; }
            public string storedProcedure { get; set; }
            public List<string> sqlParameters { get; set; } = new List<string>();
            public Dictionary<string, SubMeta> storedProcedureSubreports { get; set; }
                = new Dictionary<string, SubMeta>(StringComparer.OrdinalIgnoreCase);
        }

        // ==== Helpers ====
        private static object CoerceValue(string raw)
        {
            if (raw == null) return DBNull.Value;
            if (DateTime.TryParse(raw, out var d)) return d;
            return raw;
        }

        private static DataTable ExecToDataTable(string spName, IEnumerable<string> paramNames,
                                                 Func<string, string> resolve, string tableName)
        {
            var dt = new DataTable(string.IsNullOrWhiteSpace(tableName) ? "Data" : tableName);
            using (var con = new SqlConnection(ConnString))
            using (var cmd = new SqlCommand("dbo." + spName, con) { CommandType = CommandType.StoredProcedure })
            using (var adp = new SqlDataAdapter(cmd))
            {
                if (paramNames != null)
                {
                    foreach (var p in paramNames)
                    {
                        var raw = resolve(p);
                        cmd.Parameters.AddWithValue("@" + p, CoerceValue(raw) ?? DBNull.Value);
                    }
                }
                adp.Fill(dt);
            }
            return dt;
        }

        private static void ForceDbLogin(ReportDocument doc, string server, string db, string user, string pass)
        {
            // 1) Global logon
            doc.SetDatabaseLogon(user, pass, server, db);
            // 2) DataSourceConnections
            try { for (int i = 0; i < doc.DataSourceConnections.Count; i++) doc.DataSourceConnections[i].SetConnection(server, db, user, pass); } catch { }
            // 3) Apply to all tables
            var loi = new TableLogOnInfo { ConnectionInfo = new ConnectionInfo { ServerName = server, DatabaseName = db, UserID = user, Password = pass, IntegratedSecurity = false } };
            foreach (Table t in doc.Database.Tables) t.ApplyLogOnInfo(loi);
            foreach (ReportDocument sr in doc.Subreports)
            {
                sr.SetDatabaseLogon(user, pass, server, db);
                try { for (int i = 0; i < sr.DataSourceConnections.Count; i++) sr.DataSourceConnections[i].SetConnection(server, db, user, pass); } catch { }
                foreach (Table t in sr.Database.Tables) t.ApplyLogOnInfo(loi);
            }
        }

        private static void DumpTables(ReportDocument doc, string logFile, string tag)
        {
            File.AppendAllText(logFile, $"[TABLES {tag}] MAIN:\n");
            foreach (Table t in doc.Database.Tables) File.AppendAllText(logFile, $"  - {t.Name} | {t.Location}\n");
            foreach (ReportDocument sr in doc.Subreports)
            {
                File.AppendAllText(logFile, $"[TABLES {tag}] SUB: {sr.Name}\n");
                foreach (Table t in sr.Database.Tables) File.AppendAllText(logFile, $"  - {t.Name} | {t.Location}\n");
            }
        }

        [HttpGet]
        [Route("api/crystalreport/wps/export-pdf")]
        public IHttpActionResult ExportPDF(string reportName)
        {
            ReportDocument rptDoc = null;
            try
            {
                if (string.IsNullOrWhiteSpace(reportName))
                    return BadRequest("Parameter 'reportName' wajib diisi.");

                Directory.CreateDirectory(LogDir);
                var logFile = Path.Combine(LogDir, "crystal_debug_log.txt");

                // metadata
                var metadataFilePath = System.Web.Hosting.HostingEnvironment.MapPath($"~/Metadata/{reportName}.json");
                if (!File.Exists(metadataFilePath))
                    return InternalServerError(new FileNotFoundException("File metadata tidak ditemukan.", metadataFilePath));

                var meta = JsonConvert.DeserializeObject<ReportMeta>(File.ReadAllText(metadataFilePath));
                if (meta == null || string.IsNullOrWhiteSpace(meta.rptFile))
                    return InternalServerError(new InvalidDataException("Metadata tidak valid atau rptFile kosong."));

                var rptPath = Path.Combine(NetworkReportRoot, meta.rptFile);
                if (!File.Exists(rptPath))
                    return InternalServerError(new FileNotFoundException($"File .rpt tidak ditemukan di: {rptPath}", rptPath));

                // resolver param 2 arah
                var qp = Request.GetQueryNameValuePairs().ToDictionary(kv => kv.Key, kv => kv.Value, StringComparer.OrdinalIgnoreCase);
                string username = qp.ContainsKey("Username") ? qp["Username"] : "Unknown";
                string Resolver(string key)
                {
                    if (qp.TryGetValue(key, out var v) && !string.IsNullOrWhiteSpace(v)) return v;
                    if (key.Equals("TglAwal", StringComparison.OrdinalIgnoreCase) && qp.TryGetValue("StartDate", out v)) return v;
                    if (key.Equals("TglAkhir", StringComparison.OrdinalIgnoreCase) && qp.TryGetValue("EndDate", out v)) return v;
                    if (key.Equals("StartDate", StringComparison.OrdinalIgnoreCase) && qp.TryGetValue("TglAwal", out v)) return v;
                    if (key.Equals("EndDate", StringComparison.OrdinalIgnoreCase) && qp.TryGetValue("TglAkhir", out v)) return v;
                    if (key.Equals("Username", StringComparison.OrdinalIgnoreCase)) return username;
                    return null;
                }

                // main (opsional)
                DataTable dtMain = null;
                var hasMainSp = !string.IsNullOrWhiteSpace(meta.storedProcedure);
                if (hasMainSp)
                {
                    foreach (var p in meta.sqlParameters ?? Enumerable.Empty<string>())
                        if (Resolver(p) == null) return BadRequest($"Parameter '{p}' wajib dikirim.");

                    dtMain = ExecToDataTable(meta.storedProcedure, meta.sqlParameters, Resolver,
                                             string.IsNullOrWhiteSpace(meta.dataset) ? "MainData" : meta.dataset);

                    if (dtMain.Columns.Contains("InOut"))
                    {
                        dtMain.Columns["InOut"].ColumnName = "InOut_Old";
                        dtMain.Columns.Add("InOut", typeof(bool));
                        foreach (DataRow r in dtMain.Rows)
                            r["InOut"] = (r["InOut_Old"] != DBNull.Value && Convert.ToInt32(r["InOut_Old"]) == 1);
                        dtMain.Columns.Remove("InOut_Old");
                    }
                    if (dtMain.Rows.Count == 0) return BadRequest("Data tidak ditemukan.");
                }

                // load RPT
                rptDoc = new ReportDocument();
                rptDoc.Load(rptPath);
                DumpTables(rptDoc, logFile, "AFTER-LOAD");

                // paksa login (kalau ada objek DB Fields tersisa)
                ForceDbLogin(rptDoc, DbServer, DbName, DbUser, DbPass);

                if (hasMainSp) rptDoc.SetDataSource(dtMain);

                // validasi nama subreport
                var rptSubNames = rptDoc.Subreports.Cast<ReportDocument>().Select(s => s.Name)
                                   .ToHashSet(StringComparer.OrdinalIgnoreCase);
                File.AppendAllText(logFile, $"[{DateTime.Now}] REPORT={reportName} RPT={meta.rptFile}\n");
                File.AppendAllText(logFile, $"Subreports in RPT: {string.Join(", ", rptSubNames)}\n");

                // === supply data per subreport (bisa multi-table) ===
                if (meta.storedProcedureSubreports != null && meta.storedProcedureSubreports.Count > 0)
                {
                    foreach (var kv in meta.storedProcedureSubreports)
                    {
                        var subName = kv.Key;
                        var sMeta = kv.Value;

                        if (!rptSubNames.Contains(subName))
                        {
                            File.AppendAllText(logFile, $"[WARN] Subreport '{subName}' tidak ada di file RPT.\n");
                            continue;
                        }

                        // normalisasi ke "sources"
                        var sources = new List<SubSource>();
                        if (sMeta.sources != null && sMeta.sources.Count > 0)
                        {
                            sources.AddRange(sMeta.sources);
                        }
                        else if (!string.IsNullOrWhiteSpace(sMeta.storedProcedure))
                        {
                            sources.Add(new SubSource
                            {
                                storedProcedure = sMeta.storedProcedure,
                                dataset = string.IsNullOrWhiteSpace(sMeta.dataset) ? "SubData" : sMeta.dataset,
                                sqlParameters = sMeta.sqlParameters ?? new List<string>()
                            });
                        }
                        else
                        {
                            File.AppendAllText(logFile, $"[WARN] Subreport '{subName}' tidak punya sumber data di metadata.\n");
                            continue;
                        }

                        // Build DataSet berisi SEMUA tabel yang dibutuhkan subreport tsb
                        var dsSub = new DataSet(subName);
                        DataTable firstSchema = null;

                        foreach (var src in sources)
                        {
                            var tableName = string.IsNullOrWhiteSpace(src.dataset) ? "SubData" : src.dataset;

                            // SKIP eksekusi jika noFill atau SP kosong → supply DataTable kosong
                            if (src.noFill || string.IsNullOrWhiteSpace(src.storedProcedure))
                            {
                                var dtEmpty = new DataTable(tableName);
                                dsSub.Tables.Add(dtEmpty);
                                File.AppendAllText(logFile, $"[SUB {subName}] placeholder empty table '{tableName}' (no fill).\n");
                                continue;
                            }

                            try
                            {
                                var dt = ExecToDataTable(src.storedProcedure, src.sqlParameters, Resolver, tableName);
                                dsSub.Tables.Add(dt);
                                if (firstSchema == null) firstSchema = dt.Clone(); // opsional
                                File.AppendAllText(logFile, $"[SUB {subName}] filled {tableName} via {src.storedProcedure}\n");
                            }
                            catch (SqlException ex) when (ex.Number == 2812) // SP not found
                            {
                                var dtEmpty = firstSchema != null ? firstSchema.Clone() : new DataTable(tableName);
                                dtEmpty.TableName = tableName;
                                dsSub.Tables.Add(dtEmpty);
                                File.AppendAllText(logFile, $"[SUB {subName}] WARN: SP '{src.storedProcedure}' not found → supplying empty table '{tableName}'.\n");
                            }
                            catch (Exception exAny)
                            {
                                var dtEmpty = firstSchema != null ? firstSchema.Clone() : new DataTable(tableName);
                                dtEmpty.TableName = tableName;
                                dsSub.Tables.Add(dtEmpty);
                                File.AppendAllText(logFile, $"[SUB {subName}] WARN: failed to fill '{tableName}' via '{src.storedProcedure}' → {exAny.Message} → supplying empty table.\n");
                            }
                        }

                        // Set dataset ke subreport (bukan DataTable tunggal)
                        rptDoc.Subreports[subName].SetDataSource(dsSub);

                        // set parameter di subreport (kalau ada)
                        foreach (ParameterFieldDefinition pf in rptDoc.Subreports[subName].DataDefinition.ParameterFields)
                        {
                            var raw = Resolver(pf.Name);
                            if (raw == null) continue;
                            rptDoc.SetParameterValue(pf.Name, CoerceValue(raw), subName);
                            File.AppendAllText(logFile, $"[SUB SET PARAM] {subName}.{pf.Name} = {raw}\n");
                        }
                    }
                }

                // parameter induk
                foreach (ParameterFieldDefinition pf in rptDoc.DataDefinition.ParameterFields)
                {
                    var raw = Resolver(pf.Name);
                    if (raw == null && pf.Name.Equals("Username", StringComparison.OrdinalIgnoreCase)) raw = username;
                    if (raw != null) rptDoc.SetParameterValue(pf.Name, CoerceValue(raw));
                }

                DumpTables(rptDoc, logFile, "BEFORE-EXPORT");

                // export
                byte[] pdfBytes;
                using (var s = rptDoc.ExportToStream(ExportFormatType.PortableDocFormat))
                using (var ms = new MemoryStream())
                {
                    s.CopyTo(ms);
                    pdfBytes = ms.ToArray();
                }

                rptDoc.Close(); rptDoc.Dispose(); rptDoc = null;

                var http = new HttpResponseMessage(HttpStatusCode.OK) { Content = new ByteArrayContent(pdfBytes) };
                http.Content.Headers.ContentType = new MediaTypeHeaderValue("application/pdf");
                http.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("inline")
                {
                    FileName = string.IsNullOrWhiteSpace(meta.outputFile) ? (reportName + ".pdf") : meta.outputFile
                };
                return ResponseMessage(http);
            }
            catch (Exception ex)
            {
                Directory.CreateDirectory(LogDir);
                File.AppendAllText(Path.Combine(LogDir, "crystal_error_log.txt"), $"[{DateTime.Now}] {ex}\n");
                return InternalServerError(new Exception("Terjadi kesalahan saat memproses report. Lihat crystal_error_log.txt"));
            }
            finally
            {
                if (rptDoc != null) { try { rptDoc.Close(); } catch { } try { rptDoc.Dispose(); } catch { } }
            }
        }
    }
}
