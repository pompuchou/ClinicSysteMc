using ClinicSysteMc.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ClinicSysteMc.ViewModel.Converters
{
    internal class JYYconvert
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private readonly string _loadpath;
        private readonly DateTime _qdate;

        public JYYconvert(string loadpath)
        {
            this._loadpath = loadpath;
            this._qdate = DateTime.Now;
        }

        internal async Task Transform(IProgress<ProgressReportModel> progress)
        {
            string[] Lines = System.IO.File.ReadAllLines(_loadpath, System.Text.Encoding.Default);

            int totalN = Lines.Length;  // No titles

            int table_N = 500; // after testing 500 is better
            int total_div = totalN / table_N;
            int residual = totalN % table_N;
            ProgressReportModel report = new ProgressReportModel();

            log.Info($"  start async process.");
            List<Task<PTresult>> tasks = new List<Task<PTresult>>();

            // 將_data分拆成幾個小的Array
            for (int i = 0, idx = 0; i <= total_div; i++, idx += table_N)
            {
                string[] dummy;
                if (i < total_div)
                {
                    dummy = new string[table_N];
                    Array.Copy(Lines, idx, dummy, 0, table_N);
                }
                else
                {
                    dummy = new string[residual];
                    Array.Copy(Lines, idx, dummy, 0, residual);
                }
                tasks.Add(ImportJYY_async(dummy, progress, report));
            }

            PTresult[] result = await Task.WhenAll(tasks);

            int total_NewPT = (from p in result
                               select p.NewPT).Sum();
            int total_AllPT = (from p in result
                               select p.AllPT).Sum();
            log.Info($"  end async process.");

            string output = $"共處理{total_AllPT}筆教養院住民上傳資料, 其中{total_NewPT}筆新住民上傳資料.";
            log.Info(output);
            Logging.Record_admin("JYY add/change", output);
        }

        private async Task<PTresult> ImportJYY_async(string[] Lines, IProgress<ProgressReportModel> progress, ProgressReportModel report)
        {
            int totalN = Lines.Length;
            int add_N = 0;
            int all_N = 0;

            await Task.Run(() =>
            {
                log.Info($"    enter ImportJYY_async.");
                foreach (string Line in Lines)
                {
                    string[] lineStr = Line.Split(',');
                    string sYM = lineStr[0]; // 1  分年月,
                    string sCli = lineStr[1]; // 2  診所代碼,
                    string sIid = lineStr[2]; // 3  機構代碼,
                    string sUid = lineStr[3]; // 4  身分證字號,
                    string sBd = lineStr[4]; // 5  生日,
                    string sInsD = lineStr[5]; // 6  入院日期,
                    string sOutD = lineStr[6]; // 7  出院日期,
                    string sCname = lineStr[7]; // 8  姓名,

                    using (BSDataContext dc = new BSDataContext())
                    {
                        var q = from p in dc.tbl_upload
                                where (p.YM == sYM) && (p.uid == sUid)
                                select p;
                        if (q.Count() == 0)
                        {
                            try
                            {
                                tbl_upload newUP = new tbl_upload()
                                {
                                    YM = sYM,
                                    Cli = sCli,
                                    iid = sIid,
                                    uid = sUid,
                                    bd = sBd,
                                    InsD = sInsD,
                                    OutD = sOutD,
                                    cname = sCname,
                                    QDATE = _qdate
                                };
                                dc.tbl_upload.InsertOnSubmit(newUP);
                                dc.SubmitChanges();
                                add_N++;
                            }
                            catch (Exception ex)
                            {
                                Logging.Record_error(ex.Message);
                                log.Error(ex.Message);
                            }
                        }
                    }
                    all_N++;
                    report.PercentageComeplete = all_N * 100 / totalN;
                    progress.Report(report);
                }
                log.Info($"    exit ImportJYY_async.");
            });
            return new PTresult()
            {
                NewPT = add_N,
                AllPT = all_N
            };
        }
    }
}