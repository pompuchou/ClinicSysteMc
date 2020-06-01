using ClinicSysteMc.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ClinicSysteMc.ViewModel.Converters
{
    internal class HOSPconvert
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private readonly string _loadpath;
        private readonly DateTime _qdate;

        public HOSPconvert(string loadpath)
        {
            this._loadpath = loadpath;
            this._qdate = DateTime.Now;
        }

        internal async Task Transform(IProgress<ProgressReportModel> progress)
        {
            string[] Lines = System.IO.File.ReadAllLines(_loadpath, System.Text.Encoding.Default);

            int totalN = Lines.Length - 1;  // -1 because line 1 is titles, so I should begin with 1 to total_N + 1

            int table_N = 5000; // after testing 5000 is better
            int total_div = totalN / table_N;
            int residual = totalN % table_N;
            ProgressReportModel report = new ProgressReportModel();

            log.Info($"  start async process.");
            List<Task<PTresult>> tasks = new List<Task<PTresult>>();

            // 將_data分拆成幾個小的Array
            for (int i = 0, idx = 1; i <= total_div; i++, idx += table_N)
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
                tasks.Add(ImportHOSP_async(dummy, progress, report));
            }

            PTresult[] result = await Task.WhenAll(tasks);

            int total_NewPT = (from p in result
                               select p.NewPT).Sum();
            int total_ChangePT = (from p in result
                                  select p.ChangePT).Sum();
            int total_AllPT = (from p in result
                               select p.AllPT).Sum();
            log.Info($"  end async process.");

            string output = $"共處理{total_AllPT}筆特約機構資料, 其中{total_NewPT}筆新特約機構, 修改{total_ChangePT}筆特約機構資料.";
            log.Info(output);
            Logging.Record_admin("HOSP add/change", output);
        }

        private async Task<PTresult> ImportHOSP_async(string[] Lines, IProgress<ProgressReportModel> progress, ProgressReportModel report)
        {
            int totalN = Lines.Length;
            int add_N = 0;
            int change_N = 0;
            int all_N = 0;

            await Task.Run(() =>
            {
                log.Info($"    enter ImportHOSP_async.");
                foreach (string Line in Lines)
                {
                    string[] lineStr = Line.Split(',');
                    string sDiv = lineStr[0].Trim('\"'); // 1  分區別,
                    string sCod = lineStr[1].Trim('\"'); // 2  醫事機構代碼,
                    string sNam = lineStr[2].Trim('\"'); // 3  醫事機構名稱,
                    string sAdr = lineStr[3].Trim('\"'); // 4  機構地址,
                    string sLoc = lineStr[4].Trim('\"'); // 5  電話區域號碼,
                    string sTel = lineStr[5].Trim('\"'); // 6  電話號碼,
                    string sCls = lineStr[6].Trim('\"'); // 7  特約類別,
                    string sFor = lineStr[7].Trim('\"'); // 8  型態別,
                    string sTyp = lineStr[8].Trim('\"'); // 9  醫事機構種類,
                    string sDat = lineStr[9].Trim('\"'); // 10 終止合約或歇業日期,
                    string sSta = lineStr[10].Trim('\"'); // 11 開業狀況

                    using (BSDataContext dc = new BSDataContext())
                    {
                        var q = from p in dc.NHI_hosp
                                where (p.NHI_code == sCod)
                                select p;
                        if (q.Count() == 0)
                        {
                            try
                            {
                                NHI_hosp newNHI = new NHI_hosp()
                                {
                                    Div = Char.Parse(sDiv),
                                    NHI_code = sCod,
                                    Nam = sNam,
                                    Adr = sAdr,
                                    Loc = sLoc,
                                    Tel = sTel,
                                    Clas = Char.Parse(sCls),
                                    Form = sFor,
                                    Typ = Char.Parse(sTyp),
                                    end_date = sDat,
                                    Stat = Char.Parse(sSta),
                                    QDATE = _qdate
                                };
                                dc.NHI_hosp.InsertOnSubmit(newNHI);
                                dc.SubmitChanges();
                                add_N++;
                            }
                            catch (Exception ex)
                            {
                                Logging.Record_error(ex.Message);
                                log.Error($"{sCod}: [{ex.Message}]");
                            }
                        }
                        else
                        {
                            try
                            {
                                // only one if any
                                bool bChanged = false;
                                string strChange = string.Empty;
                                NHI_hosp oldNHI = q.First();

                                if (oldNHI.Div != char.Parse(sDiv))
                                {
                                    strChange += $"分區別: {oldNHI.Div} => {sDiv};";
                                    oldNHI.Div = char.Parse(sDiv);
                                    bChanged = true;
                                }         // 1  分區別,
                                if (oldNHI.Nam != sNam)
                                {
                                    strChange += $"醫事機構名稱: {oldNHI.Nam} => {sNam};";
                                    oldNHI.Nam = sNam;
                                    bChanged = true;
                                }         // 3  醫事機構名稱,
                                if (oldNHI.Adr != sAdr)
                                {
                                    strChange += $"機構地址: {oldNHI.Adr} => {sAdr};";
                                    oldNHI.Adr = sAdr;
                                    bChanged = true;
                                }         // 4  機構地址,
                                if (oldNHI.Loc != sLoc)
                                {
                                    strChange += $"電話區域號碼: {oldNHI.Loc} => {sLoc};";
                                    oldNHI.Loc = sLoc;
                                    bChanged = true;
                                }         // 5  電話區域號碼,
                                if (oldNHI.Tel != sTel)
                                {
                                    strChange += $"電話號碼: {oldNHI.Tel} => {sTel};";
                                    oldNHI.Tel = sTel;
                                    bChanged = true;
                                }         // 6  電話號碼,
                                if (oldNHI.Clas != char.Parse(sCls))
                                {
                                    strChange += $"特約類別: {oldNHI.Clas} => {sCls};";
                                    oldNHI.Clas = char.Parse(sCls);
                                    bChanged = true;
                                }       // 7  特約類別,
                                if (oldNHI.Form != sFor)
                                {
                                    strChange += $"型態別: {oldNHI.Form} => {sFor};";
                                    oldNHI.Form = sFor;
                                    bChanged = true;
                                }         // 8  型態別,
                                if (oldNHI.Typ != char.Parse(sTyp))
                                {
                                    strChange += $"醫事機構種類: {oldNHI.Typ} => {sTyp};";
                                    oldNHI.Typ = char.Parse(sTyp);
                                    bChanged = true;
                                }     // 9  醫事機構種類,
                                if (oldNHI.end_date != sDat)
                                {
                                    strChange += $"終止合約或歇業日期: {oldNHI.end_date} => {sDat};";
                                    oldNHI.end_date = sDat;
                                    bChanged = true;
                                }     // 10 終止合約或歇業日期,
                                if (oldNHI.Stat != char.Parse(sSta))
                                {
                                    strChange += $"開業狀況: {oldNHI.Stat} => {sSta};";
                                    oldNHI.Stat = char.Parse(sSta);
                                    bChanged = true;
                                }     // 11 開業狀況
                                if (bChanged)
                                {
                                    // 做記錄
                                    Logging.Record_admin("Change hosp data", $"{sCod}: {strChange}");
                                    log.Info($"Change hosp data: {sCod}: {strChange}");
                                    change_N++;
                                }
                                // 做實改變
                                oldNHI.QDATE = _qdate;
                                dc.SubmitChanges();
                            }
                            catch (Exception ex)
                            {
                                Logging.Record_error(ex.Message);
                                log.Error($"{sCod}: [{ex.Message}]");
                            }
                        }
                    }
                    all_N++;
                    report.PercentageComeplete = all_N * 100 / totalN;
                    progress.Report(report);
                }
                log.Info($"    exit ImportHOSP_async.");
            });
            return new PTresult()
            {
                NewPT = add_N,
                ChangePT = change_N,
                AllPT = all_N
            };
        }
    }
}