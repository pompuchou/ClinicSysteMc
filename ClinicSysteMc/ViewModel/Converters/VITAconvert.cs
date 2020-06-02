using ClinicSysteMc.Model;
using Hardcodet.Wpf.TaskbarNotification;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ClinicSysteMc.ViewModel.Converters
{
    internal class VITAconvert : IDisposable
    {
        private readonly object[,] _data;
        private readonly DateTime _qdate;
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private readonly TaskbarIcon tb = new TaskbarIcon();
        private bool _disposed = false;

        public VITAconvert(object[,] Data)
        {
            _data = Data;
            _qdate = DateTime.Now;
        }

        public async Task Transform(Progress<ProgressReportModel> progress)
        {
            // 檢查檔案格式
            // 可以算出總筆數,第一行是標題,不算
            // 開單日期	採檢日期	病患姓名	院所病歷號	身分證號	生日	檢驗項目	代檢費
            string[] strT = { "開單日期", "採檢日期", "病患姓名", "院所病歷號", "身分證號", "生日", "檢驗項目", "代檢費" };
            for (int i = 1; i <= strT.Length; i++)
            {
                if ((string)_data[5, i] != strT[i - 1])
                {
                    // 寫入Error Log
                    Logging.Record_error(" 輸入的賽亞對帳檔案格式不對");
                    log.Error("輸入的賽亞對帳檔案格式不對");
                    tb.ShowBalloonTip("錯誤", "檔案格式不對", BalloonIcon.Error);
                    return;
                }
            }

            // 通過測試
            Logging.Record_admin("賽亞對帳檔格式", "correct");
            log.Info("輸入的賽亞對帳檔案格式正確");
            tb.ShowBalloonTip("正確", "賽亞對帳檔格式正確", BalloonIcon.Info);

            // _data is a 2-dimentional array
            // _data all begin with 1, in dimension 1, and dimension 2
            int totalN = _data.GetUpperBound(0) - 6;  // -6 because line 1-5 is titles, and a total, so I should begin with 6 to total_N
            // now I should divide the array into 500 lines each and store it into a list.

            int table_N = 100;
            int total_div = totalN / table_N;
            int residual = totalN % table_N;
            int item_n = strT.Length;
            ProgressReportModel report = new ProgressReportModel();

            log.Info($"  start async process.");
            List<Task<ODRresult>> tasks = new List<Task<ODRresult>>();

            for (int i = 0, idx = 5 * item_n + 1; i <= total_div; i++, idx += (table_N * item_n))
            {
                object[,] dummy;
                if (i < total_div)
                {
                    dummy = new object[table_N, item_n];
                    Array.Copy(_data, idx, dummy, 0, table_N * item_n);
                }
                else
                {
                    dummy = new object[residual, item_n];
                    Array.Copy(_data, idx, dummy, 0, residual * item_n);
                }
                tasks.Add(ImportVITA_async(dummy, progress, report));
            }

            ODRresult[] result = await Task.WhenAll(tasks);

            int total_NewODR = (from p in result
                                select p.NewODR).Sum();
            int total_AllODR = (from p in result
                                select p.AllODR).Sum();
            log.Info($"  end async process.");

            string output = $"共處理{total_AllODR}筆對帳資料, 其中{total_NewODR}筆新對帳資料.";
            log.Info(output);
            tb.ShowBalloonTip("完成", output, BalloonIcon.Info);
            Logging.Record_admin("ODR add/change", output);

            this.Dispose();
            return;
        }

        private async Task<ODRresult> ImportVITA_async(object[,] data, IProgress<ProgressReportModel> progress, ProgressReportModel report)
        {
            int totalN = data.GetUpperBound(0);
            int add_N = 0;
            int all_N = 0;

            await Task.Run(() =>
            {
                log.Info($"    enter ImportVITA_async.");
                // 要有迴路, 來讀一行一行的xls, 能夠判斷
                for (int i = 0; i <= totalN; i++)
                {
                    // 先判斷是否已經在資料表中, 如果不是就insert否則判斷要不要update
                    BSDataContext dc = new BSDataContext();
                    string sKaiDan = string.Empty;

                    // 先判斷開單日期是否空白, 原本第1, 現在第0
                    if (string.IsNullOrEmpty((string)data[i, 0]))
                    {
                        // 寫入Error Log
                        // 沒有開單日期是不行的
                        //Logging.Record_error("醫令代碼是空的");
                        log.Error("開單日期是空的");
                        // 20200528 發現這裡用return是不對的, continue才對
                        //return;
                        continue;
                    }

                    // 開單日期	採檢日期	病患姓名	院所病歷號	身分證號	生日	檢驗項目	代檢費

                    // 再判斷是否已在資料表中
                    // 0 開單日期
                    sKaiDan = (string)data[i, 0];    // 開單日期,第0欄
                    string[] saKaiDan = sKaiDan.Split('/');
                    DateTime dKaiDan = DateTime.Parse($"{int.Parse(saKaiDan[0])+1911}/{saKaiDan[1]}/{saKaiDan[2]}");
                    // 1 採檢日期
                    string sCaiJian = (string)data[i, 1];
                    string[] saCaiJian = sCaiJian.Split('/');
                    DateTime dCaiJian = DateTime.Parse($"{int.Parse(saCaiJian[0]) + 1911}/{saCaiJian[1]}/{saCaiJian[2]}");
                    // 2 病患姓名
                    string sCname = (string)data[i, 2];
                    // 3 院所病歷號
                    string sCid = (string)data[i, 3];
                    // 4 身分證號
                    string sUID = (string)data[i, 4]; // 身分證號, 第4欄 
                    // 5 生日
                    string sBD = (string)data[i, 5];
                    string[] saBD = sBD.Split('/');
                    DateTime dBD = DateTime.Parse($"{int.Parse(saBD[0]) + 1911}/{saBD[1]}/{saBD[2]}");
                    // 6 檢驗項目
                    string sItems = (string)data[i, 6];
                    // 7 代檢費
                    int iBill = int.Parse(data[i, 7].ToString());

                    var od = from d in dc.VITA_bill 
                             where (d.KaiDan == dKaiDan) && (d.CaiJian == dCaiJian) && (d.uid == sUID) && (d.items == sItems)
                             select d;    // this is a querry
                    if (od.Count() == 0)
                    {
                        // insert
                        // 沒這個醫令可以新增這個醫令
                        // 填入資料
                        try
                        {
                            VITA_bill newBill = new VITA_bill()
                            {
                                KaiDan = dKaiDan,          // 0 開單日期
                                CaiJian = dCaiJian,        // 1 採檢日期
                                cname = sCname,            // 2 病患姓名
                                cid = sCid,                // 3 院所病歷號
                                uid = sUID,                // 4 身分證號
                                bd = dBD,                  // 5 生日
                                items = sItems,            // 6 檢驗項目
                                bill = iBill,              // 7 代檢費
                                QDATE = _qdate
                            };
                            dc.VITA_bill.InsertOnSubmit(newBill);
                            dc.SubmitChanges();
                            log.Info($"Add a new Vita Bill: {sKaiDan}, {sUID}.");
                            add_N++;
                        }
                        catch (Exception ex)
                        {
                            Logging.Record_error(ex.Message);
                            log.Error(ex.Message);
                        }
                }
                all_N++;
                    report.PercentageComeplete = all_N * 100 / totalN;
                    progress.Report(report);
                }
                log.Info($"    exit ImportVITA_async.");
            });

            return new ODRresult()
            {
                NewODR = add_N,
                AllODR = all_N
            };
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (_disposed)
            {
                return;
            }

            if (disposing)
            {
                // Free any other managed objects here.
            }

            _disposed = true;
        }
    }
}