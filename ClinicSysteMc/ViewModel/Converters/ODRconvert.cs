using ClinicSysteMc.Model;
using Hardcodet.Wpf.TaskbarNotification;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ClinicSysteMc.ViewModel.Converters
{
    internal class ODRconvert : IDisposable
    {
        private readonly object[,] _data;
        private readonly DateTime _qdate;
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private readonly TaskbarIcon tb = new TaskbarIcon();
        private bool _disposed = false;

        public ODRconvert(object[,] Data)
        {
            _data = Data;
            _qdate = DateTime.Now;
        }

        public async void Transform(Progress<ProgressReportModel> progress)
        {
            // 檢查檔案格式
            // 可以算出總筆數,第一行是標題,不算
            string[] strT = {"醫令碼", "英文規格", "生效日期", "截止日期", "健保碼", "醫令簡碼", "中文規格", "學名", "類別", "健保價",
                             "自費價", "批價單位", "批價比率", "使用單位", "頻率", "途徑", "天數", "調劑方式", "最小劑量", "最大總量",
                             "最大天數", "展開方式", "集合醫令明細", "劑型", "副作用", "用途", "用藥指示", "外觀", "成分含量", "廠牌",
                             "用藥/排程說明", "藥品備註", "許可證字號", "安全存量", "臨界存量", "給付類別", "疫苗給付類別", "特定治療項目",
                             "檢驗代碼", "案件註記", "服務機構代號", "處置碼", "檢查儀器", "停用日期", "有效醫令", "管制藥品", "管制藥品",
                             "磨粉", "病摘", "療程", "診斷書", "門診使用", "門診缺藥", "替換代碼", "常用", "列印", "檢核類型", "檢核起",
                             "檢核迄", "檢核性別", "異動人員", "異動日期"};
            for (int i = 1; i <= strT.Length; i++)
            {
                if ((string)_data[1, i] != strT[i - 1])
                {
                    // 寫入Error Log
                    Logging.Record_error(" 輸入的醫令資料檔案格式不對");
                    log.Error("輸入的醫令資料檔案格式不對");
                    tb.ShowBalloonTip("錯誤", "檔案格式不對", BalloonIcon.Error);
                    return;
                }
            }

            // 通過測試
            Logging.Record_admin("計價檔格式", "correct");
            log.Info("輸入的醫令資料檔案格式正確");
            tb.ShowBalloonTip("正確", "檔案格式正確", BalloonIcon.Info);

            // _data is a 2-dimentional array
            // _data all begin with 1, in dimension 1, and dimension 2
            int totalN = _data.GetUpperBound(0) - 1;  // -1 because line 1 is titles, so I should begin with 2 to total_N + 1
            // now I should divide the array into 500 lines each and store it into a list.

            int table_N = 250;
            int total_div = totalN / table_N;
            int residual = totalN % table_N;
            int item_n = strT.Length;

            log.Info($"  start async process.");
            List<Task<ODRresult>> tasks = new List<Task<ODRresult>>();

            for (int i = 0, idx = item_n + 1; i <= total_div; i++, idx += (table_N * item_n))
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
                tasks.Add(ImportODR_async(dummy, progress));
            }

            ODRresult[] result = await Task.WhenAll(tasks);

            int total_NewODR = (from p in result
                                select p.NewODR).Sum();
            int total_ChangeODR = (from p in result
                                   select p.ChangeODR).Sum();
            int total_AllODR = (from p in result
                                select p.AllODR).Sum();
            log.Info($"  end async process.");

            string output = $"共處理{total_AllODR}筆資料, 其中{total_NewODR}筆新醫令, 修改{total_ChangeODR}筆醫令.";
            log.Info(output);
            tb.ShowBalloonTip("完成", output, BalloonIcon.Info);
            Logging.Record_admin("ODR add/change", output);

            this.Dispose();
            return;
        }

        private async Task<ODRresult> ImportODR_async(object[,] data, IProgress<ProgressReportModel> progress)
        {
            int totalN = data.GetUpperBound(0);
            int add_N = 0;
            int change_N = 0;
            int all_N = 0;
            ProgressReportModel report = new ProgressReportModel();

            await Task.Run(() =>
            {
                // 要有迴路, 來讀一行一行的xls, 能夠判斷
                for (int i = 0; i <= totalN; i++)
                {
                    // 先判斷是否已經在資料表中, 如果不是就insert否則判斷要不要update
                    CSDataContext dc = new CSDataContext();
                    string strRID = string.Empty;

                    // 先判斷醫令代碼是否空白, 原本第1, 現在第0
                    if (string.IsNullOrEmpty((string)data[i, 0]))
                    {
                        // 寫入Error Log
                        // 沒有醫令代碼是不行的
                        //Logging.Record_error("醫令代碼是空的");
                        log.Error("醫令代碼是空的");
                        return;
                    }

                    // 再判斷是否已在資料表中
                    strRID = (string)data[i, 0];    //醫令代碼,第0欄
                    var od = from d in dc.p_order
                             where d.rid == strRID
                             select d;    // this is a querry

                    if (od.Count() == 0)
                    {
                        // insert
                        // 沒這個醫令可以新增這個醫令
                        // 填入資料
                        try
                        {
                            p_order newOd = new p_order()
                            {
                                rid = strRID,
                                r01 = (string)data[i, 1],  //英文規格, 第2欄, 20200513 改成1欄
                                r02 = (string)data[i, 2],     //生效日期,第3欄, 20200513 改成2欄
                                r03 = (string)data[i, 3],  //截止日期,第4欄, 20200513 改成3欄
                                r04 = (string)data[i, 4], // 健保碼, 第5欄, 20200513 改成4欄
                                r06 = (string)data[i, 6],  //中文規格, 第7欄, 20200513 改成6欄
                                r07 = (string)data[i, 7],  //學名, 第8欄, 20200513 改成7欄
                                r08 = (string)data[i, 8],  //類別,第9欄, 20200513 改成8欄
                                r09 = (string)data[i, 9],  //健保價,第10欄, 20200513 改成9欄
                                r10 = (string)data[i, 10],  //自費價, 第11欄, 20200513 改成10欄
                                r13 = (string)data[i, 13],  //使用單位, 第14欄, 20190611 改成第16欄, 20200319 改成14欄, 20200513 改成13欄
                                r14 = (string)data[i, 11],  //批價單位, 第15欄, 20190611 改成第14欄, 20200319 改成12欄, 20200513 改成11欄
                                r15 = (string)data[i, 14],  //頻率, 第16欄, 20190611 改成第17欄, 20200319 改成15欄, 20200513 改成14欄
                                r16 = (string)data[i, 15],  //途徑, 第17欄, 20190611 改成第18欄, 20200319 改成16欄, 20200513 改成15欄
                                r18 = (string)data[i, 17],  //調劑方式, 第19欄, 20190611 改成第20欄, 20200319 改成18欄, 20200513 改成17欄
                                r19 = (string)data[i, 12],  //批價比率, 第20欄, 20190611 改成第15欄, 20200319 改成13欄, 20200513 改成12欄
                                r25 = (string)data[i, 23],  //劑型, 第26欄, 20200319 改成24欄, 20200513 改成23欄
                                r26 = (string)data[i, 24],  //副作用, 第27欄, 20200319 改成25欄, 20200513 改成24欄
                                r27 = (string)data[i, 25],  //用途, 第28欄, 20200319 改成26欄, 20200513 改成25欄
                                r28 = (string)data[i, 26],  //用藥指示, 第29欄, 20200319 改成27欄, 20200513 改成26欄
                                r29 = (string)data[i, 27],  //外觀, 第30欄, 20200319 改成28欄, 20200513 改成27欄
                                r30 = (string)data[i, 28],  //程分含量, 第31欄, 20200319 改成29欄, 20200513 改成28欄
                                r31 = (string)data[i, 29],  //廠牌, 第32欄, 20200319 改成30欄, 20200513 改成29欄
                                r32 = (string)data[i, 30],  //用藥/排程說明, 第33欄, 20200319 改成31欄, 20200513 改成30欄
                                r33 = (string)data[i, 31],  //藥品備註, 第34欄, 20200319 改成32欄, 20200513 改成31欄
                                r34 = (string)data[i, 32],  //許可證字號, 第35欄, 20200319 改成33欄, 20200513 改成32欄
                                r40 = (string)data[i, 38],  //檢驗代碼, 第41欄, 20200319 改成39欄, 20200513 改成38欄
                                r48 = (string)data[i, 46],  //管制藥品, 第49欄, 20190611 改成第48欄, 20200319 改成47欄, 20200513 改成46欄
                                r52 = (string)data[i, 52],  //門診缺藥, 第53欄, 20190929 改成第54欄, 20200319 改成53欄, 20200513 改成52欄
                                r60 = (string)data[i, 60],  //異動人員, 第61欄, 20190929 改成第62欄, 20200319 改成61欄, 20200513 改成60欄
                                r61 = (string)data[i, 61],  //異動日期, 第62欄, 20190929 改成第63欄, 20200319 改成62欄, 20200513 改成61欄
                                QDATE = _qdate
                            };
                            dc.p_order.InsertOnSubmit(newOd);
                            dc.SubmitChanges();
                            Logging.Record_admin("Add a new order", strRID);
                            log.Info($"Add a new order: {strRID} {data[i, 1]}");
                            add_N++;
                        }
                        catch (Exception ex)
                        {
                            Logging.Record_error(ex.Message);
                            log.Error(ex.Message);
                        }
                    }
                    else
                    {
                        // update
                        // 有此醫令喔, 走update方向
                        // 拿oldOd比較data[i),如果不同就修改,並且記錄

                        p_order oldOd = (from d in dc.p_order
                                         where d.rid == strRID
                                         select d).First();  // this is a record
                        string strChange = string.Empty;
                        bool bChange = false;

                        try
                        {
                            // 英文規格, 第2欄, 20200513 改成1欄
                            if (oldOd.r01 != (string)data[i, 1])
                            {
                                strChange += $"改英文規格: {oldOd.r01} => {data[i, 1]}; ";
                                bChange = true;
                                oldOd.r01 = (string)data[i, 1];
                            }
                            // 生效日期,第3欄, 20200513 改成2欄
                            if (oldOd.r02 != (string)data[i, 2])
                            {
                                strChange += $"改生效日期: {oldOd.r02} => {data[i, 2]}; ";
                                bChange = true;
                                oldOd.r02 = (string)data[i, 2];
                            }
                            // 截止日期,第4欄, 20200513 改成3欄
                            if (oldOd.r03 != (string)data[i, 3])
                            {
                                strChange += $"改截止日期: {oldOd.r03} => {data[i, 3]}; ";
                                bChange = true;
                                oldOd.r03 = (string)data[i, 3];
                            }
                            // 健保碼, 第5欄, 20200513 改成4欄
                            if (oldOd.r04 != (string)data[i, 4])
                            {
                                strChange += $"改健保碼: {oldOd.r04} => {data[i, 4]}; ";
                                bChange = true;
                                oldOd.r04 = (string)data[i, 4];
                            }
                            // 中文規格, 第7欄, 20200513 改成6欄
                            if (oldOd.r06 != (string)data[i, 6])
                            {
                                strChange += $"改中文規格: {oldOd.r06} => {data[i, 6]}; ";
                                bChange = true;
                                oldOd.r06 = (string)data[i, 6];
                            }
                            // 學名, 第8欄, 20200513 改成7欄
                            if (oldOd.r07 != (string)data[i, 7])
                            {
                                strChange += $"改學名: {oldOd.r07} => {data[i, 7]}; ";
                                bChange = true;
                                oldOd.r07 = (string)data[i, 7];
                            }
                            // 類別,第9欄, 20200513 改成8欄
                            if (oldOd.r08 != (string)data[i, 8])
                            {
                                strChange += $"改類別: {oldOd.r08} => {data[i, 8]}; ";
                                bChange = true;
                                oldOd.r08 = (string)data[i, 8];
                            }
                            // 健保價,第10欄, 20200513 改成9欄
                            if (oldOd.r09 != (string)data[i, 9])
                            {
                                strChange += $"改健保價: {oldOd.r09} => {data[i, 9]}; ";
                                bChange = true;
                                oldOd.r09 = (string)data[i, 9];
                            }
                            // 自費價, 第11欄, 20200513 改成10欄
                            if (oldOd.r10 != (string)data[i, 10])
                            {
                                strChange += $"改自費價: {oldOd.r10} => {data[i, 10]}; ";
                                bChange = true;
                                oldOd.r10 = (string)data[i, 10];
                            }
                            // 使用單位, 第14欄, 20190611 改成第16欄, 20200319 改成14欄, 20200513 改成13欄
                            if (oldOd.r13 != (string)data[i, 13])
                            {
                                strChange += $"改使用單位: {oldOd.r13} => {data[i, 13]}; ";
                                bChange = true;
                                oldOd.r13 = (string)data[i, 13];
                            }
                            // 批價單位, 第15欄, 20190611 改成第14欄, 20200319 改成12欄, 20200513 改成11欄
                            if (oldOd.r14 != (string)data[i, 11])
                            {
                                strChange += $"改批價單位: {oldOd.r14} => {data[i, 11]}; ";
                                bChange = true;
                                oldOd.r14 = (string)data[i, 11];
                            }
                            // 頻率, 第16欄, 20190611 改成第17欄, 20200319 改成15欄, 20200513 改成14欄
                            if (oldOd.r15 != (string)data[i, 14])
                            {
                                strChange += $"改頻率: {oldOd.r15} => {data[i, 14]}; ";
                                bChange = true;
                                oldOd.r15 = (string)data[i, 14];
                            }
                            // 途徑, 第17欄, 20190611 改成第18欄, 20200319 改成16欄, 20200513 改成15欄
                            if (oldOd.r16 != (string)data[i, 15])
                            {
                                strChange += $"改途徑: {oldOd.r16} => {data[i, 15]}; ";
                                bChange = true;
                                oldOd.r16 = (string)data[i, 15];
                            }
                            // 調劑方式, 第19欄, 20190611 改成第20欄, 20200319 改成18欄, 20200513 改成17欄
                            if (oldOd.r18 != (string)data[i, 17])
                            {
                                strChange += $"改調劑方式: {oldOd.r18} => {data[i, 17]}; ";
                                bChange = true;
                                oldOd.r18 = (string)data[i, 17];
                            }
                            // 批價比率, 第20欄, 20190611 改成第15欄, 20200319 改成13欄, 20200513 改成12欄
                            if (oldOd.r19 != (string)data[i, 12])
                            {
                                strChange += $"改批價比率: {oldOd.r19} => {data[i, 12]}; ";
                                bChange = true;
                                oldOd.r19 = (string)data[i, 12];
                            }
                            // 劑型, 第26欄, 20200319 改成24欄, 20200513 改成23欄
                            if (oldOd.r25 != (string)data[i, 23])
                            {
                                strChange += $"改劑型: {oldOd.r25} => {data[i, 23]}; ";
                                bChange = true;
                                oldOd.r25 = (string)data[i, 23];
                            }
                            // 副作用, 第27欄, 20200319 改成25欄, 20200513 改成24欄
                            if (oldOd.r26 != (string)data[i, 24])
                            {
                                strChange += $"改副作用: {oldOd.r26} => {data[i, 24]}; ";
                                bChange = true;
                                oldOd.r26 = (string)data[i, 24];
                            }
                            // 用途, 第28欄, 20200319 改成26欄, 20200513 改成25欄
                            if (oldOd.r27 != (string)data[i, 25])
                            {
                                strChange += $"改用途: {oldOd.r27} => {data[i, 25]}; ";
                                bChange = true;
                                oldOd.r27 = (string)data[i, 25];
                            }
                            // 用藥指示, 第29欄, 20200319 改成27欄, 20200513 改成26欄
                            if (oldOd.r28 != (string)data[i, 26])
                            {
                                strChange += $"改用藥指示: {oldOd.r28} => {data[i, 26]}; ";
                                bChange = true;
                                oldOd.r28 = (string)data[i, 26];
                            }
                            // 外觀, 第30欄, 20200319 改成28欄, 20200513 改成27欄
                            if (oldOd.r29 != (string)data[i, 27])
                            {
                                strChange += $"改外觀: {oldOd.r29} => {data[i, 27]}; ";
                                bChange = true;
                                oldOd.r29 = (string)data[i, 27];
                            }
                            // 成分含量, 第31欄, 20200319 改成29欄, 20200513 改成28欄
                            if (oldOd.r30 != (string)data[i, 28])
                            {
                                strChange += $"改成分含量: {oldOd.r30} => {data[i, 28]}; ";
                                bChange = true;
                                oldOd.r30 = (string)data[i, 28];
                            }
                            // 廠牌, 第32欄, 20200319 改成30欄, 20200513 改成29欄
                            if (oldOd.r31 != (string)data[i, 29])
                            {
                                strChange += $"改廠牌: {oldOd.r31} => {data[i, 29]}; ";
                                bChange = true;
                                oldOd.r31 = (string)data[i, 29];
                            }
                            // 用藥/排程說明, 第33欄, 20200319 改成31欄, 20200513 改成30欄
                            if (oldOd.r32 != (string)data[i, 30])
                            {
                                strChange += $"改用藥排程說明: {oldOd.r32} => {data[i, 30]}; ";
                                bChange = true;
                                oldOd.r32 = (string)data[i, 30];
                            }
                            // 藥品備註, 第34欄, 20200319 改成32欄, 20200513 改成31欄
                            if (oldOd.r33 != (string)data[i, 31])
                            {
                                strChange += $"改藥品備註: {oldOd.r33} => {data[i, 31]}; ";
                                bChange = true;
                                oldOd.r33 = (string)data[i, 31];
                            }
                            // 許可證字號, 第35欄, 20200319 改成33欄, 20200513 改成32欄
                            if (oldOd.r34 != (string)data[i, 32])
                            {
                                strChange += $"改許可證字號: {oldOd.r34} => {data[i, 32]}; ";
                                bChange = true;
                                oldOd.r34 = (string)data[i, 32];
                            }
                            // 檢驗代碼, 第41欄, 20200319 改成39欄, 20200513 改成38欄
                            if (oldOd.r40 != (string)data[i, 38])
                            {
                                strChange += $"改檢驗代碼: {oldOd.r40} => {data[i, 38]}; ";
                                bChange = true;
                                oldOd.r40 = (string)data[i, 38];
                            }
                            // 管制藥品, 第49欄, 20190611 改成第48欄, 20200319 改成47欄, 20200513 改成46欄
                            if (oldOd.r48 != (string)data[i, 46])
                            {
                                strChange += $"改管制藥品: {oldOd.r48} => {data[i, 46]}; ";
                                bChange = true;
                                oldOd.r48 = (string)data[i, 46];
                            }
                            // 門診缺藥, 第53欄, 20190929 改成第54欄, 20200319 改成53欄, 20200513 改成52欄
                            if (oldOd.r52 != (string)data[i, 52])
                            {
                                strChange += $"改門診缺藥: {oldOd.r52} => {data[i, 52]}; ";
                                bChange = true;
                                oldOd.r52 = (string)data[i, 52];
                            }
                            // 異動人員, 第61欄, 20190929 改成第62欄, 20200319 改成61欄, 20200513 改成60欄
                            if (oldOd.r60 != (string)data[i, 60])
                            {
                                strChange += $"改異動人員: {oldOd.r60} => {data[i, 60]}; ";
                                bChange = true;
                                oldOd.r60 = (string)data[i, 60];
                            }
                            // 異動日期, 第62欄, 20190929 改成第63欄, 20200319 改成62欄, 20200513 改成61欄
                            if (oldOd.r61 != (string)data[i, 61])
                            {
                                strChange += $"改異動日期: {oldOd.r61} => {data[i, 61]}; ";
                                bChange = true;
                                oldOd.r61 = (string)data[i, 61];
                            }
                            if (bChange == true)
                            {
                                //  做實改變
                                oldOd.QDATE = _qdate;
                                dc.SubmitChanges();
                                // 做記錄
                                Logging.Record_admin("Change order data", $"{strRID}: {strChange}");
                                log.Info($"Change order data: [{strRID}] {data[i, 1]}: {strChange}");
                                change_N++;
                            }
                        }
                        catch (Exception ex)
                        {
                            Logging.Record_error(ex.Message);
                            log.Error(ex.Message);
                        }
                    }
                    all_N++;
                    report.PercentageComeplete = (all_N * 100) / totalN;
                    progress.Report(report);
                }
            });
            return new ODRresult()
            {
                NewODR = add_N,
                ChangeODR = change_N,
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