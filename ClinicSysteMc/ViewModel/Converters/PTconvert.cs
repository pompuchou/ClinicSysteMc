using ClinicSysteMc.Model;
using Hardcodet.Wpf.TaskbarNotification;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ClinicSysteMc.ViewModel.Converters
{
    public class PTconvert : IDisposable
    {
        private readonly object[,] _data;
        private readonly DateTime _qdate;
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private readonly TaskbarIcon tb = new TaskbarIcon();
        private bool _disposed = false;

        public PTconvert(object[,] Data)
        {
            _data = Data;
            _qdate = DateTime.Now;
        }

        public async void Transform()
        {
            // 檢查檔案格式
            // 可以算出總筆數,第一行是標題,不算
            string[] strT = { "病歷號", "姓名", "性別", "室內電話", "手機門號", "電子郵件", "傳送日期", "身分證號", "生日", "地址", "提醒" };
            for (int i = 1; i <= strT.Length; i++)
            {
                if ((string)_data[1, i] != strT[i - 1])
                {
                    // 寫入Error Log
                    Logging.Record_error(" 輸入的病患資料檔案格式不對");
                    log.Error("輸入的病患資料檔案格式不對");
                    tb.ShowBalloonTip("錯誤", "檔案格式不對", BalloonIcon.Error);
                    return;
                }
            }

            // 通過測試
            Logging.Record_admin("病患檔案格式", "correct");
            log.Info("輸入的病患資料檔案格式正確");
            tb.ShowBalloonTip("正確", "檔案格式正確", BalloonIcon.Info);

            //System.Windows.MessageBox.Show(_data.GetUpperBound(0).ToString());
            //System.Windows.MessageBox.Show(_data.GetUpperBound(1).ToString());
            // _data is a 2-dimentional array
            // _data all begin with 1, in dimension 1, and dimension 2
            int totalN = _data.GetUpperBound(0) - 1;  // -1 because line 1 is titles, so I should begin with 2 to total_N + 1
            // now I should divide the array into 500 lines each and store it into a list.

            int table_N = 500;
            int total_div = totalN / table_N;
            int residual = totalN % table_N;
            int item_n = strT.Length;

            log.Info($"  start async process.");
            List<Task<PTresult>> tasks = new List<Task<PTresult>>();

            // 將_data分拆成幾個小的Array
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
                tasks.Add(ImportPT_async(dummy));
            }

            PTresult[] result = await Task.WhenAll(tasks);

            int total_NewPT = (from p in result
                               select p.NewPT).Sum();
            int total_ChangePT = (from p in result
                                  select p.ChangePT).Sum();
            int total_AllPT = (from p in result
                               select p.AllPT).Sum();
            log.Info($"  end async process.");

            string output = $"共處理{total_AllPT}筆資料, 其中{total_NewPT}筆新病歷, 修改{total_ChangePT}筆病歷.";
            log.Info(output);
            tb.ShowBalloonTip("完成", output, BalloonIcon.Info);
            Logging.Record_admin("PT add/change", output);

            this.Dispose();
            return;
        }

        private async Task<PTresult> ImportPT_async(object[,] data)
        {
            int totalN = data.GetUpperBound(0);
            int add_N = 0;
            int change_N = 0;
            int all_N = 0;

            await Task.Run(() =>
            {
                // 要有迴路, 來讀一行一行的xls, 能夠判斷
                for (int i = 0; i <= totalN; i++)
                {
                    // 先判斷是否已經在資料表中, 如果不是就insert否則判斷要不要update
                    // 如何判斷是否已經在資料表中?
                    CSDataContext dc = new CSDataContext();
                    string strUID = string.Empty;
                    // 先判斷身分證字號是否空白, 原本第8, 現在第7
                    if (string.IsNullOrEmpty((string)data[i, 7]))
                    {
                        // 寫入Error Log
                        // 沒有身分證字號是不行的
                        Logging.Record_error("身分證字號是空的");
                        log.Error("身分證字號是空的");
                        return;
                    }
                    // 再判斷是否已在資料表中
                    strUID = (string)data[i, 7];    //身分證號,第7欄
                    var pt = from p in dc.tbl_patients
                             where p.uid == strUID
                             select p;    // this is a querry
                    if (pt.Count() == 0)
                    {
                        // insert
                        // 沒這個人可以新增這個人
                        // 填入資料
                        try
                        {
                            tbl_patients newPt = new tbl_patients();
                            if (string.IsNullOrEmpty((string)data[i, 0]))
                            {
                                // 寫入Error Log
                                Logging.Record_error($"{strUID} 沒有病歷號碼");
                                log.Error($"{strUID} 沒有病歷號碼");
                            }
                            else
                            {
                                newPt.cid = long.Parse((string)data[i, 0]);  // 病歷號, 第1欄
                            }
                            newPt.uid = strUID;     // 身分證號,第8欄
                            if (string.IsNullOrEmpty((string)data[i, 1]))
                            {
                                // 寫入Error Log
                                Logging.Record_error($"{strUID} 沒有姓名");
                                log.Error($"{strUID} 沒有姓名");
                            }
                            else
                            {
                                newPt.cname = (string)data[i, 1];  //姓名,第2欄
                            }
                            newPt.mf = (string)data[i, 2]; // 性別, 第3欄
                            if (string.IsNullOrEmpty((string)data[i, 8]))
                            {
                                // 寫入Error Log
                                Logging.Record_error($"{strUID} 沒有生日資料");
                                log.Error($"{strUID} 沒有生日資料");
                            }
                            else
                            {
                                string strD = (string)data[i, 8];   // 生日, 第9欄
                                newPt.bd = DateTime.Parse($"{strD.Substring(0, 4)}/{strD.Substring(4, 2)}/{strD.Substring(6, 2)}");
                            }
                            newPt.p01 = (string)data[i, 3];  // 市內電話, 第4欄
                            newPt.p02 = (string)data[i, 4];  // 手機電話, 第5欄
                            newPt.p03 = (string)data[i, 9];  // 地址,第10欄
                            newPt.p04 = (string)data[i, 10];  // 提醒,第11欄
                            newPt.QDATE = _qdate;
                            // 2020026新增QDATE

                            dc.tbl_patients.InsertOnSubmit(newPt);
                            dc.SubmitChanges();

                            // 20190929 加姓名, 病歷號
                            Logging.Record_admin("Add a new patient", $"{data[i, 0]} {strUID} {data[i, 1]}");
                            log.Info($"Add a new patient: {data[i, 0]} {strUID} {data[i, 1]}");
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
                        // 有此人喔, 走update方向
                        // 拿pt比較ws.cells(i),如果不同就修改,並且記錄
                        tbl_patients oldPt = (from p in dc.tbl_patients
                                              where p.uid == strUID
                                              select p).ToList()[0];     // this is a record
                        string strChange = string.Empty;
                        bool bChange = false;
                        try
                        {
                            // 病歷號, 20200512加上修改病歷號
                            if (string.IsNullOrEmpty((string)data[i, 0]))
                            {
                                // 寫入Error Log
                                Logging.Record_error($"{strUID} 沒有病歷號碼");
                                log.Error($"{strUID} 沒有病歷號碼");
                            }
                            else if (oldPt.cid != long.Parse((string)data[i, 0]))
                            {
                                strChange += $"改病歷號: {oldPt.cid}=>{data[i, 0]}; ";
                                bChange = true;
                                oldPt.cid = long.Parse((string)data[i, 0]);  // 病歷號, 第1欄
                            }
                            // 姓名
                            if (string.IsNullOrEmpty((string)data[i, 1]))
                            {
                                // 寫入Error Log
                                Logging.Record_error(strUID + " 沒有姓名");
                                log.Error($"{strUID} 沒有姓名");
                            }
                            else if (oldPt.cname != (string)data[i, 1])
                            {
                                strChange += $"改名: {oldPt.cname}=>{data[i, 1]}; ";
                                bChange = true;
                                oldPt.cname = (string)data[i, 1];  // 姓名,第2欄
                            }
                            // 性別
                            if (oldPt.mf != (string)data[i, 2])
                            {
                                strChange += $"改性別: {oldPt.mf}=>{data[i, 2]}; ";
                                bChange = true;
                                oldPt.mf = (string)data[i, 2];  // 性別, 第3欄
                            }
                            // 生日
                            if (string.IsNullOrEmpty((string)data[i, 8]))
                            {
                                // 寫入Error Log
                                Logging.Record_error($"{strUID} 沒有生日資料");
                                log.Error($"{strUID} 沒有生日資料");
                            }
                            else
                            {
                                string strBD = (string)data[i, 8];   // 生日, 第9欄
                                DateTime dBD = DateTime.Parse($"{strBD.Substring(0, 4)}/{strBD.Substring(4, 2)}/{strBD.Substring(6, 2)}");
                                if (oldPt.bd != dBD)
                                {
                                    strChange += $"改生日: {oldPt.bd}=>{dBD}; ";
                                    bChange = true;
                                    oldPt.bd = dBD; // 生日,第9欄
                                }
                            }
                            // 市內電話
                            if ((oldPt.p01 ?? string.Empty) != ((string)data[i, 3] ?? string.Empty))
                            {
                                strChange += $"改市內電話: {oldPt.p01}=>{data[i, 3]}; ";
                                bChange = true;
                                oldPt.p01 = (string)data[i, 3];  // 市內電話,第4欄
                            }

                            // 手機電話
                            if ((oldPt.p02 ?? string.Empty) != ((string)data[i, 4] ?? string.Empty))
                            {
                                strChange += $"改手機電話: {oldPt.p02}=>{data[i, 4]}; ";
                                bChange = true;
                                oldPt.p02 = (string)data[i, 4];  // 手機電話,第5欄
                            }

                            // 地址
                            if ((oldPt.p03 ?? string.Empty) != ((string)data[i, 9] ?? string.Empty))
                            {
                                strChange += $"改地址: {oldPt.p03}=>{data[i, 9]}; ";
                                bChange = true;
                                oldPt.p03 = (string)data[i, 9];  // 地址,第10欄
                            }

                            // 提醒
                            if ((oldPt.p04 ?? string.Empty) != ((string)data[i, 10] ?? string.Empty))
                            {
                                strChange += $"改提醒: {oldPt.p04}=>{data[i, 10]}; ";
                                bChange = true;
                                oldPt.p04 = (string)data[i, 10];  // 提醒,第11欄
                            }

                            if (bChange)
                            {
                                // 做實改變
                                // 2020026新增QDATE
                                oldPt.QDATE = _qdate;
                                dc.SubmitChanges();
                                // 做記錄
                                // 20190929 加姓名, 病歷號
                                Logging.Record_admin("Change patient data", $"{data[i, 0]} {strUID} {data[i, 1]}: {strChange}");
                                log.Info($"Change patient data: {data[i, 0]} {strUID} {data[i, 1]}: {strChange}");
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
                }
            });
            return new PTresult()
            {
                NewPT = add_N,
                ChangePT = change_N,
                AllPT = all_N
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