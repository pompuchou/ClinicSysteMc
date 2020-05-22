using ClinicSysteMc.Model;
using Hardcodet.Wpf.TaskbarNotification;
using System;
using System.Linq;

namespace ClinicSysteMc.ViewModel.Converters
{
    internal class LABconvert : IDisposable
    {
        private readonly object[,] _data;
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private readonly TaskbarIcon tb = new TaskbarIcon();
        private bool _disposed = false;

        public LABconvert(object[,] Data)
        {
            _data = Data;
        }

        public void Transform()
        {
            // 檢查檔案格式
            // 檢查第一行的標題,看看是否符合
            string[] strT = new string[] {"", "身份證字號", "病患姓名", "出生日期", "性別", "原病歷號碼", "原就醫日期",
                                        "檢驗單工號", "開單日(收件日)", "開單時間", "檢驗日期", "報告日期", "報告時間",
                                        "就醫序號"};
            for (int i = 1; i <= strT.Length; i++)
            {
                if (strT[i - 1] != ((string)_data[1, i] ?? string.Empty) )
                {
                    // 寫入Error Log
                    Logging.Record_error("輸入的常誠檢驗資料檔案格式不對");
                    log.Error("輸入的常誠檢驗資料檔案格式不對");
                    tb.ShowBalloonTip("錯誤", "輸入的常誠檢驗資料檔案格式不對", BalloonIcon.Error);
                    return;
                }
            }

            // 通過測試
            Logging.Record_admin("Lab file format", "correct");
            log.Info("輸入的常誠檢驗資料檔案格式正確");
            tb.ShowBalloonTip("正確", "常誠檢驗檔案格式正確", BalloonIcon.Info);

            // 要有迴路, 來讀一行一行的xls, 能夠判斷
            // 檔案結構複雜, 不好用for next, 應該用while
            // 一次性讀檔, 不用update
            // totalN+1 是excel檔的總rows數
            int totalN = _data.GetUpperBound(0) - 1;
            int ind = 1;   // index, 從第二行開始
            string strUid = string.Empty;
            string strLid = string.Empty;
            DateTime dL05 = DateTime.Parse("1901/01/01");
            CSDataContext dc = new CSDataContext();

            while (ind <= totalN)
            {
                ind++; //next line
                if ((string)_data[ind, 1] == "***")
                {
                    try
                    {
                        // 檢驗單工號, 第8欄, 檢查是否空白, 空白不行
                        if (_data[ind, 8].ToString().Length == 0 || _data[ind, 2].ToString().Length == 0 ||
                            !DateTime.TryParse((string)_data[ind, 12], out DateTime d1))
                        {
                            strUid = string.Empty;
                            dL05 = DateTime.Parse("1901/01/01");
                            strLid = string.Empty;
                            // 寫入Error Log
                            string output = "輸入檢驗資料時,缺少檢驗單工號,或身分證字號, 或沒有報告日期";
                            Logging.Record_error(output);
                            log.Error(output);
                            continue; // continue while就可以跳下一行
                        }
                        else
                        {
                            strLid = _data[ind, 8].ToString().Trim();
                            var La = from l in dc.tbl_lab
                                     where l.lid == strLid
                                     select l; // a query for searching duplicates
                            if (La.Count() != 0) //如果有重複,不但這行不要讀了, 連帶後面也都不要讀(strLid=""), 直到下次"***"
                            {
                                strUid = string.Empty;
                                dL05 = DateTime.Parse("1901/01/01");
                                strLid = string.Empty;
                                continue; //跳下一行
                            }
                            // 身分證字號, 第2欄, 檢查是否空白, 空白不行
                            strUid = _data[ind, 2].ToString().Trim();
                            // 報告日期, 第12欄, 檢查是否空白, 空白不行
                            dL05 = d1;
                            // 檢查檢驗單工號是否存在,如果有就不要存了
                        }

                        // 寫入資料庫tbl_Lab, uid, lid, cname, bd, mf, cid, l01, l02, l03, l04, l05, l06
                        // 有些變數共用uid, lid, l05
                        //l01, 原就醫日期,刻意留白, 第7欄
                        tbl_lab newLb = new tbl_lab()
                        {
                            uid = strUid,  //身分證字號,第2欄
                            cname = _data[ind, 3].ToString().Trim(),  //病患姓名, 第3欄
                            mf = _data[ind, 5].ToString().Trim(),   //性別,第5欄
                            cid = _data[ind, 6].ToString().Trim(),  //原病歷號碼,第6欄
                            lid = strLid,  //檢驗單工號, 第8欄
                            l03 = _data[ind, 10].ToString().Trim(), //開單時間, 第10欄
                            l05 = dL05,    //報告日期, 第12欄
                            l06 = _data[ind, 13].ToString().Trim()  //報告時間,第13欄
                        };
                        if (DateTime.TryParse((string)_data[ind, 4], out DateTime d2))  //出生日期, 第4欄
                        {
                            newLb.bd = d2;
                        }
                        if (DateTime.TryParse((string)_data[ind, 9], out DateTime d3))  //開單日(收件日), 第9欄
                        {
                            newLb.l02 = d3;
                        }
                        if (DateTime.TryParse((string)_data[ind, 11], out DateTime d4))  //檢驗日期, 第11欄
                        {
                            newLb.l04 = d4;
                        }

                        dc.tbl_lab.InsertOnSubmit(newLb);
                        dc.SubmitChanges();
                    }
                    catch (Exception ex)
                    {
                        // 寫入錯誤訊息
                        Logging.Record_error(ex.Message);
                        log.Error(ex.Message);
                    }
                }
                else
                {
                    try
                    {
                        //如果沒讀過"***"就略過,以防檔案有錯
                        if (strLid.Length == 0 || strUid.Length == 0 || dL05 == DateTime.Parse("1901/01/01")) continue;

                        // 寫入資料庫tbl_Lab_record: uid, lid, l05, iid, l07
                        tbl_lab_record newLbrd = new tbl_lab_record()
                        {
                            uid = strUid,    //身分證字號
                            lid = strLid,    //檢驗單工號
                            l05 = dL05,  //報告日期
                            iid = _data[ind, 1].ToString().Trim(),    //檢驗代碼, 第1欄
                            l07 = _data[ind, 4].ToString().Trim(),    //檢驗值, 第4欄
                            l09 = _data[ind, 5].ToString().Trim()   //異常, 第5欄
                        };
                        dc.tbl_lab_record.InsertOnSubmit(newLbrd);
                        // 寫入資料庫p_lab_temp: l05, iid, l08, l09, l10, l11
                        p_lab_temp newTemp = new p_lab_temp()
                        {
                            l05 = dL05,  //報告日期
                            iid = _data[ind, 1].ToString().Trim(),    //檢驗代碼, 第1欄
                            l08 = _data[ind, 2].ToString().Trim(),   //檢驗名稱, 第2欄
                            l10 = _data[ind, 6].ToString().Trim(),   //單位, 第6欄
                            l11 = _data[ind, 7].ToString().Trim()   //參考值, 第7欄
                        };
                        dc.p_lab_temp.InsertOnSubmit(newTemp);
                        dc.SubmitChanges();
                    }
                    catch (Exception ex)
                    {
                        // 寫入錯誤訊息
                        Logging.Record_error(ex.Message);
                        log.Error(ex.Message);
                    }
                }
            }
            
            this.Dispose();
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