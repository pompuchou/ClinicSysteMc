using AutoIt;
using ClinicSysteMc.Model;
using Hardcodet.Wpf.TaskbarNotification;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ClinicSysteMc.ViewModel.Converters
{
    internal class DEPchange
    {
        // 20190606 created, 目的再深化自動化
        // 20190608 加好了try, record_adm, record_err
        // 目前穩定,已經使用了大約一年, 意思是用前身AutoIt版本, 大概是201806中開始的
        // 20190607 created
        // 20200518 transcribed into c-sharp

        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private readonly TaskbarIcon tb = new TaskbarIcon();
        private readonly string _strYM;

        public DEPchange(string YM)
        {
            _strYM = YM;
        }

        public void Change()
        {
            // Dim output As DEP_return = Change_DEP(strYM)
            // MessageBox.Show("修改了" + output.m.ToString + "筆, 請匯入門診資料")
            string savepath = @"C:\vpn\change_dep";
            int change_N;
            DateTime minD;
            DateTime maxD;
            List<sp_change_depResult> ListChange;

            // 存放目錄,不存在就要建立一個
            if (!(System.IO.Directory.Exists(savepath))) System.IO.Directory.CreateDirectory(savepath);

            #region Making CSV

            try
            {
                // 呼叫SQL stored procedure
                using (CSDataContext dc = new CSDataContext())
                {
                    ListChange = dc.sp_change_dep(_strYM).ToList();
                }

                // 自動產生名字
                string savefile = $"\\change_dep_{_strYM}_{DateTime.Now.Year}{(DateTime.Now.Month + 100).ToString().Substring(1)}";
                savefile += $"{(DateTime.Now.Day + 100).ToString().Substring(1)}_{DateTime.Now.TimeOfDay}";
                savefile = savefile.Replace(":", "").Replace(".", "");
                savepath += $"{savefile}.csv";

                // 製作csv檔 writing to csv
                System.IO.StreamWriter sw = new System.IO.StreamWriter(savepath);
                int i = 1;
                change_N = ListChange.Count;
                if (change_N == 0)
                {
                    tb.ShowBalloonTip("完成", "沒有什麼需要修改的", BalloonIcon.Info);
                    log.Info("change department: 沒有什麼需要修改的");
                    Logging.Record_admin("change department", "沒有什麼需要修改的");
                }
                else
                {
                    minD = DateTime.Parse("9999/12/31");
                    maxD = DateTime.Parse("0001/01/01");
                    foreach (var c in ListChange)
                    {
                        sw.Write(c.o); // 欄位名叫o
                        if (i < change_N) sw.Write(sw.NewLine);
                        DateTime tempD = DateTime.Parse($"{c.o.Substring(0, 4)}/{c.o.Substring(4, 2)}/{c.o.Substring(6, 2)}");
                        // 找尋最大的值
                        if (tempD.CompareTo(maxD) > 0) maxD = tempD;
                        // 找尋最小的值
                        if (tempD.CompareTo(minD) < 0) minD = tempD;
                        i++;
                    }
                    // 20200518 放在foreach的loop迴圈裡是錯誤的, 我把它放出來了
                    string output = $"{minD:d}~{maxD:d}, 共{change_N}筆需要修改";
                    tb.ShowBalloonTip("需修改:", output, BalloonIcon.Info);
                    log.Info($"change department: {output}");
                    Logging.Record_admin("change department", output);
                    sw.Close();
                }
            }
            catch (System.Exception ex)
            {
                Logging.Record_error(ex.Message);
                log.Error(ex.Message);
                return;
            }

            #endregion Making CSV

            #region Environment

            try
            {
                // 營造環境
                if (AutoItX.WinExists("看診清單") == 1) //如果直接存在就直接叫用
                {
                    AutoItX.WinActivate("看診清單");
                }
                else
                {
                    Thesis.LogIN();
                    // 打開"看診清單"
                    AutoItX.Run(@"C:\Program Files (x86)\THESE\杏雲醫療資訊系統\THCClinic.exe", @"C:\Program Files (x86)\THESE\杏雲醫療資訊系統\");
                    AutoItX.WinWaitActive("看診清單");
                    AutoItX.WinActivate("看診清單");
                }
            }
            catch (Exception ex)
            {
                Logging.Record_error(ex.Message);
                log.Error(ex.Message);
                return;
            }

            #endregion Environment

            #region Execute change department

            try
            {
                string strD_o = string.Empty;
                string strV_o = string.Empty;
                string strR_o = string.Empty;

                foreach (var c in ListChange)
                {
                    string strD_n = c.o.Substring(0, 8);
                    string strV_n = c.o.Substring(8, 1);
                    string strR_n = c.o.Substring(9, 2);
                    bool changed = false;

                    string strNr = c.o.Substring(11, 3);
                    string strDEP = c.o.Substring(14, 2);

                    // 先檢查是否換日, 如果有就換到新日期;
                    if (strD_n != strD_o)
                    {
                        // 製造3個AutoIT VB程式, 1. changeDP_DATE, 針對"看診清單", [NAME:dtpSDate]
                        // 一個參數, 格式YYYYMMDD
                        AutoItX.Run($"C:\\vpn\\exe\\changeDP_DATE.exe {strD_n}", @"C:\vpn\exe\");
                        log.Info($"Change Date from {strD_o} to {strD_n}.  Then sleep 1000ms.");
                        strD_o = strD_n;
                        changed = true;
                        AutoItX.Sleep(1000);
                    }

                    // 再檢查是否換午別, 如果有就換到新的午別
                    if (strV_n != strV_o)
                    {
                        // 製造3個AutoIT VB程式, 2. changeDP_VIST, 針對"看診清單", [NAME:cmbVist]
                        // 一個參數, 格式V
                        AutoItX.Run($"C:\\vpn\\exe\\changeDP_VIST.exe {strV_n}", @"C:\vpn\exe\");
                        log.Info($"Change Vist from {strV_o} to {strV_n}. Then sleep 1000ms.");
                        strV_o = strV_n;
                        changed = true;
                        AutoItX.Sleep(1000);
                    }

                    // 最後檢查是否換診間, 如果有就換到新的診間
                    if (strR_n != strR_o)
                    {
                        // 製造3個AutoIT VB程式, 3. changeDP_ROOM, 針對"看診清單", [NAME:cmbRmno]
                        // 一個參數, 格式RR
                        AutoItX.Run($"C:\\vpn\\exe\\changeDP_ROOM.exe {strR_n}", @"C:\vpn\exe\");
                        log.Info($"Change Date from {strR_o} to {strR_n}. Then sleep 1000ms.");
                        strR_o = strR_n;
                        changed = true;
                        AutoItX.Sleep(1000);
                    }

                    if (changed)
                    {
                        AutoItX.ControlClick("看診清單", "", "[NAME:btnRefresh]");
                        log.Info($"Press Refresh button. [NAME:btnRefresh].");
                        AutoItX.WinWaitActive("看診清單");
                        changed = false;
                    }

                    // 按下新的號碼;
                    // Seq NO
                    // 輸入診號
                    AutoItX.ControlSend("看診清單", "", "[NAME:txbSqno]", strNr);
                    log.Info($"Enter {strNr}.");

                    // 按鈕
                    AutoItX.ControlClick("看診清單", "", "[NAME:btnGo]");
                    log.Info($"Press GO button. [NAME:btnGo].");

                    // 進入問診畫面
                    AutoItX.WinWaitActive("問診畫面");
                    AutoItX.Sleep(500);

                    // 製造一個AutoIT VB程式, changeDP_DEP, 針對"問診畫面"[NAME:cmbDept]
                    // 一個參數, 格式DD
                    AutoItX.Run($"C:\\vpn\\exe\\changeDP_DEP.exe {strDEP}", @"C:\vpn\exe\");

                    AutoItX.Sleep(500);
                    int idx = 0;

                    // 先問「此病患已批價, 是否繼續?」, THCClinic, 還可能問重大傷病, 超過8種藥物, 可能有重複用藥畫面
                    do
                    {
                        AutoItX.Send("{F9}");
                        AutoItX.Sleep(100);
                        idx++; // time out for 10 sec at most
                    } while (AutoItX.WinExists("THCClinic") == 0 && idx < 100);

                    // 下一個畫面「確定要重複開立收據」
                    idx = 0;
                    do
                    {
                        AutoItX.ControlClick("THCClinic", "", "[CLASSNN:Button1]");
                        AutoItX.Sleep(100);
                        idx++; // time out for 10 sec at most
                    } while (AutoItX.WinExists("These.CludCln.Accounting") == 0 && idx < 100);

                    // 是否重印收據
                    AutoItX.WinWaitActive("These.CludCln.Accounting");
                    AutoItX.ControlClick("These.CludCln.Accounting", "", "[CLASSNN:Button2]");
                    AutoItX.Sleep(1000);
                }

                Logging.Record_admin("change department", $"修改了{change_N}筆");
                tb.ShowBalloonTip("完成", $"修改了{change_N}筆, 請匯入門診資料.", BalloonIcon.Info);
            }
            catch (Exception ex)
            {
                Logging.Record_error(ex.Message);
                log.Error(ex.Message);
                return;
            }

            #endregion Execute change department
        }
    }
}