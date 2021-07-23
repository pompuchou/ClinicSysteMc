using AutoIt;
using ClinicSysteMc.Model;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace ClinicSysteMc.ViewModel.Converters
{
    public partial class Dash
    {
        private async Task StoryBoard()
        {
            await Task.Run(() => 
            {
                // Step 1. Making CSV
                List<sp_change_depResult> ListChange = MakeCSV();
                log.Info("1. CSV made.");

                // 如果有錯誤,離開程式
                if (ListChange is null) return;
                log.Info("2. ListChange is not null.");

                // 沒有資料就離開程式
                if (ListChange.Count == 0) return;
                log.Info("3. ListChange count > 0.");

                // 顯示共幾筆, 目前第幾筆
                this.Dispatcher.Invoke((Action)(() =>
                {
                    this.total_n.Content = ListChange.Count.ToString();
                    this.current_n.Content = "0";
                }));
                log.Info("4. Number displayed.");

                // Step 2. Environment
                Environment();
                log.Info("5. Environment created.");

                AutoItX.WinWaitActive("看診清單");
                AutoItX.WinActivate("看診清單"); //要在這裡activate, 為何沒有用?

                // Step 2.5 wait for responsiveness
                WaitForResponsiveness("看診清單", "[NAME:cmbVist]");

                // Step 3. 開始讀取
                #region Execute change department
                int idx = 0;
                foreach (var c in ListChange)
                {
                    // Step 3.1 比較異同
                    string strD_n = c.o.Substring(0, 8); // 目標的日期
                    string strV_n = c.o.Substring(8, 1); // 目標的上下午
                    string strR_n = c.o.Substring(9, 2); // 目標的診間
                    log.Info($"6. case {idx + 1}.");

                    do
                    {
                        // 跳出機制, 如果按下停止, 就跳出整個程式, 計算已經完成的數量, 顯示出來
                        // 20210722 還沒有寫
                        if (_stopflag)
                        {
                            Summerize(idx);
                            return;
                        }
                    } while (!CompareV(strD_n, strV_n, strR_n));
                    log.Info($"7. after date, vist, room number set.");

                    // 號碼, 科別
                    string strNr_n = c.o.Substring(11, 3);  // 目標的看診號
                    string strDEP_n = c.o.Substring(14, 2); // 目標的科別

                    // Step 3.2 按下新的號碼;
                    // Seq NO
                    // 輸入診號
                    AutoItX.ControlSend("看診清單", "", "[NAME:txbSqno]", strNr_n);
                    log.Info($"8. Enter {strNr_n}.");

                    // 按鈕
                    AutoItX.ControlClick("看診清單", "", "[NAME:btnGo]");
                    log.Info($"9. Press GO button. [NAME:btnGo].");

                    // 如果進不去問診畫面呢? 沒有這個號碼
                    // ????

                    // Step 3.3 進入問診畫面
                    AutoItX.WinWaitActive("問診畫面");
                    WaitForResponsiveness("問診畫面", "[NAME:cmbDept]");
                    log.Info($"stop 1");

                    do
                    {
                        log.Info($"stop 2");
                        // 跳出機制, 如果按下停止, 就跳出整個程式, 計算已經完成的數量, 顯示出來
                        // 20210722 還沒有寫
                        if (_stopflag)
                        {
                            Summerize(idx);
                            return;
                        }
                    } while (!CompareP(strNr_n, strDEP_n));
                    log.Info($"stop 3");

                    // Step 3.4 收尾動作, 最後回到"看診清單"
                    Finishing();
                    log.Info($"stop 4");

                    idx++;
                    // Step 3.5 顯示共幾筆, 目前第幾筆
                    this.Dispatcher.Invoke((Action)(() =>
                    {
                        this.current_n.Content = idx.ToString();
                    }));
                }

                // Step 4 中斷或結束, 應自成一個sub
                Summerize(idx);
                return;

                #endregion Execute change department
            });
        }
    }
}