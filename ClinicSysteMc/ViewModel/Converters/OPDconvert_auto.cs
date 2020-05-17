using AutoIt;
using System;
using System.Threading;

namespace ClinicSysteMc.ViewModel.Converters
{
    internal class OPDconvert_auto
    {
        // 20190608 created
        // 20190608 add try, record_adm, record_err
        // 目的是自動惠入門診資料
        private readonly DateTime _begindate;

        private readonly DateTime _enddate;

        public OPDconvert_auto(DateTime begindate, DateTime enddate)
        {
            _begindate = begindate;
            _enddate = enddate;
        }

        public void Convert()
        {
            string output;

            #region Environment

            try
            {
                // 營造環境
                if (AutoItX.WinExists("處方清單") == 1) //如果直接存在就直接叫用
                {
                    AutoItX.WinActivate("處方清單");
                }
                else
                {
                    Thesis.LogIN();
                    // 打開"處方清單", 找不到control,只好用mouse去按
                    AutoItX.WinActivate("杏雲雲端醫療服務");
                    // 先maximize
                    AutoItX.WinSetState("杏雲雲端醫療服務", "", 3);  //0 close; 1 @SW_RESTORE; 2 @SW_MINIMIZE; 3 @SW_MAXIMIZE
                    AutoItX.MouseMove(280, 280);
                    AutoItX.MouseClick();
                    AutoItX.Sleep(500);
                    AutoItX.ControlClick("杏雲雲端醫療服務", "", "[NAME:btnPrescription]");
                    Thread.Sleep(10000);
                }

                // 打開備份
                AutoItX.WinWaitActive("處方清單");
                AutoItX.ControlClick("處方清單", "", "[NAME:btnBackup]");
                AutoItX.WinActivate("處方清單備份選項");
                AutoItX.WinWaitActive("處方清單備份選項");
                AutoItX.ControlClick("處方清單備份選項", "", "[NAME:txbBackupPath]", "LEFT", 2);
                AutoItX.Send("{Tab}");
                AutoItX.Send("{Tab}");
                AutoItX.Send("{Enter}"); // first choice Desktop
                                         // 這裡的等待很重要, 太短來不及讀, 500可以, 100 不行, 200 一半一半, 250 100%
                AutoItX.Sleep(300);
                // 尋找XML, 若有就刪除
                output = AutoItX.ControlGetText("處方清單備份選項", "", "[NAME:txbBackupPath]");
                output += $"\\{_begindate.Year}\\{_begindate.Year}{(_begindate.Month + 100).ToString().Substring(1)}.xml";
                if (System.IO.File.Exists(output)) System.IO.File.Delete(output);
                // AutoItX.ControlSend("處方清單備份選項", "", "[NAME:txbBackupPath]", "C:\vpn")
            }
            catch (Exception ex)
            {
                output = string.Empty;
                Logging.Record_error(ex.Message);
            }

            #endregion Environment

            #region Producing XML

            string BeginDate = $"{_begindate.Year}{(_begindate.Month + 100).ToString().Substring(1)}{(_begindate.Day + 100).ToString().Substring(1)}";
            string EndDate = $"{_enddate.Year}{(_enddate.Month + 100).ToString().Substring(1)}{(_enddate.Day + 100).ToString().Substring(1)}";
            string Execution = $"C:\\vpn\\exe\\changePresDTP.exe {BeginDate}{EndDate}";
            AutoItX.Run(Execution, @"C:\vpn\exe\");
            // 檢查XML做好了嗎?
            do
            {
                Thread.Sleep(100);
            } while (!System.IO.File.Exists(output));
            // XML好了就把頁面關掉
            AutoItX.Sleep(500);
            AutoItX.ControlClick("處方清單備份選項", "", "[NAME:Cancel_Button]");
            AutoItX.Sleep(100);
            AutoItX.ControlClick("處方清單", "", "[NAME:BtnEXIT]");

            #endregion Producing XML

            OPDconvert o = new OPDconvert(output);
            o.Transform();
        }
    }
}