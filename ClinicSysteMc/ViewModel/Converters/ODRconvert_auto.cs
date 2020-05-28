using AutoIt;
using ClinicSysteMc.Model;
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace ClinicSysteMc.ViewModel.Converters
{
    internal class ODRconvert_auto
    {
        // 20190610 created
        // 目的是自動匯入批價項目資料
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public async Task Convert(Progress<ProgressReportModel> progress)
        {
            Microsoft.Office.Interop.Excel.Application MyExcel = new Microsoft.Office.Interop.Excel.Application();

            #region Environment

            // 殺掉所有的EXCEL
            foreach (Process p in Process.GetProcessesByName("EXCEL"))
            {
                p.Kill();
            }

            // 營造環境
            // 各類資料維護
            // 計價標準維護      (這就是我們的標的)
            if (AutoItX.WinExists("計價標準檔維護") == 1) //如果直接存在就直接叫用
            {
                AutoItX.WinActivate("計價標準檔維護");
            }
            else
            {
                if (AutoItX.WinExists("各類資料維護") == 1)
                {
                    AutoItX.WinActivate("各類資料維護");
                }
                else
                {
                    Thesis.LogIN();
                    // 從"杏雲雲端醫療服務"叫用"各類資料維護"
                    // 打開"處方清單", 找不到control,只好用mouse去按
                    AutoItX.WinActivate("杏雲雲端醫療服務");
                    // 先maximize
                    AutoItX.WinSetState("杏雲雲端醫療服務", "", 3);  //0 close; 1 @SW_RESTORE; 2 @SW_MINIMIZE; 3 @SW_MAXIMIZE
                    AutoItX.MouseClick("LEFT", AutoItX.WinGetPos("杏雲雲端醫療服務").X + 200, AutoItX.WinGetPos("杏雲雲端醫療服務").Y + 175);
                    AutoItX.Sleep(500);
                    AutoItX.ControlClick("杏雲雲端醫療服務", "", "[NAME:btnDBaseMaint]");
                }
                // 從"各類資料維護"叫用"計價標準檔維護"
                AutoItX.Sleep(10000);
                AutoItX.ControlSetText("各類資料維護", "", "[NAME:txbQuery]", "計價標準檔維護");
                // AutoItX.ControlSend("各類資料維護", "", "[NAME:txbQuery]", "計價標準檔維護")
                AutoItX.ControlClick("各類資料維護", "", "[NAME:btnQuery]");
                AutoItX.MouseClick("LEFT", AutoItX.WinGetPos("各類資料維護").X + 100, AutoItX.WinGetPos("各類資料維護").Y + 135, 2);
                log.Info("Click inquiry.");
                //AutoItX.Sleep(2000);
                //AutoItX.WinActivate("計價標準檔維護");
            }

            log.Info("Start to Click.");
            do
            {
                // 20200517 一直按到行為止
                // 20190610 模仿昨天成功的經驗
                // 打開EXCEL檔
                AutoItX.Send("{Alt}");
                AutoItX.Send("{Down}");
                AutoItX.Send("{Down}");
                AutoItX.Send("{Down}");
                AutoItX.Send("{Down}");
                AutoItX.Send("{Down}");
                AutoItX.Send("{Down}");
                // 20200319 修改程式,再往下一格
                AutoItX.Send("{Down}");
                AutoItX.Send("{Enter}");
                AutoItX.Sleep(10000);
            } while (Process.GetProcessesByName("EXCEL").Length == 0);
            log.Info("Excel exists now.");

            MyExcel = (Microsoft.Office.Interop.Excel.Application)Marshal.GetActiveObject("Excel.Application");
            do
            {
                AutoItX.Sleep(100);
            } while (!MyExcel.Visible);
            log.Info("Excel appears now.");

            #endregion Environment

            #region Saving XLSX files

            // ====================================================================================================================================
            // 製作自動檔名
            string temp_filepath = @"C:\vpn\odr";
            // 20190609 因為不小心多一個空格, 搞了好久除錯, 很辛苦啊
            // System.Runtime.InteropServices.COMException '發生例外狀況於 HRESULT: 0x800A03EC'
            // 存放目錄,不存在就要建立一個
            if (!System.IO.Directory.Exists(temp_filepath))
            {
                System.IO.Directory.CreateDirectory(temp_filepath);
            }
            // 自動產生名字
            string temp_file = $"\\odr_{DateTime.Now.Year}{(DateTime.Now.Month + 100).ToString().Substring(1)}";
            temp_file += $"{(DateTime.Now.Day + 100).ToString().Substring(1)}_{DateTime.Now.TimeOfDay}";
            temp_file = temp_file.Replace(":", "").Replace(".", "");
            temp_filepath += $"{temp_file}.xlsx";
            // wb.SaveAs(temp_filepath, Excel.XlFileFormat.xlCSV, vbNull, vbNull, False, False, Excel.XlSaveAsAccessMode.xlNoChange, vbNull, vbNull, vbNull, vbNull, vbNull)
            Microsoft.Office.Interop.Excel.Workbook wb = MyExcel.ActiveWorkbook;
            wb.SaveAs(temp_filepath, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook);
            Microsoft.Office.Interop.Excel.Worksheet ws = wb.ActiveSheet;

            #endregion Saving XLSX files

            // 丟出的是一個object [,]
            ODRconvert odr = new ODRconvert(ws.UsedRange.Value2);
            await odr.Transform(progress);

            #region Ending

            // 殺掉所有的EXCEL
            foreach (Process p in Process.GetProcessesByName("EXCEL"))
            {
                p.Kill();
            }
            AutoItX.WinClose("計價標準檔維護");
            AutoItX.WinClose("各類資料維護");

            #endregion Ending
        }
    }
}