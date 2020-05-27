using AutoIt;
using ClinicSysteMc.Model;
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Controls;

namespace ClinicSysteMc.ViewModel.Converters
{
    internal class PTconvert_auto
    {
        // 20190610 created
        // 目的是自動匯入病患資料
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public void Convert(Progress<ProgressReportModel> progress)
        {
            Microsoft.Office.Interop.Excel.Application MyExcel = new Microsoft.Office.Interop.Excel.Application();

            #region Environment

            // 殺掉所有的EXCEL
            foreach (Process p in Process.GetProcessesByName("EXCEL"))
            {
                p.Kill();
            }
            // 營造環境
            Process[] isCust = Process.GetProcessesByName("THCustomerFilter");   // 處方清單
            if (isCust.Length == 0)    // 如果沒有打開
            {
                // 測試"看診清單"是否有打開
                Thesis.LogIN();
                // 打開"各類特殊 追蹤與紀錄查詢"
                AutoItX.Run(@"C:\Program Files (x86)\THESE\杏雲醫療資訊系統\THCustomerFilter.exe", @"C:\Program Files (x86)\THESE\杏雲醫療資訊系統\");
                System.Threading.Thread.Sleep(2000);
            }
            // 準備好
            AutoItX.WinActivate("各類特殊 追蹤與紀錄查詢");
            AutoItX.WinWaitActive("各類特殊 追蹤與紀錄查詢");
            AutoItX.ControlClick("各類特殊 追蹤與紀錄查詢", "", "[NAME:chk允許完整筆數呈現]");
            AutoItX.Sleep(1000);
            // [NAME:btn病歷號查詢]
            AutoItX.ControlClick("各類特殊 追蹤與紀錄查詢", "", "[NAME:btn病歷號查詢]");
            AutoItX.Sleep(1000);
            // 病歷號查詢
            // [NAME:TextBox]
            // [NAME:OKButton]
            AutoItX.ControlSend("病歷號查詢", "", "[NAME:TextBox]", "0000000001~9999999999");
            AutoItX.ControlClick("病歷號查詢", "", "[NAME:OKButton]");
            log.Info("Click inquiry.");
            //AutoItX.WinWaitActive("各類特殊 追蹤與紀錄查詢");
            //AutoItX.WinWaitActive("各類特殊 追蹤與紀錄查詢");
            //log.Info("WinWait ends.");
            //AutoItX.Sleep(20000);

            // 20190610 模仿昨天成功的經驗
            // [NAME:btn匯出EXCEL]
            // aut.WinWait("[dlgPrintMethodAsk]",, 1000)
            // aut.ControlClick("[dlgPrintMethodAsk]", "", "[NAME:OK_Button]")

            log.Info("Start to Click.");
            do
            {
                AutoItX.ControlClick("各類特殊 追蹤與紀錄查詢", "", "[NAME:btn匯出EXCEL]");
                AutoItX.Sleep(10000);
                log.Info("Click export one time.");
            } while (Process.GetProcessesByName("EXCEL").Length == 0);
            log.Info("Excel exists now.");

            // aut.Sleep(10000), 用等的,等10秒大多有效,但不能保證,且也許不用10秒,這樣就浪費了, 應該要個別化
            // 好在發現visibility可以有效等到整個檔案製作完成
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
            string temp_filepath = @"C:\vpn\pt";
            // 20190609 因為不小心多一個空格, 搞了好久除錯, 很辛苦啊
            // System.Runtime.InteropServices.COMException '發生例外狀況於 HRESULT: 0x800A03EC'
            // 存放目錄,不存在就要建立一個
            if (!System.IO.Directory.Exists(temp_filepath))
            {
                System.IO.Directory.CreateDirectory(temp_filepath);
            }
            // 自動產生名字
            string temp_file = $"\\pt_{DateTime.Now.Year}{(DateTime.Now.Month + 100).ToString().Substring(1)}";
            temp_file += $"{(DateTime.Now.Day + 100).ToString().Substring(1)}_{DateTime.Now.TimeOfDay}";
            temp_file = temp_file.Replace(":", "").Replace(".", "");
            temp_filepath += $"{temp_file}.xlsx";
            // wb.SaveAs(temp_filepath, Excel.XlFileFormat.xlCSV, vbNull, vbNull, False, False, Excel.XlSaveAsAccessMode.xlNoChange, vbNull, vbNull, vbNull, vbNull, vbNull)
            Microsoft.Office.Interop.Excel.Workbook wb = MyExcel.ActiveWorkbook;
            wb.SaveAs(temp_filepath, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook);
            Microsoft.Office.Interop.Excel.Worksheet ws = wb.ActiveSheet;

            #endregion Saving XLSX files

            // 丟出的是一個object [,]
            PTconvert pt = new PTconvert(ws.UsedRange.Value2);
            pt.Transform(progress);


            #region Ending

            // 殺掉所有的EXCEL
            foreach (Process p in Process.GetProcessesByName("EXCEL"))
            {
                p.Kill();
            }
            AutoItX.WinClose("各類特殊 追蹤與紀錄查詢");

            #endregion Ending
        }
    }
}