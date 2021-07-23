using AutoIt;
using System;

namespace ClinicSysteMc.ViewModel.Converters
{
    public partial class Dash
    {
        internal bool CompareV(string strD_n, string strV_n, string strR_n)
        {
            try
            {
                // 比較看診清單, 核對日期, 上下午, 診間
                string tmpD_o = AutoItX.ControlGetText("看診清單", "", "[NAME:dtpSDate]"); // 杏翔系統的日期
                string tmpV_o = AutoItX.ControlGetText("看診清單", "", "[NAME:cmbVist]"); // 杏翔系統的上下午
                string tmpR_o = AutoItX.ControlGetText("看診清單", "", "[NAME:cmbRmno]"); // 杏翔系統的診間            
                string strD_o = DateTime.Parse(tmpD_o).ToString("yyyyMMdd"); // 杏翔系統的日期
                string strV_o = tmpV_o.Substring(0, 1); // 杏翔系統的上下午, 前一碼
                string strR_o = tmpR_o.Substring(0, 2); // 杏翔系統的診間, 前兩碼

                bool changed = false;

                // 先檢查是否換日, 如果有就換到新日期;
                if (strD_n != strD_o)
                {
                    // 製造3個AutoIT VB程式, 1. changeDP_DATE, 針對"看診清單", [NAME:dtpSDate]
                    // 一個參數, 格式YYYYMMDD
                    AutoItX.Run($"C:\\vpn\\exe\\changeDP_DATE.exe {strD_n}", @"C:\vpn\exe\");
                    changed = true;
                    log.Info($"Change Date from {strD_o} to {strD_n}.");
                    // 等待可以反應
                    AutoItX.Sleep(500);
                    //WaitForResponsiveness("看診清單", "[NAME:cmbVist]");
                }

                // 再檢查是否換午別, 如果有就換到新的午別
                if (strV_n != strV_o)
                {
                    // 製造3個AutoIT VB程式, 2. changeDP_VIST, 針對"看診清單", [NAME:cmbVist]
                    // 一個參數, 格式V
                    AutoItX.Run($"C:\\vpn\\exe\\changeDP_VIST.exe {strV_n}", @"C:\vpn\exe\");
                    log.Info($"Change Vist from {strV_o} to {strV_n}. Then sleep 1000ms.");
                    changed = true;
                    // 等待可以反應
                    AutoItX.Sleep(500);
                    //WaitForResponsiveness("看診清單", "[NAME:cmbVist]");
                }

                // 最後檢查是否換診間, 如果有就換到新的診間
                if (strR_n != strR_o)
                {
                    // 製造3個AutoIT VB程式, 3. changeDP_ROOM, 針對"看診清單", [NAME:cmbRmno]
                    // 一個參數, 格式RR
                    AutoItX.Run($"C:\\vpn\\exe\\changeDP_ROOM.exe {strR_n}", @"C:\vpn\exe\");
                    log.Info($"Change Date from {strR_o} to {strR_n}. Then sleep 1000ms.");
                    changed = true;
                    // 等待可以反應
                    AutoItX.Sleep(500);
                    //WaitForResponsiveness("看診清單", "[NAME:cmbVist]");
                }

                if (changed)
                {
                    AutoItX.ControlClick("看診清單", "", "[NAME:btnRefresh]");
                    log.Info($"Press Refresh button. [NAME:btnRefresh].");
                    AutoItX.WinWaitActive("看診清單");
                    // 等待可以反應
                    WaitForResponsiveness("看診清單", "[NAME:cmbVist]");
                }

                // 再檢查一次
                tmpD_o = AutoItX.ControlGetText("看診清單", "", "[NAME:dtpSDate]"); // 杏翔系統的日期
                tmpV_o = AutoItX.ControlGetText("看診清單", "", "[NAME:cmbVist]"); // 杏翔系統的上下午
                tmpR_o = AutoItX.ControlGetText("看診清單", "", "[NAME:cmbRmno]"); // 杏翔系統的診間            
                strD_o = DateTime.Parse(tmpD_o).ToString("yyyyMMdd"); // 杏翔系統的日期
                strV_o = tmpV_o.Substring(0, 1); // 杏翔系統的上下午, 前一碼
                strR_o = tmpR_o.Substring(0, 2); // 杏翔系統的診間, 前兩碼

                if (strD_n == strD_o && strV_n == strV_o && strR_n == strR_o)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                string o = $"Something wrong in CompareV: {ex.Message}";
                Logging.Record_error(o);
                log.Error(o);
                return false;
            }


            // 如此可確保 日期, 上下午, 診間正確
        }
    }
}
