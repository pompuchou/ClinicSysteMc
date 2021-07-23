using AutoIt;
using System;

namespace ClinicSysteMc.ViewModel.Converters
{
    public partial class Dash
    {
        internal bool CompareP(string strNr_n, string strDEP_n)
        {
            try
            {
                // 比較看診號, 科別
                // [NAME: txbCASENO]
                string tmpNr_o = AutoItX.ControlGetText("問診畫面", "", "[NAME:txbCASENO]"); // 杏翔系統的看診號
                string tmpDEP_o = AutoItX.ControlGetText("問診畫面", "", "[NAME:cmbDept]"); // 杏翔系統的科別
                string strNr_o = tmpNr_o.Substring(11, 3); // 杏翔系統的看診號
                string strDEP_o = tmpDEP_o.Substring(0, 2); // 杏翔系統的科別, 前二碼

                if (strNr_n == strNr_o)
                {
                    // 製造一個AutoIT VB程式, changeDP_DEP, 針對"問診畫面"[NAME:cmbDept]
                    // 一個參數, 格式DD
                    AutoItX.Run($"C:\\vpn\\exe\\changeDP_DEP.exe {strDEP_n}", @"C:\vpn\exe\");
                }

                tmpNr_o = AutoItX.ControlGetText("問診畫面", "", "[NAME:txbCASENO]"); // 杏翔系統的看診號
                tmpDEP_o = AutoItX.ControlGetText("問診畫面", "", "[NAME:cmbDept]"); // 杏翔系統的科別
                strNr_o = tmpNr_o.Substring(11, 3); // 杏翔系統的看診號
                strDEP_o = tmpDEP_o.Substring(0, 2); // 杏翔系統的科別, 前二碼

                if (strNr_n == strNr_o && strDEP_n == strDEP_o)
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
                string o = $"Something wrong in CompareP: {ex.Message}";
                Logging.Record_error(o);
                log.Error(o);
                return false;
            }
        }
    }
}
