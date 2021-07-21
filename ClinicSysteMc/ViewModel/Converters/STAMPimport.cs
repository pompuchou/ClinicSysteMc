using AutoIt;
using ClinicSysteMc.Model;
using Hardcodet.Wpf.TaskbarNotification;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Serialization;

namespace ClinicSysteMc.ViewModel.Converters
{
    internal class STAMPimport
    {
        // 20200603 created
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private readonly TaskbarIcon tb = new TaskbarIcon();
        private readonly string _strYM;
        private readonly DateTime _bd;
        private readonly DateTime _ed;

        public STAMPimport(string YM, DateTime BD, DateTime ED)
        {
            _strYM = YM;
            _bd = BD;
            _ed = ED;
        }

        public void Import()
        {
            log.Info("  Begin Import.");

            #region Environment

            try
            {
                // 印花稅總繳明細表
                log.Info("  Check 印花稅總繳明細表.");
                if (AutoItX.WinExists("印花稅總繳明細表") == 1) //如果直接存在就直接叫用
                {
                    AutoItX.WinActivate("印花稅總繳明細表");
                    log.Info("  印花稅總繳明細表 exists.");
                }
                else
                {
                    log.Info("  印花稅總繳明細表 doesn't exist.");
                    Thesis.LogIN();
                    // 打開"印花稅總繳明細表", 找不到control,只好用mouse去按
                    AutoItX.WinActivate("杏雲雲端醫療服務");
                    // 先maximize
                    AutoItX.WinSetState("杏雲雲端醫療服務", "", 3);  //0 close; 1 @SW_RESTORE; 2 @SW_MINIMIZE; 3 @SW_MAXIMIZE
                    AutoItX.MouseMove(280, 280);
                    AutoItX.MouseClick();
                    AutoItX.Sleep(500);
                    AutoItX.ControlClick("杏雲雲端醫療服務", "", "[NAME:btnStampTax]");
                    log.Info("  處方清單 opened.");
                    AutoItX.Sleep(10000);
                }
            }
            catch (Exception ex)
            {
                Logging.Record_error(ex.Message);
                log.Error(ex.Message);
                return;
            }


            #endregion
        }
    }
}
