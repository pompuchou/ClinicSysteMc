using ClinicSysteMc.ViewModel.Converters;
using ClinicSysteMc.ViewModel.Dialog;
using System;
using System.Windows.Input;

namespace ClinicSysteMc.ViewModel.Commands
{
    // B stands for begin, E for end
    internal class BEinput : ICommand
    {
        private readonly MainVM _mainVM;
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        public BEinput(MainVM MVM)
        {
            _mainVM = MVM;
        }

        public bool CanExecute(object parameter)
        {
            return true;
        }

        public void Execute(object parameter)
        {
            log.Info("Begin Excution.");
            #region ASK for begin date, end date

            DateTime bd = DateTime.Parse($"{DateTime.Now.Year}/{DateTime.Now.Month}/1");
            DateTime ed = DateTime.Now;
            if ((string)parameter == "匯入批價檔")
            {
                bd = DateTime.Parse($"{DateTime.Now.Year}/{DateTime.Now.Month}/1").AddMonths(-1);
                ed = DateTime.Parse($"{DateTime.Now.Year}/{DateTime.Now.Month}/1").AddSeconds(-1);
            }

            var dlg = new BEdialog(bd, ed);
            dlg.ShowDialog();

            if (dlg.DialogResult == false) return;

            #endregion ASK for begin date, end date
            log.Info($"  Begin: {dlg.BeginDate}; End: {dlg.EndDate}");

            if ((string)parameter == "匯入批價檔")
            {
                PIJIAconvert p = new PIJIAconvert(dlg.BeginDate, dlg.EndDate);
                p.Convert();
            }

            if ((string)parameter == "門診(自動)")
            {
                log.Info("  Going to OPDconvert_auto");
                OPDconvert_auto o = new OPDconvert_auto(dlg.BeginDate, dlg.EndDate);
                o.Convert();
            }

            // 20200518 完成工作後可以更新資料
            log.Info("  Refresh Data.");
            _mainVM.Refresh_Data();

            log.Info("End Excution.");
        }
    }
}