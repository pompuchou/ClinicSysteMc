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

            if ((string)parameter == "匯入批價檔")
            {
                PIJIAconvert p = new PIJIAconvert(dlg.BeginDate, dlg.EndDate);
                p.Convert();
            }

            if ((string)parameter == "門診(自動)")
            {
                OPDconvert_auto o = new OPDconvert_auto(dlg.BeginDate, dlg.EndDate);
                o.Convert();
            }

            // 20200518 完成工作後可以更新資料
            _mainVM.Refresh_Data();
        }
    }
}