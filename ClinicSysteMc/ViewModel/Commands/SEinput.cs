using System;
using ClinicSysteMc.ViewModel.Dialog;
using System.Windows.Input;

namespace ClinicSysteMc.ViewModel.Commands
{
    internal class SEinput : ICommand
    {
        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        public bool CanExecute(object parameter)
        {
            return true;
        }

        public void Execute(object parameter)
        {
            #region ASK for begin date, end date

            DateTime d = DateTime.Now;
            if ((string)parameter == "制檢驗上傳") d = DateTime.Now.AddMonths(-1);

            var dlg = new SEdialog(DateTime.Now.AddDays(-10), DateTime.Now);
            dlg.ShowDialog();

            if (dlg.DialogResult == false) return;

            #endregion ASK for begin date, end date
        }
    }
}