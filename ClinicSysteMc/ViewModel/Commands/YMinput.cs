using ClinicSysteMc.ViewModel.Converters;
using ClinicSysteMc.ViewModel.Dialog;
using System;
using System.Windows.Input;

namespace ClinicSysteMc.ViewModel.Commands
{
    internal class YMinput : ICommand
    {
        private readonly MainVM _mainVM;

        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        public YMinput(MainVM MVM)
        {
            _mainVM = MVM;
        }

        public bool CanExecute(object parameter)
        {
            return true;
        }

        public void Execute(object parameter)
        {
            #region ASK for YM

            DateTime d = DateTime.Now;
            if ((string)parameter == "制檢驗上傳") d = DateTime.Now.AddMonths(-1);

            var dlg = new YMdialog()
            {
                StrYM = $"{d.Year - 1911}{(d.Month + 100).ToString().Substring(1)}"
            };
            dlg.ShowDialog();

            if (dlg.DialogResult == false) return;

            #endregion ASK for YM

            if ((string)parameter == "制檢驗上傳")
            {
                LABXMLbuild x = new LABXMLbuild(dlg.StrYM);
                x.Build();
            }

            if ((string)parameter == "調整科別")
            {
                DEPchange c = new DEPchange(dlg.StrYM);
                c.Change();
            }

            // 20200518 完成工作後可以更新資料
            _mainVM.Refresh_Data();
        }
    }
}