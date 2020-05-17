using System;
using System.Windows.Input;

namespace ClinicSysteMc.ViewModel.Commands
{
    class DATArefresh : ICommand
    {
        private readonly MainVM _mainVM;

        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        public DATArefresh(MainVM MVM)
        {
            _mainVM = MVM;
        }

        public bool CanExecute(object parameter)
        {
            return true;
        }

        public void Execute(object parameter)
        {
            // 20200518 完成工作後可以更新資料
            _mainVM.Refresh_Data();
        }
    }
}
