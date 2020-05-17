using System;
using System.Windows.Input;

namespace ClinicSysteMc.ViewModel.Commands
{
    class DATArefreshi : ICommand
    {
        private readonly InfoVM _iVM;

        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        public DATArefreshi(InfoVM IVM)
        {
            _iVM = IVM;
        }

        public bool CanExecute(object parameter)
        {
            return true;
        }

        public void Execute(object parameter)
        {
            // 20200518 完成工作後可以更新資料
            _iVM.Refresh_Data();
        }
    }
}
