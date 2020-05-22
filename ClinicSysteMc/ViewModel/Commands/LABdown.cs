using System;
using System.Runtime.InteropServices;
using System.Windows.Input;

namespace ClinicSysteMc.ViewModel.Commands
{
    internal class LABdown : ICommand
    {

        private readonly LabMatchVM _lmVM;

        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        public LABdown(LabMatchVM LABmatchVM)
        {
            _lmVM = LABmatchVM;
        }

        public bool CanExecute(object parameter)
        {
            if (int.TryParse(_lmVM.StrTO, out int iFrom))
            {
                if (iFrom > 0) return true;
            }
            return false;
        }

        public void Execute(object parameter)
        {
            _lmVM.StrTO = (int.Parse(_lmVM.StrTO) - 1).ToString();
        }
    }
}
