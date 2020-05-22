using System;
using System.Windows.Input;

namespace ClinicSysteMc.ViewModel.Commands
{
    internal class LABup : ICommand
    {
        private readonly LabMatchVM _lmVM;

        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        public LABup(LabMatchVM LABmatchVM)
        {
            _lmVM = LABmatchVM;
        }

        public bool CanExecute(object parameter)
        {
            if (int.TryParse(_lmVM.StrTO, out _)) return true;
            return false;
        }

        public void Execute(object parameter)
        {
            _lmVM.StrTO = (int.Parse(_lmVM.StrTO) + 1).ToString();
        }

    }
}
