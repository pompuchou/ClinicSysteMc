﻿using ClinicSysteMc.ViewModel.Converters;
using System;
using System.Windows.Input;

namespace ClinicSysteMc.ViewModel.Commands
{
    internal class Plain : ICommand
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
            if ((string)parameter == "病患(自動)")
            {
                PTconvert_auto p = new PTconvert_auto();
                p.Convert();
            }

            if ((string)parameter == "醫令(自動)")
            {
                ODRconvert_auto odr = new ODRconvert_auto();
                odr.Convert();
            }
        }
    }
}