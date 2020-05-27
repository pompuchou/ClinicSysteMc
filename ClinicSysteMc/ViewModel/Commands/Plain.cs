﻿using ClinicSysteMc.Model;
using ClinicSysteMc.ViewModel.Converters;
using System;
using System.Windows.Input;

namespace ClinicSysteMc.ViewModel.Commands
{
    internal class Plain : ICommand
    {
        private readonly MainVM _mainVM;

        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        public Plain(MainVM MVM)
        {
            _mainVM = MVM;
        }

        public bool CanExecute(object parameter)
        {
            return true;
        }

        public void Execute(object parameter)
        {
            Progress<ProgressReportModel> progress = new Progress<ProgressReportModel>();
            progress.ProgressChanged += ReportProgress;

            if ((string)parameter == "病患(自動)")
            {
                PTconvert_auto p = new PTconvert_auto();
                p.Convert(progress);
            }

            if ((string)parameter == "醫令(自動)")
            {
                ODRconvert_auto odr = new ODRconvert_auto();
                odr.Convert(progress);
            }

            // 20200518 完成工作後可以更新資料
            _mainVM.Refresh_Data();
        }

        private void ReportProgress(object sender, ProgressReportModel e)
        {
            _mainVM.ProgressValue = e.PercentageComeplete;
        }

    }
}
