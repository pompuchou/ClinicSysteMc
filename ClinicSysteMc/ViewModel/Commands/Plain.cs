using ClinicSysteMc.Model;
using ClinicSysteMc.ViewModel.Converters;
using System;
using System.Windows.Input;

namespace ClinicSysteMc.ViewModel.Commands
{
    internal class Plain : ICommand
    {
        private readonly MainVM _mainVM;
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

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

        public async void Execute(object parameter)
        {
            Progress<ProgressReportModel> progress = new Progress<ProgressReportModel>();
            progress.ProgressChanged += ReportProgress;

            if ((string)parameter == "病患(自動)")
            {
                log.Info($"    Button PT_auto pressed.");
                PTconvert_auto p = new PTconvert_auto();
                await p.Convert(progress);
            }

            if ((string)parameter == "醫令(自動)")
            {
                log.Info($"    Button ODR_auto pressed.");
                ODRconvert_auto odr = new ODRconvert_auto();
                await odr.Convert(progress);
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
