using ClinicSysteMc.Model;
using System;
using System.Linq;
using System.Windows.Input;

namespace ClinicSysteMc.ViewModel.Commands
{
    internal class LABmatch : ICommand
    {
        private readonly LabMatchVM _lmVM;

        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        public LABmatch(LabMatchVM LABmatchVM)
        {
            _lmVM = LABmatchVM;
        }

        public bool CanExecute(object parameter)
        {
            return true;
        }

        public void Execute(object parameter)
        {
            #region 進行配對

            CSDataContext dc = new CSDataContext();
            // 20190615 tbl_lab_record連結tbl_opd_order
            var q = from cs in dc.sp_match_lab(int.Parse(_lmVM.StrFrom), int.Parse(_lmVM.StrTO)).AsEnumerable()
                    select cs;
            int n = q.First().rows_affected;
            Logging.Record_admin("檢驗檔配對", $"{n}筆配對成功");

            #endregion 進行配對

            _lmVM.Refresh_Data();
        }
    }
}