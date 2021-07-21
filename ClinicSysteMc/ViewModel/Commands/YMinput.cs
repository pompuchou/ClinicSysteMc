using ClinicSysteMc.ViewModel.Converters;
using ClinicSysteMc.ViewModel.Dialog;
using System;
using System.Windows;
using System.Windows.Input;

namespace ClinicSysteMc.ViewModel.Commands
{
    /// <summary>
    /// 這是輸入幾種需要先詢問年月的外部檔案
    /// 先詢問年月的, YM
    /// </summary>
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
                Owner = Application.Current.MainWindow,
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

            if ((string)parameter == "印花稅")
            {
                // 將YM轉換成begin_date, end_date
                // 例如: 10801, 10802 => 2019/01/01 ~ 2019/02/28
                // StrYM: Year: int.Parse(StrYM.Substring(0, 3)) + 1911
                int Y = int.Parse(dlg.StrYM.Substring(0, 3)) + 1911;
                int M = int.Parse(dlg.StrYM.Substring(3, 2));
                if (M % 2 == 0) M--;
                DateTime bd = DateTime.Parse($"{Y}/{M}/1");
                DateTime ed = bd.AddMonths(2).AddMilliseconds(-1);
                if (DateTime.Now < ed) return;

                string StrYM = $"{Y-1911}{(bd.Month + 100).ToString().Substring(1)}{(ed.Month + 100).ToString().Substring(1)}";
                STAMPimport s = new STAMPimport(StrYM, bd, ed);
                s.Import();
            }
            // 20200518 完成工作後可以更新資料
            _mainVM.Refresh_Data();
        }
    }
}