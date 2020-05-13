using ClinicSysteMc.ViewModel.Converters;
using Hardcodet.Wpf.TaskbarNotification;
using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System;
using System.Windows.Input;

namespace ClinicSysteMc.ViewModel.Commands
{
    internal class InputFile : ICommand
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private readonly TaskbarIcon tb = new TaskbarIcon();

        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object parameter)
        {
            return true;
        }

        public void Execute(object parameter)
        {
            // inputbox

            #region 讀取檔案路徑

            // 讀取要輸入的位置
            string loadpath;
            // 從杏翔病患資料輸入, 只有一種xml格式
            // 依照parameter, 不同來源: 申報匯入, 門診, 病患, 醫令, 檢驗, 指向不同方向
            OpenFileDialog oFDialog = new OpenFileDialog();
            switch ((string)parameter)
            {
                case "門診":
                    oFDialog.Filter = "xml|*.xml";
                    if (oFDialog.ShowDialog() != true) return;
                    loadpath = oFDialog.FileName;

                    OPDconvert o = new OPDconvert(loadpath);
                    o.Transform();

                    Logging.Record_admin("add opd", "匯入門診檔案 Manual");

                    break;

                case "病患":
                    oFDialog.Filter = "xlsx|*.xlsx";
                    if (oFDialog.ShowDialog() != true) return;
                    loadpath = oFDialog.FileName;

                    Microsoft.Office.Interop.Excel.Application myExcel = new Microsoft.Office.Interop.Excel.Application();
                    Workbook wb = myExcel.Workbooks.Open(loadpath);
                    Worksheet ws = wb.ActiveSheet;
                    // 丟出的是一個object [,]
                    PTconvert p = new PTconvert(ws.UsedRange.Value2);
                    p.Transform();

                    Logging.Record_admin("add/change patients", "加入/修改病患資料 Manual");

                    break;

                default:
                    break;
            }

            #endregion 讀取檔案路徑
        }
    }
}