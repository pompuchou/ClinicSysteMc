using ClinicSysteMc.Model;
using ClinicSysteMc.ViewModel.Converters;
using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System;
using System.Windows.Input;

namespace ClinicSysteMc.ViewModel.Commands
{
    internal class FILEinput : ICommand
    {
        private readonly MainVM _mainVM;
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public event EventHandler CanExecuteChanged
        { 
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        public FILEinput(MainVM MVM)
        {
            _mainVM = MVM;
        }

        public bool CanExecute(object parameter)
        {
            return true;
        }

        public async void Execute(object parameter)
        {
            // inputbox

            #region 讀取檔案路徑

            // 讀取要輸入的位置
            string loadpath;
            Progress<ProgressReportModel> progress = new Progress<ProgressReportModel>();
            progress.ProgressChanged += ReportProgress;
            // 從杏翔病患資料輸入, 只有一種xml格式
            // 依照parameter, 不同來源: 申報匯入, 門診, 病患, 醫令, 檢驗, 指向不同方向
            OpenFileDialog oFDialog = new OpenFileDialog();
            switch ((string)parameter)
            {
                case "門診":
                    oFDialog.Filter = "xml|*.xml";
                    if (oFDialog.ShowDialog() != true) return;
                    loadpath = oFDialog.FileName;
                    log.Info($"    File: [{loadpath}] is being loaded.");

                    OPDconvert o = new OPDconvert(loadpath);
                    o.Transform();

                    Logging.Record_admin("add opd", "匯入門診檔案 Manual");

                    break;

                case "病患":
                    oFDialog.Filter = "xlsx|*.xlsx";
                    if (oFDialog.ShowDialog() != true) return;
                    loadpath = oFDialog.FileName;
                    log.Info($"    File: [{loadpath}] is being loaded.");

                    Microsoft.Office.Interop.Excel.Application myExcel = new Microsoft.Office.Interop.Excel.Application();
                    Workbook wb = myExcel.Workbooks.Open(loadpath);
                    Worksheet ws = wb.ActiveSheet;
                    // 丟出的是一個object [,]
                    PTconvert p = new PTconvert(ws.UsedRange.Value2);

                    await p.Transform(progress);

                    Logging.Record_admin("add/change patients", "加入/修改病患資料 Manual");

                    break;

                case "醫令":
                    oFDialog.Filter = "xlsx|*.xlsx";
                    if (oFDialog.ShowDialog() != true) return;
                    loadpath = oFDialog.FileName;
                    log.Info($"    File: [{loadpath}] is being loaded.");

                    Microsoft.Office.Interop.Excel.Application myExcel2 = new Microsoft.Office.Interop.Excel.Application();
                    Workbook wb2 = myExcel2.Workbooks.Open(loadpath);
                    Worksheet ws2 = wb2.ActiveSheet;
                    // 丟出的是一個object [,]
                    ODRconvert odr = new ODRconvert(ws2.UsedRange.Value2);
                    await odr.Transform(progress);

                    Logging.Record_admin("add/change order", "加入/修改醫令資料 Manual");

                    break;

                case "申報匯入":
                    oFDialog.Filter = "健保申報檔|TOTFA.xml";
                    if (oFDialog.ShowDialog() != true) return;
                    loadpath = oFDialog.FileName;
                    log.Info($"    File: [{loadpath}] is being loaded.");

                    TOTconvert tot = new TOTconvert(loadpath);
                    tot.Transform();

                    Logging.Record_admin("import xml", "匯入健保申報檔 Manual");

                    break;

                case "檢驗":
                    oFDialog.Filter = "xls|*.xls";
                    if (oFDialog.ShowDialog() != true) return;
                    loadpath = oFDialog.FileName;
                    log.Info($"    File: [{loadpath}] is being loaded.");

                    Microsoft.Office.Interop.Excel.Application myExcel3 = new Microsoft.Office.Interop.Excel.Application();
                    Workbook wb3 = myExcel3.Workbooks.Open(loadpath);
                    Worksheet ws3 = wb3.ActiveSheet;
                    // 丟出的是一個object [,]
                    LABconvert lab = new LABconvert(ws3.UsedRange.Value2);
                    lab.Transform();

                    Logging.Record_admin("add lab data", "加入檢驗資料 Manual");

                    break;

                case "健保匯入":
                    oFDialog.Filter = "健保藥物檔案|*.b5";
                    if (oFDialog.ShowDialog() != true) return;
                    loadpath = oFDialog.FileName;
                    log.Info($"    File: [{loadpath}] is being loaded.");

                    B5convert b5 = new B5convert(loadpath);
                    b5.Transform(progress);

                    Logging.Record_admin("add b5 data", $"加入健保藥物資料{loadpath}");

                    break;

                default:
                    break;
            }

            // 20200518 完成工作後可以更新資料
            _mainVM.Refresh_Data();

            #endregion 讀取檔案路徑
        }

        private void ReportProgress(object sender, ProgressReportModel e)
        {
            _mainVM.ProgressValue = e.PercentageComeplete;
        }
    }
}