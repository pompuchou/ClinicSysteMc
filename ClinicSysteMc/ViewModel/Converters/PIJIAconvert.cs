using AutoIt;
using ClinicSysteMc.Model;
using Hardcodet.Wpf.TaskbarNotification;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace ClinicSysteMc.ViewModel.Converters
{
    internal class PIJIAconvert
    {
        private readonly DateTime _begindate;
        private readonly DateTime _enddate;
        private string BeginDate;
        private string EndDate;
        private int[] header_order;
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private readonly TaskbarIcon tb = new TaskbarIcon();

        public PIJIAconvert(DateTime begindate, DateTime enddate)
        {
            _begindate = begindate;
            _enddate = enddate;
        }

        public async void Convert()
        {
            // 20190608 created, 現在要匯入批價檔

            #region Environment

            try
            {
                // 營造環境
                Process[] isAdvn = Process.GetProcessesByName("THCAdvancedBillingReport");
                if (isAdvn.Length == 0)     // 測試"日報表清單"是否有打開
                {
                    // 有無登入系統?
                    Thesis.LogIN();
                    // 如果沒有打開就打開"日報表清單"
                    AutoItX.Run(@"C:\Program Files (x86)\THESE\杏雲醫療資訊系統\THCAdvancedBillingReport.exe", @"C:\Program Files (x86)\THESE\杏雲醫療資訊系統\");
                    System.Threading.Thread.Sleep(300);
                }
                // [FrmMain] v1.0.0.67
                AutoItX.WinWaitActive("[FrmMain] v");
                // [NAME:btnDailyIncome]
                // 按下去
                AutoItX.ControlClick("[FrmMain] v", "", "[NAME:btnDailyIncome]");
                AutoItX.WinWaitActive("日收入報表A");
                // 日收入報表A
                // [NAME:chk允許完整筆數呈現]
                AutoItX.ControlClick("日收入報表A", "", "[NAME:chk允許完整筆數呈現]");
                // [NAME:chkIncludeInvalid]
                AutoItX.ControlClick("日收入報表A", "", "[NAME:chkIncludeInvalid]");

                BeginDate = $"{_begindate.Year}{(_begindate.Month + 100).ToString().Substring(1)}{(_begindate.Day + 100).ToString().Substring(1)}";
                EndDate = $"{_enddate.Year}{(_enddate.Month + 100).ToString().Substring(1)}{(_enddate.Day + 100).ToString().Substring(1)}";
                string Execution = $"C:\\vpn\\exe\\changeBillDTP.exe {BeginDate}{EndDate}";
                AutoItX.Run(Execution, @"C:\vpn\exe\");

                log.Info($"執行changeBillDTP ");
            }
            catch (Exception ex)
            {
                Logging.Record_error(ex.Message);
                log.Error(ex.Message);
            }

            #endregion Environment

            // 資料都從excel檔讀進記憶體, 也寫入csv檔
            Dictionary<string, object[,]> data = RetrievePIJIA();

            // 空值就結束, 沒有後面什麼事
            if (data is null) return;

            string strYM = $"{_begindate.Year - 1911}{(_begindate.Month + 100).ToString().Substring(1)}";
            List<Task<PIJIAresult>> tasks = new List<Task<PIJIAresult>>();

            #region 進行讀取資料

            // 20190609 created
            // 模仿匯入opd
            // add及update, update要清空CASENO, G
            // 20190612 重大修改,key值改為三個YM, bid, uid, 原因是bid在同一個月內重複太多了
            // 用varchar, 不要用varchar,不然比較時會出錯, VDATE空值仍會有8個空白
            try
            {
                // 開始回圈
                // 讀取每一筆檔案
                log.Info($"begin async process.");
                foreach (var d in data)
                {
                    // 20200519 程式碼移出, 程式不那麼臃腫, 較好維護
                    tasks.Add(WriteIntoSQL_async(d.Value, strYM));
                }
                PIJIAresult[] result = await Task.WhenAll(tasks);

                int total_NewPIJIA = (from p in result
                                      select p.NewPIJIA).Sum();
                int total_ChangePIJIA = (from p in result
                                         select p.ChangePIJIA).Sum();
                int total_AllPIJIA = (from p in result
                                      select p.AllPIJIA).Sum();

                log.Info($"end async process.");

                string output = $"共處理{total_AllPIJIA}筆資料, 其中{total_NewPIJIA}筆新批價, 修改{total_ChangePIJIA}筆批價.";
                log.Info(output);
                tb.ShowBalloonTip("完成", output, BalloonIcon.Info);
                Logging.Record_admin("PIJIA add/change", output);

                return;
            }
            catch (Exception ex)
            {
                Logging.Record_error(ex.Message);
            }

            #endregion 進行讀取資料

            // 20200519 移出程式碼, 不要顯得太臃腫
            Matching();
        }

        private Dictionary<string, object[,]> RetrievePIJIA()
        {
            // 20200519 獨立成一個副程式

            #region Declaration

            Microsoft.Office.Interop.Excel.Application MyExcel;
            Dictionary<string, object[,]> output = new Dictionary<string, object[,]>();

            #endregion Declaration

            #region The Loop of 讀取檔案

            // 20190609 今天竟然完成了最難的部分
            // a FOR loop, LIST of A, B, C, D, E, F, G, H, I, J, K, L, M
            // Making a list
            // A 周孫元診所; B 聖愛; C 啟智; D 由根; E 方舟; F 景仁; G 香園; H 觀音; I 桃園; J 誠信; K 祥育; L 春暉; M 世美
            // 20200319 新增 N 華光
            string[] lArea = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N" };
            foreach (string a in lArea)
            {
                try
                {
                    // 殺掉所有的EXCEL
                    foreach (Process p in Process.GetProcessesByName("EXCEL"))
                    {
                        p.Kill();
                    }
                }
                catch (Exception ex)
                {
                    log.Error($"ignore this {ex.Message}");
                }

                // [NAME:cmbArea]
                // AutoItX.ControlFocus("日收入報表A", "", "[NAME:cmbArea]")
                // AutoItX.Send(a)
                AutoItX.Sleep(3000); //這裡等一下
                AutoItX.ControlSend("日收入報表A", "", "[NAME:cmbArea]", a);
                log.Info($"現在處理{a}");
                // execute AutoIT
                // 日收入報表A
                AutoItX.Sleep(500); // 這裡等一下

                log.Info("Start to Click.");
                int timeout = 0; // 20 sec timeout
                do
                {
                    AutoItX.ControlClick("日收入報表A", "", "[NAME:btnExcel]");
                    AutoItX.Sleep(3000);
                    log.Info($"按下btnExcel {timeout + 1} time.");
                    timeout++;
                } while (Process.GetProcessesByName("EXCEL").Length == 0 && timeout < 2);

                Process[] pr = Process.GetProcessesByName("EXCEL");
                if (pr.Length > 0)
                {
                    log.Info("Excel exists now.");
                    //  有的話,excel.application, getobject, 存檔, 存檔案位置, 供匯入用
                    //  winwait實驗可行, 可偵測excel已經完成
                    // AutoItX.WinWaitActive("活頁簿"), 實際測試失敗, 改用DO Loop
                    //  後來發現process建立後, 一段時間才會建立windows
                    do
                    {
                        AutoItX.Sleep(100);
                    } while (AutoItX.WinExists("活頁簿") == 0);
                    // AutoItX.Sleep(10000), 用等的,等10秒大多有效,但不能保證,且也許不用10秒,這樣就浪費了, 應該要個別化
                    // 好在發現visibility可以有效等到整個檔案製作完成
                    MyExcel = (Microsoft.Office.Interop.Excel.Application)Marshal.GetActiveObject("Excel.Application");

                    do
                    {
                        AutoItX.Sleep(100);
                    } while (!MyExcel.Visible);
                    // 現在開始excel 的處理
                    try
                    {
                        Microsoft.Office.Interop.Excel.Workbook wb = MyExcel.ActiveWorkbook;
                        Microsoft.Office.Interop.Excel.Worksheet ws = wb.ActiveSheet;

                        // 要刪除什麼欄位,合計等等資料
                        //  ====================================================================================================================================
                        // 檢查欄位, 如果欄位不對, 就不要處理了, 一共要以下27欄位
                        string[] header = { "狀態", "收據號", "批價人員", "作廢日期", "看診日期", "午別", "診別", "科別", "醫師", "身分",
                                            "就醫序號", "優免", "部分負擔說明", "身分證號", "患者姓名", "醫療費用", "掛號費用", "部分負擔", "押金", "自付金額",
                                            "藥費加重", "欠收", "折扣", "應收金額", "實收金額", "收據說明", "說明"};
                        // excel檔有幾欄?
                        int total_col_n = ws.UsedRange.Columns.Count;
                        int last_row = ws.UsedRange.Rows.Count;

                        // 刪除欄位
                        for (int idx = total_col_n; idx > 0; idx--)
                        {
                            if (!header.Contains((string)ws.Cells[1, idx].Value))
                            {
                                ws.Columns[idx].delete();
                            }
                        }

                        if (header_order is null)
                        {
                            log.Info("Check header_order.");
                            // 27欄位以上的都刪除掉
                            header_order = new int[header.Length + 1]; // 0不要用, 使用1 - 27

                            // 檢查格式, header.Length = 27
                            for (int i = 1; i <= header.Length; i++)
                            {
                                for (int j = 1; j <= total_col_n; j++)
                                {
                                    if ((string)ws.Cells[1, j].Value == header[i - 1])
                                    {
                                        header_order[i] = j;
                                        break;
                                    }
                                }
                                // 只要任一header成員找不到ws相對應的文字, 那就是找不到
                                if (header_order[i] == 0)
                                {
                                    // 寫入Error Log
                                    Logging.Record_error($"{a}輸入的批價檔案格式不對");
                                    log.Error($"{a}輸入的批價檔案格式不對");
                                    tb.ShowBalloonTip("錯誤", $"{a}檔案格式不對", BalloonIcon.Error);
                                    return null;
                                }
                            }

                            // 格式正確
                            Logging.Record_admin("讀取批價檔", $"{a}輸入的批價檔案格式正確");
                            log.Info($"{a}輸入的批價檔案格式正確");
                        }

                        // 刪除最後一列, 存入檔案
                        ws.Rows[last_row].delete();
                        //  ====================================================================================================================================
                        // 製作自動檔名, 並存檔
                        string temp_filepath = @"C:\vpn\bills";
                        //  20190609 因為不小心多一個空格, 搞了好久除錯, 很辛苦啊
                        //  System.Runtime.InteropServices.COMException // 發生例外狀況於 HRESULT: 0x800A03EC//
                        // 存放目錄,不存在就要建立一個
                        if (!System.IO.Directory.Exists(temp_filepath)) System.IO.Directory.CreateDirectory(temp_filepath);
                        // 自動產生名字
                        temp_filepath += $"\\bill_{a}_{BeginDate}_{EndDate}";
                        temp_filepath += $"_{DateTime.Now.Year}{(DateTime.Now.Month + 100).ToString().Substring(1)}{(DateTime.Now.Day + 100).ToString().Substring(1)}";
                        temp_filepath += DateTime.Now.TimeOfDay.ToString().Replace(":", "").Replace(".", "");
                        temp_filepath += ".csv";
                        // wb.SaveAs(temp_filepath, Excel.XlFileFormat.xlCSV, vbNull, vbNull, False, False, Excel.XlSaveAsAccessMode.xlNoChange, vbNull, vbNull, vbNull, vbNull, vbNull)
                        wb.SaveAs(temp_filepath, Microsoft.Office.Interop.Excel.XlFileFormat.xlCSV);

                        //// 刪除第一列, 存入data
                        //ws.Rows[1].delete();

                        output.Add(a, ws.UsedRange.Value2);
                    }
                    catch (Exception ex)
                    {
                        Logging.Record_error(ex.Message);
                        log.Error(ex.Message);
                    }
                }
            }

            // close windows
            // 日收入報表A
            // [NAME:Cancel_Button]
            AutoItX.ControlClick("日收入報表A", "", "[NAME:Cancel_Button]");
            // [FrmMain] v1.0.0.67
            // [NAME:btnExit]
            AutoItX.ControlClick("[FrmMain] v", "", "[NAME:btnExit]");

            #endregion The Loop of 讀取檔案

            return output;
        }

        private async Task<PIJIAresult> WriteIntoSQL_async(object[,] data, string strYM)
        {
            // 20200519 獨立成一個副程式
            CSDataContext dc = new CSDataContext();
            int totalN = data.GetUpperBound(0);
            int add_N = 0;
            int change_N = 0;
            int all_N = 0;

            await Task.Run(() =>
            {
                for (int i = 2; i <= totalN; i++)
                {
                    // 找到KEY值, YM, bid: YM=strYM, bid=Item(1), 第二個值就是bid
                    // 查詢,看看是否有重複
                    // 沒有重複就是新增, 有重複就是修改
                    var q = from o in dc.tbl_pijia
                            where (o.YM == strYM) && (o.bid == (string)data[i, 2]) && (o.uid == (string)data[i, 14])
                            select o;
                    if (q.Count() == 0) // 資料庫裡面沒有 INSERT
                    {
                        // MedFee 有小數點, 所以要先變成double, 再變成int
                        tbl_pijia newPijia = new tbl_pijia()
                        {
                            YM = strYM,
                            STATUS = (string)data[i, header_order[1]],
                            bid = (string)data[i, header_order[2]],
                            op = (string)data[i, header_order[3]],
                            VDATE = (string)data[i, header_order[4]] ?? string.Empty,
                            SDATE = (string)data[i, header_order[5]],
                            VIST = (string)data[i, header_order[6]],
                            RMNO = (string)data[i, header_order[7]],
                            DEPTNAME = (string)data[i, header_order[8]],
                            DOCTNAME = (string)data[i, header_order[9]],
                            POSINAME = (string)data[i, header_order[10]],
                            HEATH_CARD = (string)data[i, header_order[11]] ?? string.Empty,
                            Youmian = (string)data[i, header_order[12]] ?? string.Empty,
                            PAYNO = (string)data[i, header_order[13]] ?? string.Empty,
                            uid = (string)data[i, header_order[14]],
                            cname = (string)data[i, header_order[15]],
                            MedFee = (int)double.Parse((string)data[i, header_order[16]]),
                            RegFee = int.Parse((string)data[i, header_order[17]]),
                            Copay = int.Parse((string)data[i, header_order[18]]),
                            Deposit = int.Parse((string)data[i, header_order[19]]),
                            SelfPay = int.Parse((string)data[i, header_order[20]]),
                            PharmW = int.Parse((string)data[i, header_order[21]]),
                            Arrears = int.Parse((string)data[i, header_order[22]]),
                            Discount = int.Parse((string)data[i, header_order[23]]),
                            AMTreceivable = int.Parse((string)data[i, header_order[24]]),
                            AMTreceived = int.Parse((string)data[i, header_order[25]]),
                            bremark = (string)data[i, header_order[26]] ?? string.Empty,
                            remark = (string)data[i, header_order[27]] ?? string.Empty
                        };
                        dc.tbl_pijia.InsertOnSubmit(newPijia);
                        dc.SubmitChanges();
                        add_N++;
                    }
                    else
                    {
                        // 資料庫裡已經有了, 檢查是否有異,有異UPDATE
                        tbl_pijia oldPijia = q.ToList()[0];     // this is a record
                        string strChange = string.Empty;
                        bool bChange = false;
                        if (oldPijia.STATUS != (string)data[i, header_order[1]])
                        {
                            strChange += $";改狀態: {oldPijia.STATUS}=>{(string)data[i, header_order[1]]}";
                            bChange = true;
                            oldPijia.STATUS = (string)data[i, header_order[1]];
                        }
                        if (oldPijia.op != (string)data[i, header_order[3]])
                        {
                            strChange += $";改批價人員: {oldPijia.op}=>{(string)data[i, header_order[3]]}";
                            bChange = true;
                            oldPijia.op = (string)data[i, header_order[3]];
                        }
                        if (oldPijia.VDATE != ((string)data[i, header_order[4]] ?? string.Empty))
                        {
                            strChange += $";改作廢日期: {oldPijia.VDATE}=>{((string)data[i, header_order[4]] ?? string.Empty)}";
                            bChange = true;
                            oldPijia.VDATE = ((string)data[i, header_order[4]] ?? string.Empty);
                        }
                        if (oldPijia.SDATE != (string)data[i, header_order[5]])
                        {
                            strChange += $";改看診日期: {oldPijia.SDATE}=>{(string)data[i, header_order[5]]}";
                            bChange = true;
                            oldPijia.SDATE = (string)data[i, header_order[5]];
                        }
                        if (oldPijia.VIST != (string)data[i, header_order[6]])
                        {
                            strChange += $";改午別: {oldPijia.VIST}=>{(string)data[i, header_order[6]]}";
                            bChange = true;
                            oldPijia.VIST = (string)data[i, header_order[6]];
                        }
                        if (oldPijia.RMNO != (string)data[i, header_order[7]])
                        {
                            strChange += $";改診別: {oldPijia.RMNO}=>{(string)data[i, header_order[7]]}";
                            bChange = true;
                            oldPijia.RMNO = (string)data[i, header_order[7]];
                        }
                        if (oldPijia.DEPTNAME != (string)data[i, header_order[8]])
                        {
                            strChange += $";改科別: {oldPijia.DEPTNAME}=>{(string)data[i, header_order[8]]}";
                            bChange = true;
                            oldPijia.DEPTNAME = (string)data[i, header_order[8]];
                        }
                        if (oldPijia.DOCTNAME != (string)data[i, header_order[9]])
                        {
                            strChange += $";改醫師: {oldPijia.DOCTNAME}=>{(string)data[i, header_order[9]]}";
                            bChange = true;
                            oldPijia.DOCTNAME = (string)data[i, header_order[9]];
                        }
                        if (oldPijia.POSINAME != (string)data[i, header_order[10]])
                        {
                            strChange += $";改身分: {oldPijia.POSINAME}=>{(string)data[i, header_order[10]]}";
                            bChange = true;
                            oldPijia.POSINAME = (string)data[i, header_order[10]];
                        }
                        if (oldPijia.HEATH_CARD != ((string)data[i, header_order[11]] ?? string.Empty))
                        {
                            strChange += $";改就醫序號: {oldPijia.HEATH_CARD}=>{((string)data[i, header_order[11]] ?? string.Empty)}";
                            bChange = true;
                            oldPijia.HEATH_CARD = ((string)data[i, header_order[11]] ?? string.Empty);
                        }
                        if (oldPijia.Youmian != ((string)data[i, header_order[12]] ?? string.Empty))
                        {
                            strChange += $";改優免: {oldPijia.Youmian}=>{((string)data[i, header_order[12]] ?? string.Empty)}";
                            bChange = true;
                            oldPijia.Youmian = ((string)data[i, header_order[12]] ?? string.Empty);
                        }
                        if (oldPijia.PAYNO != ((string)data[i, header_order[13]] ?? string.Empty))
                        {
                            strChange += $";改部分負擔: {oldPijia.PAYNO}=>{((string)data[i, header_order[13]] ?? string.Empty)}";
                            bChange = true;
                            oldPijia.PAYNO = ((string)data[i, header_order[13]] ?? string.Empty);
                        }
                        if (oldPijia.cname != (string)data[i, header_order[15]])
                        {
                            strChange += $";改患者姓名: {oldPijia.cname}=>{(string)data[i, header_order[15]]}";
                            bChange = true;
                            oldPijia.cname = (string)data[i, header_order[15]];
                        }
                        if (oldPijia.MedFee != (int)double.Parse((string)data[i, header_order[16]]))
                        {
                            strChange += $";改醫療費用: {oldPijia.MedFee}=>{(string)data[i, header_order[16]]}";
                            bChange = true;
                            oldPijia.MedFee = (int)double.Parse((string)data[i, header_order[16]]);
                        }
                        if (oldPijia.RegFee != int.Parse((string)data[i, header_order[17]]))
                        {
                            strChange += $";改掛號費用: {oldPijia.RegFee}=>{(string)data[i, header_order[17]]}";
                            bChange = true;
                            oldPijia.RegFee = int.Parse((string)data[i, header_order[17]]);
                        }
                        if (oldPijia.Copay != int.Parse((string)data[i, header_order[18]]))
                        {
                            strChange += $";改部分負擔: {oldPijia.Copay}=>{(string)data[i, header_order[18]]}";
                            bChange = true;
                            oldPijia.Copay = int.Parse((string)data[i, header_order[18]]);
                        }
                        if (oldPijia.Deposit != int.Parse((string)data[i, header_order[19]]))
                        {
                            strChange += $";改押金: {oldPijia.Deposit}=>{(string)data[i, header_order[19]]}";
                            bChange = true;
                            oldPijia.Deposit = int.Parse((string)data[i, header_order[19]]);
                        }
                        if (oldPijia.SelfPay != int.Parse((string)data[i, header_order[20]]))
                        {
                            strChange += $";改自付金額: {oldPijia.SelfPay}=>{(string)data[i, header_order[20]]}";
                            bChange = true;
                            oldPijia.SelfPay = int.Parse((string)data[i, header_order[20]]);
                        }
                        if (oldPijia.PharmW != int.Parse((string)data[i, header_order[21]]))
                        {
                            strChange += $";改藥費加重: {oldPijia.PharmW}=>{(string)data[i, header_order[21]]}";
                            bChange = true;
                            oldPijia.PharmW = int.Parse((string)data[i, header_order[21]]);
                        }
                        if (oldPijia.Arrears != int.Parse((string)data[i, header_order[22]]))
                        {
                            strChange += $";改欠收: {oldPijia.Arrears}=>{(string)data[i, header_order[22]]}";
                            bChange = true;
                            oldPijia.Arrears = int.Parse((string)data[i, header_order[22]]);
                        }
                        if (oldPijia.Discount != int.Parse((string)data[i, header_order[23]]))
                        {
                            strChange += $";改折扣: {oldPijia.Discount}=>{(string)data[i, header_order[23]]}";
                            bChange = true;
                            oldPijia.Discount = int.Parse((string)data[i, header_order[23]]);
                        }
                        if (oldPijia.AMTreceivable != int.Parse((string)data[i, header_order[24]]))
                        {
                            strChange += $";改應收金額: {oldPijia.AMTreceivable}=>{(string)data[i, header_order[24]]}";
                            bChange = true;
                            oldPijia.AMTreceivable = int.Parse((string)data[i, header_order[24]]);
                        }
                        if (oldPijia.AMTreceived != int.Parse((string)data[i, header_order[25]]))
                        {
                            strChange += $";改實收金額: {oldPijia.AMTreceived}=>{(string)data[i, header_order[25]]}";
                            bChange = true;
                            oldPijia.AMTreceived = int.Parse((string)data[i, header_order[25]]);
                        }
                        if (oldPijia.bremark != ((string)data[i, header_order[26]] ?? string.Empty))
                        {
                            strChange += $";改收據說明: {oldPijia.bremark}=>{((string)data[i, header_order[26]] ?? string.Empty)}";
                            bChange = true;
                            oldPijia.bremark = ((string)data[i, header_order[26]] ?? string.Empty);
                        }
                        if (oldPijia.remark != ((string)data[i, header_order[27]] ?? string.Empty))
                        {
                            strChange += $";改說明: {oldPijia.remark}=>{((string)data[i, header_order[27]] ?? string.Empty)}";
                            bChange = true;
                            oldPijia.remark = ((string)data[i, header_order[27]] ?? string.Empty);
                        }
                        if (bChange)
                        {
                            // tbl_opd的Pijia欄位也要歸零
                            var r = from opd in dc.tbl_opd
                                    where opd.CASENO == oldPijia.CASENO
                                    select opd;
                            tbl_opd opdOPD = r.ToList()[0];
                            opdOPD.Pijia = null;
                            // CASENO, G要歸零
                            oldPijia.CASENO = null;
                            oldPijia.G = null;
                            // 做實改變
                            dc.SubmitChanges();
                            change_N++;
                            // 做記錄
                            Logging.Record_admin("修改批價資料", $"{strYM}-{(string)data[i, 13]}: {strChange}");
                        }
                    }
                    all_N++;
                }
            });

            return new PIJIAresult()
            {
                NewPIJIA = add_N,
                ChangePIJIA = change_N,
                AllPIJIA = all_N
            };
        }

        private void Matching()
        {
            #region 進行配對

            // 現再來配對, 使用Stored Procedure
            // 第一步Pijia配上CASENO
            // 第二步檢查CASENO是否1to1配上Pijia, 若是進行配對,並顯示正確,若否回傳錯誤幾筆,並且紀錄下來

            // 20200519 transcribed
            // 20190614 created
            // 目的是將tbl_pijia和tbl_opd配對起來
            // 分為兩步
            // 第一步將tbl_pijia配上CASENO
            CSDataContext dc = new CSDataContext();

            var q1 = from cs in dc.sp_CASENO_for_pijia().AsEnumerable()
                     select cs;
            string strOutput = $"{_begindate:d}_{_enddate:d}: {q1.First().rows_affected}筆配對";
            Logging.Record_admin("批價檔配對STEP1 Pijia", strOutput);
            log.Info($"批價檔配對STEP1 Pijia: {strOutput}");
            tb.ShowBalloonTip("批價檔配對STEP1 Pijia", strOutput, BalloonIcon.Info);
            var q2 = from pj in dc.sp_PIJIA_for_opd().AsEnumerable()
                     select pj;
            strOutput = string.Empty;
            foreach (var q in q2)
            {
                strOutput += $"{q.CASENO} {q.SDATE} {q.VIST} {q.RMNO} {q.bid} {q.cname};";
            }
            if (strOutput == string.Empty)
            {
                Logging.Record_admin("批價檔配對STEP2 OPD", "沒有重複");
                log.Info($"批價檔配對STEP2 OPD: 沒有重複");
                tb.ShowBalloonTip("批價檔配對STEP2 OPD", "沒有重複", BalloonIcon.Info);
            }
            else
            {
                strOutput += ";請修正後再上傳";
                Logging.Record_admin("批價檔配對STEP2 OPD", strOutput);
                log.Info($"批價檔配對STEP2 OPD: {strOutput}");
                tb.ShowBalloonTip("批價檔配對STEP2 OPD, CASE有重複值:", strOutput, BalloonIcon.Info);
            }

            #endregion 進行配對
        }
    }
}