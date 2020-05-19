using AutoIt;
using ClinicSysteMc.Model;
using Hardcodet.Wpf.TaskbarNotification;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;

namespace ClinicSysteMc.ViewModel.Converters
{
    internal class PIJIAconvert
    {
        private readonly DateTime _begindate;
        private readonly DateTime _enddate;
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private readonly TaskbarIcon tb = new TaskbarIcon();

        public PIJIAconvert(DateTime begindate, DateTime enddate)
        {
            _begindate = begindate;
            _enddate = enddate;
        }

        public void Convert()
        {
            // 20190608 created, 現在要匯入批價檔

            #region Declaration

            Microsoft.Office.Interop.Excel.Application MyExcel;
            List<String> filepath = new List<String>(); // 存放pijia檔
            string strYM = $"{_begindate.Year - 1911}{(_begindate.Month + 100).ToString().Substring(1)}";
            CSDataContext dc = new CSDataContext();

            #endregion Declaration

            #region Environment

            try
            {
                // 殺掉所有的EXCEL
                foreach (Process p in Process.GetProcessesByName("EXCEL"))
                {
                    p.Kill();
                }
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
                AutoItX.Sleep(500);
                // 日收入報表A
                // [NAME:chk允許完整筆數呈現]
                AutoItX.ControlClick("日收入報表A", "", "[NAME:chk允許完整筆數呈現]");
                // [NAME:chkIncludeInvalid]
                AutoItX.ControlClick("日收入報表A", "", "[NAME:chkIncludeInvalid]");
                AutoItX.Run($"C:\\vpn\\exe\\changeBillDTP.exe {_begindate}{_enddate}", @"C:\vpn\exe\");
            }
            catch (Exception ex)
            {
                Logging.Record_error(ex.Message);
            }

            #endregion Environment

            #region The Loop of 讀取檔案

            // 20190609 今天竟然完成了最難的部分
            // a FOR loop, LIST of A, B, C, D, E, F, G, H, I, J, K, L, M
            // Making a list
            // A 周孫元診所; B 聖愛; C 啟智; D 由根; E 方舟; F 景仁; G 香園; H 觀音; I 桃園; J 誠信; K 祥育; L 春暉; M 世美
            // 20200319 新增 N 華光
            List<String> lArea = new List<String>() { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N" };
            foreach (string a in lArea)
            {
                // [NAME:cmbArea]
                // AutoItX.ControlFocus("日收入報表A", "", "[NAME:cmbArea]")
                // AutoItX.Send(a)
                AutoItX.Sleep(1000); //這裡等一下
                AutoItX.ControlSend("日收入報表A", "", "[NAME:cmbArea]", a);
                // execute AutoIT
                // 日收入報表A
                AutoItX.Sleep(500); // 這裡等一下
                                    // [NAME:dtpStart]    input begin_date
                                    // [NAME:dtpEnd]      input end_date
                                    // [NAME:btnExcel]    click
                AutoItX.ControlClick("日收入報表A", "", "[NAME:btnExcel]");
                // EXCEL management
                //  ? 怎麼判斷有EXCEL
                //  查看有沒有EXCEL的process
                // Dim i As Int16
                // Do Until Process.GetProcessesByName("EXCEL").Length > 0
                //     AutoItX.Sleep(10)
                //     i += 1
                // Loop
                // MessageBox.Show(i)
                // MessageBox.Show("hi")
                AutoItX.Sleep(2000); // 20190614 這個點的等待真的很重要, 1000已經無法成功, 1500有9成成功, 選用2000
                                     // 20190609 這個點的等待很重要, <600 都找不到EXCEL; 700 可成功; 經過測試大約是200個循環左右
                                     //  20190609 原本還擔心這個方法沒效, 原來要等700ms以上, 就可以正常, 這也是未來可能出錯的地方, 如果有其它原因造成EXCEL開啟延後,就會錯誤
                Process[] pr = Process.GetProcessesByName("EXCEL");
                if (pr.Length > 0)
                {
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
                        // 檢查欄位, 如果欄位不對, 就不要處理了
                        // 要有: 狀態 收據號 批價人員 作廢日期 看診日期 午別 診別 科別 醫師 身分 就醫序號 優免 部分負擔 身分證號 患者姓名 醫療費用 掛號費用 部分負擔 押金 自付金額	藥費加重
                        //       欠收 折扣 應收金額 實收金額 收據說明	說明
                        // 刪除: 項次 病歷號 性別 生日 年齡 還款金額 電話 地址 國籍

                        // 檢查檔案格式
                        //  可以算出總筆數,第一行是標題,不算
                        // Dim listToAdd As New List(Of String) From {"狀態", "收據號", "批價人員", "作廢日期", "看診日期", "午別", "診別", "科別", "醫師", "身分", "就醫序號", "優免", "部分負擔",
                        // "身分證號", "患者姓名", "醫療費用", "掛號費用", "部分負擔", "押金", "自付金額", "藥費加重", "欠收", "折扣", "應收金額", "實收金額", "收據說明", "說明"}
                        // 20200319 修改,原本「部分負擔」改成「部分負擔說明」
                        List<string> listToAdd = new List<string>() { "狀態", "收據號", "批價人員", "作廢日期", "看診日期",
                            "午別", "診別", "科別", "醫師", "身分", "就醫序號", "優免", "部分負擔說明", "身分證號", "患者姓名", "醫療費用",
                            "掛號費用", "部分負擔", "押金", "自付金額", "藥費加重", "欠收", "折扣", "應收金額", "實收金額", "收據說明", "說明" };
                        List<string> listToDel = new List<string>() { "項次", "病歷號", "性別", "生日", "年齡", "還款金額", "電話", "地址", "國籍" };
                        //  檢查是否有充足欄位?
                        int j = 1;  // index
                        bool x = false;
                        do
                        {
                            if (ws.Cells[1, j].value == string.Empty)
                            {
                                x = true;
                            }
                            else
                            {
                                listToAdd.Remove(ws.Cells[1, j].value);
                                j++;
                            }
                        } while (!x);     // 當欄位空白就跳出迴圈
                        int totalColumn = j - 1;
                        if (listToAdd.Count == 0)
                        {
                            //                     Record_adm("匯入批價檔", "檔案格式正確")
                            //  格式正確
                        }
                        else
                        {
                            string output = string.Empty;
                            for (int i = 1; i <= listToAdd.Count; i++)
                            {
                                output += $"{listToAdd[i - 1]}, ";
                            }
                            Logging.Record_error("匯入批價檔格式不合,缺「" + output.Substring(0, output.Length - 2) + "」欄位");
                            //  格式不合,缺欄位
                        }
                        //  刪除欄位
                        x = false;
                        List<int> colToDel = new List<int>();
                        for (int i = 1; i <= totalColumn; i++)
                        {
                            if (listToDel.Remove(ws.Cells[1, j].value))
                            {
                                colToDel.Add(j);
                            }
                        }
                        for (j = 1; j <= colToDel.Count; j++)
                        {
                            ws.Columns[colToDel[colToDel.Count - j]].delete();
                        }
                        //  ====================================================================================================================================
                        // 製作自動檔名, 並存檔
                        string temp_filepath = @"C:\vpn\bills";
                        //  20190609 因為不小心多一個空格, 搞了好久除錯, 很辛苦啊
                        //  System.Runtime.InteropServices.COMException // 發生例外狀況於 HRESULT: 0x800A03EC//
                        // 存放目錄,不存在就要建立一個
                        if (!System.IO.Directory.Exists(temp_filepath)) System.IO.Directory.CreateDirectory(temp_filepath);
                        // 自動產生名字
                        temp_filepath += $"\\bill_{a}_{_begindate}_{_enddate}";
                        temp_filepath += $"_{DateTime.Now.Year}{(DateTime.Now.Month + 100).ToString().Substring(1)}{(DateTime.Now.Day + 100).ToString().Substring(1)}";
                        temp_filepath += DateTime.Now.TimeOfDay.ToString().Replace(":", "").Replace(".", "");
                        temp_filepath += ".csv";
                        filepath.Add(temp_filepath);
                        // wb.SaveAs(temp_filepath, Excel.XlFileFormat.xlCSV, vbNull, vbNull, False, False, Excel.XlSaveAsAccessMode.xlNoChange, vbNull, vbNull, vbNull, vbNull, vbNull)
                        wb.SaveAs(temp_filepath, Microsoft.Office.Interop.Excel.XlFileFormat.xlCSV);
                        wb.Close();
                        // 殺掉這個process
                        pr[0].Kill();
                        // 修剪csv
                        List<String> Lines = new List<String>(System.IO.File.ReadAllLines(temp_filepath, System.Text.Encoding.Default));
                        for (int LinesIndex = Lines.Count - 1; LinesIndex >= 0; LinesIndex--)
                        {
                            string Line = Lines[LinesIndex];
                            if (Line.Substring(0, 5) == ",,,,," || Line.Substring(0, 2) == "狀態")
                            {
                                // 該行有包含7的內容，則刪除
                                Lines.RemoveAt(LinesIndex);
                            }
                        }
                        // 覆寫整個檔案
                        System.IO.File.WriteAllLines(temp_filepath, Lines.ToArray(), System.Text.Encoding.Default);
                    }
                    catch (Exception ex)
                    {
                        Logging.Record_error(ex.Message);
                    }
                }
                else
                {
                    //  沒有的話,下一個
                    // do nothing
                    AutoItX.Sleep(500);
                }
                //  loop back, NEXT, 怎麼知道可以下一步了?
            }

            // close windows
            // 日收入報表A
            // [NAME:Cancel_Button]
            AutoItX.ControlClick("日收入報表A", "", "[NAME:Cancel_Button]");
            // [FrmMain] v1.0.0.67
            // [NAME:btnExit]
            //        aut.ControlClick("[FrmMain] v", "", "[NAME:btnExit]")

            #endregion The Loop of 讀取檔案

            #region 進行讀取資料

            // 20190609 created
            // 模仿匯入opd
            // add及update, update要清空CASENO, G
            // 20190612 重大修改,key值改為三個YM, bid, uid, 原因是bid在同一個月內重複太多了
            // 用varchar, 不要用varchar,不然比較時會出錯, VDATE空值仍會有8個空白
            try
            {
                int totalN = filepath.Count;
                // 開始回圈
                // 讀取每一筆檔案
                foreach (string f in filepath)
                {
                    List<string> Lines = new List<string>(System.IO.File.ReadAllLines(f, System.Text.Encoding.Default));
                    foreach (string Line in Lines)
                    {
                        // 用","分隔, 這也是CSV的意義
                        string[] LItem = Line.Split(',');
                        // 找到KEY值, YM, bid: YM=strYM, bid=Item(1), 第二個值就是bid
                        // 查詢,看看是否有重複
                        // 沒有重複就是新增, 有重複就是修改
                        var q = from o in dc.tbl_pijia
                                where (o.YM == strYM) && (o.bid == LItem[1]) && (o.uid == LItem[13])
                                select o;
                        if (q.Count() == 0) // 資料庫裡面沒有 INSERT
                        {
                            tbl_pijia newPijia = new tbl_pijia()
                            {
                                YM = strYM,
                                STATUS = LItem[0],
                                bid = LItem[1],
                                op = LItem[2],
                                VDATE = LItem[3],
                                SDATE = LItem[4],
                                VIST = LItem[5],
                                RMNO = LItem[6],
                                DEPTNAME = LItem[7],
                                DOCTNAME = LItem[8],
                                POSINAME = LItem[9],
                                HEATH_CARD = LItem[10],
                                Youmian = LItem[11],
                                PAYNO = LItem[12],
                                uid = LItem[13],
                                cname = LItem[14],
                                MedFee = int.Parse(LItem[15]),
                                RegFee = int.Parse(LItem[16]),
                                Copay = int.Parse(LItem[17]),
                                Deposit = int.Parse(LItem[18]),
                                SelfPay = int.Parse(LItem[19]),
                                PharmW = int.Parse(LItem[20]),
                                Arrears = int.Parse(LItem[21]),
                                Discount = int.Parse(LItem[22]),
                                AMTreceivable = int.Parse(LItem[23]),
                                AMTreceived = int.Parse(LItem[24]),
                                bremark = LItem[25],
                                remark = LItem[26]
                            };
                            dc.tbl_pijia.InsertOnSubmit(newPijia);
                            dc.SubmitChanges();
                        }
                        else
                        {
                            // 資料庫裡已經有了, 檢查是否有異,有異UPDATE
                            tbl_pijia oldPijia = q.ToList()[0];     // this is a record
                            string strChange = string.Empty;
                            bool bChange = false;
                            if (oldPijia.STATUS != LItem[0])
                            {
                                strChange += $";改狀態: {oldPijia.STATUS}=>{LItem[0]}";
                                bChange = true;
                                oldPijia.STATUS = LItem[0];
                            }
                            if (oldPijia.op != LItem[2])
                            {
                                strChange += $";改批價人員: {oldPijia.op}=>{LItem[2]}";
                                bChange = true;
                                oldPijia.op = LItem[2];
                            }
                            if (oldPijia.VDATE != LItem[3])
                            {
                                strChange += $";改作廢日期: {oldPijia.VDATE}=>{LItem[3]}";
                                bChange = true;
                                oldPijia.VDATE = LItem[3];
                            }
                            if (oldPijia.SDATE != LItem[4])
                            {
                                strChange += $";改看診日期: {oldPijia.SDATE}=>{LItem[4]}";
                                bChange = true;
                                oldPijia.SDATE = LItem[4];
                            }
                            if (oldPijia.VIST != LItem[5])
                            {
                                strChange += $";改午別: {oldPijia.VIST}=>{LItem[5]}";
                                bChange = true;
                                oldPijia.VIST = LItem[5];
                            }
                            if (oldPijia.RMNO != LItem[6])
                            {
                                strChange += $";改診別: {oldPijia.RMNO}=>{LItem[6]}";
                                bChange = true;
                                oldPijia.RMNO = LItem[6];
                            }
                            if (oldPijia.DEPTNAME != LItem[7])
                            {
                                strChange += $";改科別: {oldPijia.DEPTNAME}=>{LItem[7]}";
                                bChange = true;
                                oldPijia.DEPTNAME = LItem[7];
                            }
                            if (oldPijia.DOCTNAME != LItem[8])
                            {
                                strChange += $";改醫師: {oldPijia.DOCTNAME}=>{LItem[8]}";
                                bChange = true;
                                oldPijia.DOCTNAME = LItem[8];
                            }
                            if (oldPijia.POSINAME != LItem[9])
                            {
                                strChange += $";改身分: {oldPijia.POSINAME}=>{LItem[9]}";
                                bChange = true;
                                oldPijia.POSINAME = LItem[9];
                            }
                            if (oldPijia.HEATH_CARD != LItem[10])
                            {
                                strChange += $";改就醫序號: {oldPijia.HEATH_CARD}=>{LItem[10]}";
                                bChange = true;
                                oldPijia.HEATH_CARD = LItem[10];
                            }
                            if (oldPijia.Youmian != LItem[11])
                            {
                                strChange += $";改優免: {oldPijia.Youmian}=>{LItem[11]}";
                                bChange = true;
                                oldPijia.Youmian = LItem[11];
                            }
                            if (oldPijia.PAYNO != LItem[12])
                            {
                                strChange += $";改部分負擔: {oldPijia.PAYNO}=>{LItem[12]}";
                                bChange = true;
                                oldPijia.PAYNO = LItem[12];
                            }
                            if (oldPijia.cname != LItem[14])
                            {
                                strChange += $";改患者姓名: {oldPijia.cname}=>{LItem[14]}";
                                bChange = true;
                                oldPijia.cname = LItem[14];
                            }
                            if (oldPijia.MedFee != int.Parse(LItem[15]))
                            {
                                strChange += $";改醫療費用: {oldPijia.MedFee}=>{LItem[15]}";
                                bChange = true;
                                oldPijia.MedFee = int.Parse(LItem[15]);
                            }
                            if (oldPijia.RegFee != int.Parse(LItem[16]))
                            {
                                strChange += $";改掛號費用: {oldPijia.RegFee}=>{LItem[16]}";
                                bChange = true;
                                oldPijia.RegFee = int.Parse(LItem[16]);
                            }
                            if (oldPijia.Copay != int.Parse(LItem[17]))
                            {
                                strChange += $";改部分負擔: {oldPijia.Copay}=>{LItem[17]}";
                                bChange = true;
                                oldPijia.Copay = int.Parse(LItem[17]);
                            }
                            if (oldPijia.Deposit != int.Parse(LItem[18]))
                            {
                                strChange += $";改押金: {oldPijia.Deposit}=>{LItem[18]}";
                                bChange = true;
                                oldPijia.Deposit = int.Parse(LItem[18]);
                            }
                            if (oldPijia.SelfPay != int.Parse(LItem[19]))
                            {
                                strChange += $";改自付金額: {oldPijia.SelfPay}=>{LItem[19]}";
                                bChange = true;
                                oldPijia.SelfPay = int.Parse(LItem[19]);
                            }
                            if (oldPijia.PharmW != int.Parse(LItem[20]))
                            {
                                strChange += $";改藥費加重: {oldPijia.PharmW}=>{LItem[20]}";
                                bChange = true;
                                oldPijia.PharmW = int.Parse(LItem[20]);
                            }
                            if (oldPijia.Arrears != int.Parse(LItem[21]))
                            {
                                strChange += $";改欠收: {oldPijia.Arrears}=>{LItem[21]}";
                                bChange = true;
                                oldPijia.Arrears = int.Parse(LItem[21]);
                            }
                            if (oldPijia.Discount != int.Parse(LItem[22]))
                            {
                                strChange += $";改折扣: {oldPijia.Discount}=>{LItem[22]}";
                                bChange = true;
                                oldPijia.Discount = int.Parse(LItem[22]);
                            }
                            if (oldPijia.AMTreceivable != int.Parse(LItem[23]))
                            {
                                strChange += $";改應收金額: {oldPijia.AMTreceivable}=>{LItem[23]}";
                                bChange = true;
                                oldPijia.AMTreceivable = int.Parse(LItem[23]);
                            }
                            if (oldPijia.AMTreceived != int.Parse(LItem[24]))
                            {
                                strChange += $";改實收金額: {oldPijia.AMTreceived}=>{LItem[24]}";
                                bChange = true;
                                oldPijia.AMTreceived = int.Parse(LItem[24]);
                            }
                            if (oldPijia.bremark != LItem[25])
                            {
                                strChange += $";改收據說明: {oldPijia.bremark}=>{LItem[25]}";
                                bChange = true;
                                oldPijia.bremark = LItem[25];
                            }
                            if (oldPijia.remark != LItem[26])
                            {
                                strChange += $";改說明: {oldPijia.remark}=>{LItem[26]}$";
                                bChange = true;
                                oldPijia.remark = LItem[26];
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
                                // 做記錄
                                Logging.Record_admin("修改批價資料", $"{strYM}-{LItem[13]}: {strChange}");
                            }
                        }
                    }
                    Logging.Record_admin("新增批價檔: ", f);
                }

                // 現再來配對, 使用Stored Procedure
                // 第一步Pijia配上CASENO
                // 第二步檢查CASENO是否1to1配上Pijia, 若是進行配對,並顯示正確,若否回傳錯誤幾筆,並且紀錄下來
            }
            catch (Exception ex)
            {
                Logging.Record_error(ex.Message);
            }

            #endregion 進行讀取資料

            #region 進行配對

            // 20190614 created
            // 目的是將tbl_pijia和tbl_opd配對起來
            // 分為兩步
            // 第一步將tbl_pijia配上CASENO
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