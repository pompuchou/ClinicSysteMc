using ClinicSysteMc.Model;
using Hardcodet.Wpf.TaskbarNotification;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
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
            // Xml格式的index=2
            OpenFileDialog oFDialog = new OpenFileDialog();
            switch ((string)parameter)
            {
                case "門診":
                    oFDialog.Filter = "xml|*.xml";
                    break;

                default:
                    break;
            }

            if (oFDialog.ShowDialog() != true) return;
            loadpath = oFDialog.FileName;

            #endregion 讀取檔案路徑

            // 依照parameter, 不同來源: 申報匯入, 門診, 病患, 醫令, 檢驗, 指向不同方向
            ImportOPD(loadpath);
            Logging.Record_admin("add opd", "匯入門診檔案 Manual");
        }

        private void ImportOPD(string loadpath)
        {
            #region 宣告

            DataSet ds = new DataSet();
            DataTable dtO = new DataTable();
            DataTable dtP = new DataTable();

            #endregion 宣告

            #region 整理datatable

            //整理datatable, 分拆成兩個, 一旦可以通過,那這個檔案應該沒有問題,如果有問題,就不是正確的檔案

            try
            {
                ds.ReadXml(loadpath, XmlReadMode.ReadSchema);
                dtP = ds.Tables[0];  // dtP for tbl_opd_order, P stands for prescription
                dtP.Columns.Remove("STATUS");
                dtP.Columns.Remove("REGNO");
                dtP.Columns.Remove("PNAME");
                dtP.Columns.Remove("SEX");
                dtP.Columns.Remove("BIRTH");
                dtP.Columns.Remove("ORI_TOTAL");
                dtP.Columns.Remove("TOTAL");
                dtP.Columns.Remove("AMT8");
                dtP.Columns.Remove("RECT_NO");
                dtO = dtP.Copy();
                // 移除dtO不必要欄位, 先轉移給暫存檔, 因為要distinct, O stands for OPD
                dtO.Columns.Remove("CODE");
                dtO.Columns.Remove("ENAME");
                dtO.Columns.Remove("TIMES_DAY");
                dtO.Columns.Remove("METHODE");
                dtO.Columns.Remove("TIME_QTY1");
                dtO.Columns.Remove("DAYS");
                dtO.Columns.Remove("BILL_QTY");
                dtO.Columns.Remove("CHRONIC");
                dtO.Columns.Remove("PUT_TYPE");
                dtO.Columns.Remove("HC");
                dtO.Columns.Remove("PRICE");
                dtO.Columns.Remove("AMT");
                dtO.Columns.Remove("ORI_AMT");
                dtO.Columns.Remove("CLASS");
                dtO.Columns.Remove("PRN_CODE");
                dtO.Columns.Remove("RESULT");
                // 移除dtP不需要的欄位(for tbl_opd_order)
                dtP.Columns.Remove("VIST");
                dtP.Columns.Remove("RMNO");
                dtP.Columns.Remove("DEPTNAME");
                dtP.Columns.Remove("DOCTNAME");
                dtP.Columns.Remove("POSINAME");
                dtP.Columns.Remove("PAYNO");
                dtP.Columns.Remove("HEATH_CARD");
                dtP.Columns.Remove("STEXT");
                dtP.Columns.Remove("OTEXT");
                dtP.Columns.Remove("ICDCODE1");
                dtP.Columns.Remove("ICDCODE2");
                dtP.Columns.Remove("ICDCODE3");
                dtO = dtO.DefaultView.ToTable(true, new string[] {"CASENO", "SDATE", "VIST", "RMNO", "DEPTNAME", "DOCTNAME",
                                          "IDNO", "POSINAME", "PAYNO", "HEATH_CARD", "STEXT", "OTEXT", "ICDCODE1", "ICDCODE2",
                                          "ICDCODE3" });    // true stands for distinct
            }
            catch (Exception ex)
            {
                Logging.Record_error(ex.Message);
                log.Error(ex.Message);
                tb.ShowBalloonTip("錯誤!", ex.Message, BalloonIcon.Error);
                return;
            }

            // 通過測試
            Logging.Record_admin("OPD file format", "correct");
            log.Info("OPD XML 檔案格式正確");
            tb.ShowBalloonTip("正確", "OPD XML 檔案格式正確", BalloonIcon.Info);

            #endregion 整理datatable

            #region 進行讀取資料

            //Main.ProgressBar1.Minimum = 1
            //Main.ProgressBar1.Maximum = totalN
            CSDataContext dc = new CSDataContext();

            // 開始回圈
            foreach (DataRow dtO_Row in dtO.Rows)
            {
                //Main.ProgressBar1.Value = i + 1  // 顯示一下進度
                // 檢查案號是否已經在資料庫中, dtO.CASENO, tbl_opd.CASENO
                string strCASENO = (string)dtO_Row["CASENO"];
                if (string.IsNullOrEmpty(strCASENO))
                {
                    Logging.Record_error("在輸入門診資料時, 缺少案號CASENO");
                    log.Error("在輸入門診資料時, 缺少案號CASENO");
                    tb.ShowBalloonTip("錯誤!", "在輸入門診資料時, 缺少案號CASENO", BalloonIcon.Error);
                    // 下一個
                    continue;
                }

                var q = from o in dc.tbl_opd
                         where o.CASENO == strCASENO
                         select o;
                if (q.Count() == 0) // 資料庫裡面沒有 INSERT
                {
                    try
                    {
                        tbl_opd newOPD = new tbl_opd()
                        {
                            CASENO = strCASENO, // CASENO
                            VIST = (string)dtO_Row["VIST"], // VIST
                            RMNO = byte.Parse((string)dtO_Row["RMNO"]), // RMNO
                            uid = (string)dtO_Row["IDNO"], // uid
                            DEPTNAME = (string)dtO_Row["DEPTNAME"], // DEPTNAME
                            DOCTNAME = (string)dtO_Row["DOCTNAME"], // DOCTNAME
                            POSINAME = (string)dtO_Row["POSINAME"], // POSINAME
                            PAYNO = (string)dtO_Row["PAYNO"],  // PAYNO
                            HEATH_CARD = (string)dtO_Row["HEATH_CARD"], // HEATH_CARD
                            ICDCODE1 = (string)dtO_Row["ICDCODE1"], // ICDCODE1
                            ICDCODE2 = (string)dtO_Row["ICDCODE2"], // ICDCODE2
                            ICDCODE3 = (string)dtO_Row["ICDCODE3"], // ICDCODE3
                            INS_CODE = "A", // INS_CODE, default value "A"
                            STEXT = (string)dtO_Row["STEXT"], // STEXT
                            OTEXT = (string)dtO_Row["OTEXT"] // OTEXT
                        };

                        string tempstr;
                        tempstr = dtO_Row["SDATE"].ToString();
                        if (DateTime.TryParse($"{tempstr.Substring(0, 4)}/{tempstr.Substring(4, 2)}/{tempstr.Substring(6, 2)}", out DateTime temp_date))
                        {
                            newOPD.SDATE = temp_date;
                        }
                        dc.tbl_opd.InsertOnSubmit(newOPD);
                        dc.SubmitChanges();

                        // tbl_opd沒有資料, tbl_opd_order就一定沒有資料, 所以要加入, 這裡的挑戰是要加上醫令序
                        // datatable 此時不能使用LINQ查詢
                        List<DataRow> q2 = dtP.Select("CASENO='" + strCASENO + "'").ToList();

                        // 處理tbl_opd_order部分
                        int j = 1;
                        foreach (DataRow dtP_Row in q2)
                        {
                            tbl_opd_order newPr = new tbl_opd_order()
                            {
                                CASENO = strCASENO,
                                uid = (string)dtO_Row["IDNO"],
                                SDATE = temp_date,
                                OD_idx = (byte)(j + 1),
                                rid = (string)dtP_Row["CODE"], //CODE
                                TIMES_DAY = (string)dtP_Row["TIMES_DAY"], //TIMES_DAY
                                METHOD = (string)dtP_Row["METHODE"], //METHOD
                                TIME_QTY1 = (string)dtP_Row["TIME_QTY1"], //TIME_QTY1
                                DAYS = (string)dtP_Row["DAYS"], //DAYS
                                BILL_QTY = (string)dtP_Row["BILL_QTY"], //BILL_QTY
                                HC = (string)dtP_Row["HC"], //HC
                                PRICE = (string)dtP_Row["PRICE"], //PRICE
                                AMT = (string)dtP_Row["AMT"], //AMT
                                CLASS = (string)dtP_Row["CLASS"], //CLASS
                                CHRONIC = (string)dtP_Row["CHRONIC"] //CHRONIC
                            };
                            dc.tbl_opd_order.InsertOnSubmit(newPr);
                            dc.SubmitChanges();
                            j++;
                        }
                    }
                    catch (Exception ex)
                    {
                        log.Error(ex.Message);
                        tb.ShowBalloonTip("錯誤!", ex.Message, BalloonIcon.Error);
                        Logging.Record_error(ex.Message);
                    }
                }
                else    // 資料庫裡已經有了, 檢查是否有異,有異UPDATE
                {
                    // 先處理tbl_opd部分
                    tbl_opd oldOPD = (from p in dc.tbl_opd
                                      where p.CASENO == strCASENO
                                      select p).ToList()[0];     // this is a record
                    string strChange = string.Empty;
                    bool bChange = false;

                    try
                    {
                        string tempstr = string.Empty;
                        if (oldOPD.DEPTNAME != (string)dtO_Row["DEPTNAME"])
                        {
                            strChange += $"改科別: {oldOPD.DEPTNAME} => {dtO_Row["DEPTNAME"]}";
                            bChange = true;
                            oldOPD.DEPTNAME = (string)dtO_Row["DEPTNAME"]; // DEPTNAME
                        }

                        if (oldOPD.DOCTNAME != (string)dtO_Row["DOCTNAME"])
                        {
                            strChange += $"改醫師: {oldOPD.DOCTNAME} => {dtO_Row["DOCTNAME"]}";
                            bChange = true;
                            oldOPD.DOCTNAME = (string)dtO_Row["DOCTNAME"]; //DOCTNAME
                        }

                        if (oldOPD.POSINAME != (string)dtO_Row["POSINAME"])
                        {
                            strChange += $"改身分: {oldOPD.POSINAME} => {dtO_Row["POSINAME"]}";
                            bChange = true;
                            oldOPD.POSINAME = (string)dtO_Row["POSINAME"]; //POSINAME
                        }

                        if (oldOPD.PAYNO != (string)dtO_Row["PAYNO"])
                        {
                            strChange += $"改負擔: {oldOPD.PAYNO} => {dtO_Row["PAYNO"]}";
                            bChange = true;
                            oldOPD.PAYNO = (string)dtO_Row["PAYNO"];  //PAYNO
                        }

                        if (oldOPD.HEATH_CARD != (string)dtO_Row["HEATH_CARD"])
                        {
                            strChange += $"改卡號: {oldOPD.HEATH_CARD} => {dtO_Row["HEATH_CARD"]}";
                            bChange = true;
                            oldOPD.HEATH_CARD = (string)dtO_Row["HEATH_CARD"]; //HEATH_CARD
                        }

                        if (oldOPD.ICDCODE1 != (string)dtO_Row["ICDCODE1"])
                        {
                            strChange += $"改診斷1: {oldOPD.ICDCODE1} => {dtO_Row["ICDCODE1"]}";
                            bChange = true;
                            oldOPD.ICDCODE1 = (string)dtO_Row["ICDCODE1"]; //ICDCODE1
                        }

                        if (oldOPD.ICDCODE2 != (string)dtO_Row["ICDCODE2"])
                        {
                            strChange += $"改診斷2: {oldOPD.ICDCODE2} => {dtO_Row["ICDCODE2"]}";
                            bChange = true;
                            oldOPD.ICDCODE2 = (string)dtO_Row["ICDCODE2"]; //ICDCODE2
                        }

                        if (oldOPD.ICDCODE3 != (string)dtO_Row["ICDCODE3"])
                        {
                            strChange += $"改診斷3: {oldOPD.ICDCODE3} => {dtO_Row["ICDCODE3"]}";
                            bChange = true;
                            oldOPD.ICDCODE3 = (string)dtO_Row["ICDCODE3"]; //ICDCODE3
                        }

                        if (bChange == true)
                        {
                            // 做實改變
                            dc.SubmitChanges();
                            // 做記錄
                            Logging.Record_admin("update opd", $"{strCASENO}: {strChange}");
                        }
                    }
                    catch (Exception ex)
                    {
                        log.Error(strCASENO + ex.Message);
                        tb.ShowBalloonTip("錯誤!", strCASENO + ex.Message, BalloonIcon.Error);
                        Logging.Record_error(strCASENO + ex.Message);
                    }

                    // 再處理tbl_opd_order部分
                    // 先製造兩個list of tbl_opd_order
                    List<Prescription> oldPre = (from d in dc.tbl_opd_order
                                                 where d.CASENO == strCASENO
                                                 orderby d.rid, d.TIMES_DAY
                                                 select new Prescription()
                                                 {
                                                     CASENO = d.CASENO,
                                                     Rid = d.rid,
                                                     TIMES_DAY = d.TIMES_DAY,
                                                     METHOD = d.METHOD,
                                                     TIME_QTY1 = d.TIME_QTY1,
                                                     DAYS = d.DAYS,
                                                     BILL_QTY = d.BILL_QTY,
                                                     HC = d.HC,
                                                     PRICE = d.PRICE,
                                                     AMT = d.AMT,
                                                     CLAS = d.CLASS,
                                                     CHRONIC = d.CHRONIC
                                                 }).ToList();
                    List<Prescription> newPre = new List<Prescription>();
                    List<DataRow> q2 = dtP.Select($"CASENO='{strCASENO}'", "CODE, TIMES_DAY").ToList();
                    // 這個r.count一定大於等於1

                    // 處理tbl_opd_order部分
                    int totalP = q2.Count();
                    for (int j = 0; j < totalP; j++)
                    {
                        Prescription newP = new Prescription()
                        {
                            CASENO = strCASENO,
                            Rid = (string)q2[j]["CODE"], //CODE
                            TIMES_DAY = (string)q2[j]["TIMES_DAY"], //TIMES_DAY
                            METHOD = (string)q2[j]["METHODE"], //METHOD
                            TIME_QTY1 = (string)q2[j]["TIME_QTY1"], //TIME_QTY1
                            DAYS = (string)q2[j]["DAYS"], //DAYS
                            BILL_QTY = (string)q2[j]["BILL_QTY"], //BILL_QTY
                            HC = (string)q2[j]["HC"], //HC
                            PRICE = (string)q2[j]["PRICE"], //PRICE
                            AMT = (string)q2[j]["AMT"], //AMT
                            CLAS = (string)q2[j]["CLASS"], //CLASS
                            CHRONIC = (string)q2[j]["CHRONIC"] //CHRONIC
                        };
                        newPre.Add(newP);
                    }
                    // Now we have 2 lists now, but lists are only references
                    // 先比較兩者是否相同, 相同則跳下一筆
                    string strT = Exact(oldPre, newPre);
                    if (strT.Length != 0) // "" stands for identical
                    {
                        // 若不同則找出哪裡不同, 記錄下來
                        Logging.Record_admin("update opd order", $"{strCASENO}: {strT}");
                        // 最後把舊的刪掉, 插入新的

                        // 刪掉舊的
                        var q3 = from p in dc.tbl_opd_order
                                 where p.CASENO == strCASENO
                                 select p;
                        foreach (tbl_opd_order pr in q3)
                        {
                            dc.tbl_opd_order.DeleteOnSubmit(pr);
                        }
                        dc.SubmitChanges();
                        // 插入新的
                        // datatable 此時不能使用LINQ查詢
                        List<DataRow> q4 = dtP.Select($"CASENO='{strCASENO}'").ToList();

                        // 處理tbl_opd_order部分
                        int totalPr = q4.Count;
                        for (int j = 0; j < totalPr; j++)
                        {
                            tbl_opd_order newPr = new tbl_opd_order()
                            {
                                CASENO = strCASENO,
                                uid = oldOPD.uid,
                                SDATE = oldOPD.SDATE,
                                OD_idx = (byte)(j + 1),
                                rid = (string)q4[j]["CODE"], //CODE
                                TIMES_DAY = (string)q4[j]["TIMES_DAY"], //TIMES_DAY
                                METHOD = (string)q4[j]["METHODE"], //METHOD
                                TIME_QTY1 = (string)q4[j]["TIME_QTY1"], //TIME_QTY1
                                DAYS = (string)q4[j]["DAYS"], //DAYS
                                BILL_QTY = (string)q4[j]["BILL_QTY"], //BILL_QTY
                                HC = (string)q4[j]["HC"], //HC
                                PRICE = (string)q4[j]["PRICE"], //PRICE
                                AMT = (string)q4[j]["AMT"], //AMT
                                CLASS = (string)q4[j]["CLASS"], //CLASS
                                CHRONIC = (string)q4[j]["CHRONIC"] //CHRONIC
                            };
                            dc.tbl_opd_order.InsertOnSubmit(newPr);
                            dc.SubmitChanges();
                        }
                    }
                }
            }
            // 這樣的add opd沒什麼用
            //        Record_adm("add opd", dtO.TableName)
            dtO.Dispose();
            dtP.Dispose();
            ds.Dispose();

            #endregion 進行讀取資料
        }

        private string Exact(List<Prescription> oldPr, List<Prescription> newPr)
        {
            List<Prescription> oNon = new List<Prescription>();
            List<Prescription> nNoo = new List<Prescription>();

            foreach (Prescription oP in oldPr)
            {
                bool if_identical = false;
                foreach (Prescription nP in newPr)
                {
                    if ((oP.Rid == nP.Rid) && (oP.TIMES_DAY == nP.TIMES_DAY) && (oP.METHOD == nP.METHOD) &&
                       (oP.TIME_QTY1 == nP.TIME_QTY1) && (oP.DAYS == nP.DAYS) && (oP.BILL_QTY == nP.BILL_QTY) &&
                       (oP.HC == nP.HC) && (oP.PRICE == nP.PRICE) && (oP.CLAS == nP.CLAS) && (oP.CHRONIC == nP.CHRONIC))
                    {
                        if_identical = true;
                        break;
                    }
                }
                if (!if_identical) oNon.Add(oP);
            }

            foreach (Prescription nP in newPr)
            {
                bool if_identical = false;
                foreach (Prescription oP in oldPr)
                {
                    if ((nP.Rid == oP.Rid) && (nP.TIMES_DAY == oP.TIMES_DAY) && (nP.METHOD == oP.METHOD) &&
                       (nP.TIME_QTY1 == oP.TIME_QTY1) && (nP.DAYS == oP.DAYS) && (nP.BILL_QTY == oP.BILL_QTY) &&
                       (nP.HC == oP.HC) && (nP.PRICE == oP.PRICE) && (nP.CLAS == oP.CLAS) && (nP.CHRONIC == oP.CHRONIC))
                    {
                        if_identical = true;
                        break;
                    }
                }
                if (!if_identical) nNoo.Add(nP);
            }

            string output = string.Empty;
            if ((oNon.Count == 0) && (nNoo.Count == 0))
            {
                return output;
            }
            else
            {
                foreach (Prescription a in oNon)
                {
                    output += $"DC: {Display_by_code(a)}";
                }
                foreach (Prescription b in nNoo)
                {
                    output += $"Add: {Display_by_code(b)} ";
                }
                return output;
            }
        }

        private string Display_by_code(Prescription pr)
        {
            string strReturn = string.Empty;
            if (pr.CLAS == "藥品")
            {
                // CODE, rid
                strReturn += $"{pr.Rid}, ";
                // TIME_QTY1
                strReturn += $"{pr.TIME_QTY1}# ";
                // TIMES_DAY
                strReturn += $"{pr.TIMES_DAY} ";
                // METHOD
                strReturn += $"{pr.METHOD} x";
                // DAYS
                strReturn += $"{pr.DAYS}D; ";
            }
            else
            {
                strReturn += $"{pr.Rid}; ";
            }
            return strReturn;
        }
    }
}