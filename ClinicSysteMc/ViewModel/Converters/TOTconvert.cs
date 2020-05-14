using ClinicSysteMc.Model;
using Hardcodet.Wpf.TaskbarNotification;
using System;
using System.Linq;
using System.Xml;

namespace ClinicSysteMc.ViewModel.Converters
{
    internal class TOTconvert : IDisposable
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private readonly TaskbarIcon tb = new TaskbarIcon();
        private readonly string _loadpath;
        private bool _disposed = false;

        public TOTconvert(string loadpath)
        {
            _loadpath = loadpath;
        }

        public void Transform()
        {
            #region 宣告

            XmlDocument xdoc = new XmlDocument();
            XmlNode xOutpatient;   // root of the xml
            XmlNodeList xTDATA;    // 用來放tDATA
            XmlNodeList xDDATA;    // 用來放dDATA
            XmlNode xNodeTemp;     // 臨時的xml node操作
            string keyT3;    // 當key值, 費用年月
            string keyD1;    // 當key值, 案件分類
            string keyD2;    // 當key值,流水編號
            int dN = 0;      // dData筆數, 即當月看診人次數
            int pN = 0;      // pData筆數, 即當月處方筆數

            //讀取XML
            // 20190615 revisited: 原本想說建立防呆機制, 結果一看才發現有天然的防呆機制,就是primary key的設置, 三個表都有,有重複值自然就不會讀了
            xdoc.Load(_loadpath);
            // root node就是outpatient, outpatient下面就是兩個node: tdata,
            xOutpatient = (XmlNode)xdoc.DocumentElement;
            //選擇section
            xTDATA = xOutpatient.SelectNodes("tdata");   //這應該只有一個
            xDDATA = xOutpatient.SelectNodes("ddata");   //這應該有很多個

            #endregion 宣告

            #region 讀取tdata

            try
            {
                //TDATA只有一個item
                xNodeTemp = xTDATA.Item(0);
                //這個唯一的item下面有42個child node/item, 就是總表了
                //20190527 我已經搞懂總表了,可以寫入SQL了
                //以下寫入SQL

                // 總表重複12次, xml_tdata
                // 宣告新的一行
                keyT3 = xNodeTemp.SelectSingleNode("t3").InnerText;
                // 如果已經曾經匯入, 就不要重複匯入, t3是key值, 不可重複
                using (CSDataContext dc = new CSDataContext())
                {
                    var q = from p in dc.xml_tdata
                            where p.t3 == keyT3
                            select p;
                    if (q.Count() !=0)
                    {
                        throw new Exception($"{keyT3} 之前已經匯入, 不可重複匯入!");
                    }

                }


                xml_tdata newT = new xml_tdata()
                {
                    t1 = xNodeTemp.SelectSingleNode("t1").InnerText,
                    t2 = xNodeTemp.SelectSingleNode("t2").InnerText,
                    t3 = xNodeTemp.SelectSingleNode("t3").InnerText,
                    t4 = char.Parse(xNodeTemp.SelectSingleNode("t4").InnerText),
                    t5 = char.Parse(xNodeTemp.SelectSingleNode("t5").InnerText),
                    t6 = xNodeTemp.SelectSingleNode("t6").InnerText,
                    t37 = int.Parse(xNodeTemp.SelectSingleNode("t37").InnerText),
                    t38 = int.Parse(xNodeTemp.SelectSingleNode("t38").InnerText),
                    t39 = int.Parse(xNodeTemp.SelectSingleNode("t39").InnerText),
                    t40 = int.Parse(xNodeTemp.SelectSingleNode("t40").InnerText),
                    t41 = xNodeTemp.SelectSingleNode("t41").InnerText,
                    t42 = xNodeTemp.SelectSingleNode("t42").InnerText
                };
                //20190527 完成
                using (CSDataContext dc = new CSDataContext())
                {
                    dc.xml_tdata.InsertOnSubmit(newT);
                    dc.SubmitChanges();
                    string output = $"匯入健保申報檔年月: {keyT3}";
                    Logging.Record_admin("匯入健保申報檔", output);
                    log.Info(output);
                }
            }
            catch (Exception ex)
            {
                string output = $"讀取tdata錯誤: {ex.Message}";
                Logging.Record_error(output);
                log.Error(output);
                tb.ShowBalloonTip("錯誤!", output, BalloonIcon.Error);
                return;
            }

            #endregion 讀取tdata

            #region 讀取ddata, pdata

            //=====================================================================================
            //2019/5/27 完成的
            //DDATA有很多個item
            foreach (XmlNode dNode in xDDATA)
            {
                //取得節點[dhead]
                xNodeTemp = dNode.SelectSingleNode("dhead");
                keyD1 = xNodeTemp.SelectSingleNode("d1").InnerText;
                keyD2 = xNodeTemp.SelectSingleNode("d2").InnerText;

                //取得節點[dbody]
                xNodeTemp = dNode.SelectSingleNode("dbody");

                //取得ddata, 下面應有dhead, dbody兩個node, dhead下有d1, d2, dbody下有30欄位
                //------------> as ddata
                try
                {
                    // 宣告新的一行
                    xml_ddata newD = new xml_ddata()
                    {
                        t3 = keyT3,
                        d1 = keyD1,
                        d2 = int.Parse(keyD2),
                        d3 = xNodeTemp.SelectSingleNode("d3")?.InnerText,
                        d4 = xNodeTemp.SelectSingleNode("d4")?.InnerText,
                        d8 = xNodeTemp.SelectSingleNode("d8")?.InnerText,
                        d9 = xNodeTemp.SelectSingleNode("d9")?.InnerText,
                        d11 = xNodeTemp.SelectSingleNode("d11")?.InnerText,
                        d15 = xNodeTemp.SelectSingleNode("d15")?.InnerText,
                        d16 = xNodeTemp.SelectSingleNode("d16")?.InnerText,
                        d17 = xNodeTemp.SelectSingleNode("d17")?.InnerText,
                        d18 = char.Parse(xNodeTemp.SelectSingleNode("d18")?.InnerText),
                        d19 = xNodeTemp.SelectSingleNode("d19")?.InnerText,
                        d20 = xNodeTemp.SelectSingleNode("d20")?.InnerText,
                        d21 = xNodeTemp.SelectSingleNode("d21")?.InnerText,
                        d22 = xNodeTemp.SelectSingleNode("d22")?.InnerText,
                        d23 = xNodeTemp.SelectSingleNode("d23")?.InnerText,
                        d28 = char.Parse(xNodeTemp.SelectSingleNode("d28")?.InnerText),
                        d29 = xNodeTemp.SelectSingleNode("d29")?.InnerText,
                        d30 = xNodeTemp.SelectSingleNode("d30")?.InnerText,
                        d32 = int.Parse(xNodeTemp.SelectSingleNode("d32")?.InnerText),
                        d33 = int.Parse(xNodeTemp.SelectSingleNode("d33")?.InnerText),
                        d34 = int.Parse(xNodeTemp.SelectSingleNode("d34")?.InnerText),
                        d35 = xNodeTemp.SelectSingleNode("d35")?.InnerText,
                        d36 = int.Parse(xNodeTemp.SelectSingleNode("d36")?.InnerText),
                        d39 = int.Parse(xNodeTemp.SelectSingleNode("d39")?.InnerText),
                        d40 = int.Parse(xNodeTemp.SelectSingleNode("d40")?.InnerText),
                        d41 = int.Parse(xNodeTemp.SelectSingleNode("d41")?.InnerText),
                        d49 = xNodeTemp.SelectSingleNode("d49")?.InnerText
                    };
                    if (int.TryParse(xNodeTemp.SelectSingleNode("d27")?.InnerText, out int i27)) newD.d27 = i27;
                    if (!string.IsNullOrEmpty(xNodeTemp.SelectSingleNode("d14")?.InnerText)) newD.d14 = char.Parse(xNodeTemp.SelectSingleNode("d14").InnerText);

                    using (CSDataContext dc = new CSDataContext())
                    {
                        dc.xml_ddata.InsertOnSubmit(newD);
                        dc.SubmitChanges();
                        dN++;
                    }
                }
                catch (Exception ex)
                {
                    string output = $"讀取ddata錯誤: {ex.Message}";
                    Logging.Record_error(output);
                    log.Error(output);
                    tb.ShowBalloonTip("錯誤!", output, BalloonIcon.Error);
                    continue;
                }
                //取得[dbody]下的節點[pdata],這可能有很多個,也可能沒有半個, 要有個if, 要有個迴圈for next
                if (xNodeTemp.SelectNodes("pdata").Count != 0)
                {
                    XmlNodeList xPDATA = xNodeTemp.SelectNodes("pdata");
                    foreach (XmlNode pNode in xPDATA)
                    {
                        try
                        {
                            xml_pdata newP = new xml_pdata()
                            {
                                t3 = keyT3,
                                d1 = keyD1,
                                d2 = int.Parse(keyD2),
                                p4 = pNode.SelectSingleNode("p4")?.InnerText,
                                p6 = pNode.SelectSingleNode("p6").InnerText,
                                p7 = pNode.SelectSingleNode("p7").InnerText,
                                p8 = double.Parse(pNode.SelectSingleNode("p8").InnerText),
                                p9 = pNode.SelectSingleNode("p9").InnerText,
                                p10 = (int)float.Parse(pNode.SelectSingleNode("p10")?.InnerText),
                                p11 = (int)float.Parse(pNode.SelectSingleNode("p11")?.InnerText),
                                p12 = int.Parse(pNode.SelectSingleNode("p12")?.InnerText),
                                p13 = int.Parse(pNode.SelectSingleNode("p13")?.InnerText),
                                p14 = pNode.SelectSingleNode("p14")?.InnerText,
                                p15 = pNode.SelectSingleNode("p15")?.InnerText,
                                p16 = pNode.SelectSingleNode("p16")?.InnerText,
                                p20 = pNode.SelectSingleNode("p20")?.InnerText
                            };
                            if (int.TryParse(pNode.SelectSingleNode("p1")?.InnerText, out int i1)) newP.p1 = i1;
                            if (!string.IsNullOrEmpty(pNode.SelectSingleNode("p2")?.InnerText)) newP.p2 = char.Parse(pNode.SelectSingleNode("p2").InnerText);
                            if (!string.IsNullOrEmpty(pNode.SelectSingleNode("p3")?.InnerText)) newP.p3 = char.Parse(pNode.SelectSingleNode("p3")?.InnerText);
                            if (double.TryParse(pNode.SelectSingleNode("p5")?.InnerText, out double i5)) newP.p5 = i5;
                            if (!string.IsNullOrEmpty(pNode.SelectSingleNode("p17")?.InnerText)) newP.p17 = char.Parse(pNode.SelectSingleNode("p17")?.InnerText);

                            using (CSDataContext dc = new CSDataContext())
                            {
                                dc.xml_pdata.InsertOnSubmit(newP);
                                dc.SubmitChanges();
                                pN++;
                            }
                        }
                        catch (Exception ex)
                        {
                            string output = $"讀取pdata錯誤: {ex.Message}";
                            Logging.Record_error(output);
                            log.Error(output);
                            tb.ShowBalloonTip("錯誤!", output, BalloonIcon.Error);
                            continue;
                        }
                    }
                }
            }

            #endregion 讀取ddata, pdata

            #region 進行配對

            //20190615 連結tbl_opd
            using (CSDataContext dc = new CSDataContext())
            {
                var q = (from cs in dc.sp_match_xml().AsEnumerable()
                         select cs).First();
                int n = q.rows_affected;
                string output = $"健保上傳XML檔配對{n}筆配對成功, 共{dN}看診人次, {pN}處方";
                Logging.Record_admin("健保上傳XML檔配對", output);
                log.Info(output);
                tb.ShowBalloonTip("匯入成功!", output, BalloonIcon.Info);
            }

            #endregion 進行配對

            this.Dispose();
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (_disposed)
            {
                return;
            }

            if (disposing)
            {
                // Free any other managed objects here.
            }

            _disposed = true;
        }
    }
}