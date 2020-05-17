using ClinicSysteMc.Model;
using Hardcodet.Wpf.TaskbarNotification;
using Microsoft.Win32;
using System.Linq;
using System.Xml;
using System.IO.Compression;
using System.IO;

namespace ClinicSysteMc.ViewModel.Converters
{
    internal class LABXMLbuild
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private readonly TaskbarIcon tb = new TaskbarIcon();

        private readonly string _strYM;

        public LABXMLbuild(string YM)
        {
            _strYM = YM;
        }

        public void Build()
        {
            #region Declaration

            XmlDocument xdoc;     // TOTFA.xml
            XmlElement xElement;    // patient
            XmlElement xChildElement;
            XmlElement xElement2;
            XmlElement xChildElement2;
            string savepath;

            #endregion Declaration

            #region 寫入檔案路徑

            // 讀取要輸入的位置
            // 從杏翔病患資料輸入, 只有一種xml格式
            // Xml格式的index=2
            SaveFileDialog sFDialog = new SaveFileDialog
            {
                Filter = "xml|*.xml",
                FileName = "TOTFA.xml"
            };

            if (sFDialog.ShowDialog() == true)
            {
                savepath = sFDialog.FileName;
            }
            else
            {
                // 取消, 什麼也沒有做
                return;
            }

            #endregion 寫入檔案路徑

            try
            {
                CSDataContext dc = new CSDataContext();
                var q = from pt in dc.sp_get_hdata(_strYM).AsEnumerable()
                        select pt;
                // 建立一個 XmlDocument 物件並加入 Declaration
                xdoc = new XmlDocument();
                xdoc.AppendChild(xdoc.CreateXmlDeclaration("1.0", "big5", ""));
                // 建立根節點物件並加入 XmlDocument 中 (第0層)
                xElement = xdoc.CreateElement("patient");
                xChildElement = xElement; // 這個舉動毫無意義,但可以避免錯誤訊息
                xdoc.AppendChild(xElement);
                // 在sections下寫入一個節點名稱為section(第1層)

                foreach (var p in q)
                {
                    if (p.r1 == 1)
                    {
                        xChildElement = xdoc.CreateElement("hdata");
                        xElement.AppendChild(xChildElement);     // patient下加個hdata
                        // 第2層節點
                        xElement2 = xdoc.CreateElement("h1");
                        xElement2.InnerText = p.h1;               // h1 報告類別, 1:檢體檢驗報告
                        xChildElement.AppendChild(xElement2);     // hdata下加個h1
                        xElement2 = xdoc.CreateElement("h2");
                        xElement2.InnerText = p.h2;              // h2 醫事機構代碼
                        xChildElement.AppendChild(xElement2);    // hdata下加個h2
                        xElement2 = xdoc.CreateElement("h3");
                        xElement2.InnerText = p.h3;              // h3 醫事類別, 11:門診西醫診所
                        xChildElement.AppendChild(xElement2);    // hdata下加個h3
                        xElement2 = xdoc.CreateElement("h4");
                        xElement2.InnerText = p.h4;             // h4 費用年月
                        xChildElement.AppendChild(xElement2);    // hdata下加個h4
                        xElement2 = xdoc.CreateElement("h5");
                        xElement2.InnerText = p.h5;               // h5 申報類別, 1:送核
                        xChildElement.AppendChild(xElement2);    // hdata下加個h5
                        xElement2 = xdoc.CreateElement("h6");
                        xElement2.InnerText = p.h6;               // h5 申報類別, 1:送核
                        xChildElement.AppendChild(xElement2);    // hdata下加個h5
                        xElement2 = xdoc.CreateElement("h7");
                        xElement2.InnerText = p.h7;               // h5 申報類別, 1:送核
                        xChildElement.AppendChild(xElement2);    // hdata下加個h5
                        xElement2 = xdoc.CreateElement("h8");
                        xElement2.InnerText = p.h8.ToString();               // h5 申報類別, 1:送核
                        xChildElement.AppendChild(xElement2);    // hdata下加個h5
                        xElement2 = xdoc.CreateElement("h9");
                        xElement2.InnerText = p.h9;               // h5 申報類別, 1:送核
                        xChildElement.AppendChild(xElement2);    // hdata下加個h5
                        xElement2 = xdoc.CreateElement("h10");
                        xElement2.InnerText = p.h10;               // h5 申報類別, 1:送核
                        xChildElement.AppendChild(xElement2);    // hdata下加個h5
                        xElement2 = xdoc.CreateElement("h11");
                        xElement2.InnerText = p.h11;               // h5 申報類別, 1:送核
                        xChildElement.AppendChild(xElement2);    // hdata下加個h5
                        xElement2 = xdoc.CreateElement("h17");
                        xElement2.InnerText = p.h17.ToString();               // h5 申報類別, 1:送核
                        xChildElement.AppendChild(xElement2);    // hdata下加個h5
                        xElement2 = xdoc.CreateElement("h18");
                        xElement2.InnerText = p.h18;               // h5 申報類別, 1:送核
                        xChildElement.AppendChild(xElement2);    // hdata下加個h5
                        xElement2 = xdoc.CreateElement("h19");
                        xElement2.InnerText = p.h19;               // h5 申報類別, 1:送核
                        xChildElement.AppendChild(xElement2);    // hdata下加個h5
                        xElement2 = xdoc.CreateElement("h20");
                        xElement2.InnerText = p.h20;               //h5 申報類別, 1:送核
                        xChildElement.AppendChild(xElement2);    // hdata下加個h5
                        xElement2 = xdoc.CreateElement("h22");
                        xElement2.InnerText = p.h22;               // h5 申報類別, 1:送核
                        xChildElement.AppendChild(xElement2);    // hdata下加個h5
                        xElement2 = xdoc.CreateElement("h23");
                        xElement2.InnerText = p.h23;               // h5 申報類別, 1:送核
                        xChildElement.AppendChild(xElement2);    // hdata下加個h5
                        xElement2 = xdoc.CreateElement("h25");
                        xElement2.InnerText = p.h25;               // h5 申報類別, 1:送核
                        xChildElement.AppendChild(xElement2);    // hdata下加個h5
                        xElement2 = xdoc.CreateElement("h26");
                        xElement2.InnerText = p.h26;               // h5 申報類別, 1:送核
                        xChildElement.AppendChild(xElement2);    // hdata下加個h5
                        // 第3層節點
                        xChildElement2 = xdoc.CreateElement("rdata"); // rdata
                        xChildElement.AppendChild(xChildElement2);   // under hdata add rdata
                        xElement2 = xdoc.CreateElement("r1");
                        xElement2.InnerText = p.r1.ToString();
                        xChildElement2.AppendChild(xElement2);
                        xElement2 = xdoc.CreateElement("r2");
                        xElement2.InnerText = p.r2;
                        xChildElement2.AppendChild(xElement2);
                        xElement2 = xdoc.CreateElement("r3");
                        xElement2.InnerText = p.r3;
                        xChildElement2.AppendChild(xElement2);
                        xElement2 = xdoc.CreateElement("r4");
                        xElement2.InnerText = p.r4;
                        xChildElement2.AppendChild(xElement2);
                        xElement2 = xdoc.CreateElement("r5");
                        xElement2.InnerText = p.r5;
                        xChildElement2.AppendChild(xElement2);
                        xElement2 = xdoc.CreateElement("r6-1");
                        xElement2.InnerText = p.r6a;
                        xChildElement2.AppendChild(xElement2);
                        xElement2 = xdoc.CreateElement("r6-2");
                        xElement2.InnerText = p.r6b;
                        xChildElement2.AppendChild(xElement2);
                        xElement2 = xdoc.CreateElement("r9");
                        xElement2.InnerText = p.r9;
                        xChildElement2.AppendChild(xElement2);
                        xElement2 = xdoc.CreateElement("r10");
                        xElement2.InnerText = p.r10;
                        xChildElement2.AppendChild(xElement2);
                    }
                    else
                    {
                        // 第3層節點
                        xChildElement2 = xdoc.CreateElement("rdata"); //rdata
                        xChildElement.AppendChild(xChildElement2);   //under hdata add rdata
                        xElement2 = xdoc.CreateElement("r1");
                        xElement2.InnerText = p.r1.ToString();
                        xChildElement2.AppendChild(xElement2);
                        xElement2 = xdoc.CreateElement("r2");
                        xElement2.InnerText = p.r2;
                        xChildElement2.AppendChild(xElement2);
                        xElement2 = xdoc.CreateElement("r3");
                        xElement2.InnerText = p.r3;
                        xChildElement2.AppendChild(xElement2);
                        xElement2 = xdoc.CreateElement("r4");
                        xElement2.InnerText = p.r4;
                        xChildElement2.AppendChild(xElement2);
                        xElement2 = xdoc.CreateElement("r5");
                        xElement2.InnerText = p.r5;
                        xChildElement2.AppendChild(xElement2);
                        xElement2 = xdoc.CreateElement("r6-1");
                        xElement2.InnerText = p.r6a;
                        xChildElement2.AppendChild(xElement2);
                        xElement2 = xdoc.CreateElement("r6-2");
                        xElement2.InnerText = p.r6b;
                        xChildElement2.AppendChild(xElement2);
                        xElement2 = xdoc.CreateElement("r9");
                        xElement2.InnerText = p.r9;
                        xChildElement2.AppendChild(xElement2);
                        xElement2 = xdoc.CreateElement("r10");
                        xElement2.InnerText = p.r10;
                        xChildElement2.AppendChild(xElement2);
                    }
                }
                xdoc.Save(savepath);
                using (FileStream fs = new FileStream(savepath.Replace(".xml", ".zip"), FileMode.Create))
                using (ZipArchive arch = new ZipArchive(fs, ZipArchiveMode.Create))
                {
                    arch.CreateEntryFromFile(savepath, "TOTFA.xml");
                }
                string output = $"{_strYM} 檢驗上傳XML製作完成.";
                log.Info(output);
                tb.ShowBalloonTip("完成!", $"{output}\r\n請上傳檔案.", BalloonIcon.Info);
                Logging.Record_admin("製作檢驗上傳檔案", output);
            }
            catch (System.Exception ex)
            {
                string output = $"{ex.Message} \r\n{ex.StackTrace}";
                log.Error(output);
                tb.ShowBalloonTip("錯誤!", output, BalloonIcon.Error);
            }
        }
    }
}