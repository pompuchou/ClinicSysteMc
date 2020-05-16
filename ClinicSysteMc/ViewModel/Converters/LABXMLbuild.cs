using Microsoft.Win32;
using System.Xml;

namespace ClinicSysteMc.ViewModel.Converters
{
    internal class LABXMLbuild
    {
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


        }
    }
}
