using System;
using System.Windows;
using System.Windows.Input;
using System.Xml;
using Microsoft.VisualBasic;

namespace ClinicSysteMc.ViewModel.Commands
{
    internal class UploadLab : ICommand
    {
        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        public bool CanExecute(object parameter)
        {
            return true;
        }

        public void Execute(object parameter)
        {
            #region Declaration

            XmlDocument xdoc;     // TOTFA.xml
            XmlElement xElement;    // patient
            XmlElement xChildElement;
            XmlElement xElement2;
            XmlElement xChildElement2;
            string savepath = string.Empty;
            string strYM = $"{DateTime.Now.Year - 1911}{(DateTime.Now.Month + 100).ToString().Substring(1)}";

            #endregion Declaration

            #region ASK for YM

            strYM = Interaction.InputBox("請輸入費用年月", "詢問", strYM);
            if (!int.TryParse(strYM, out _) || strYM.Length != 5)
                {
                MessageBox.Show("格式錯誤");
                return;
            }
            else if (int.Parse(strYM.Substring(3)) < 1 || int.Parse(strYM.Substring(3)) > 12)
            {
                MessageBox.Show("格式錯誤");
                return;
            }

            #endregion ASK for YM
        }
    }
}