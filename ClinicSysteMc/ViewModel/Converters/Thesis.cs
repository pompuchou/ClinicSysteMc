using AutoIt;
using System.Diagnostics;

namespace ClinicSysteMc.ViewModel.Converters
{
    internal class Thesis
    {
        public static void LogIN()
        {
            Process[] isClud = Process.GetProcessesByName("THCludSuit");  // 主目錄
            // 測試"看診清單"是否有打開
            if (isClud.Length == 0)
            {
                AutoItX.Run(@"C:\Program Files (x86)\THESE\杏雲醫療資訊系統\THCloudStarter.exe", @"C:\Program Files (x86)\THESE\杏雲醫療資訊系統\");

                //; Wait for the Notepad to become active. The classname "Notepad" Is monitored instead of the window title
                AutoItX.WinWaitActive("登入畫面");

                //; Now that the Notepad window Is active type some text
                if (AutoItX.ControlGetText("登入畫面", "", "[NAME:txtHospitalExtensionCode]") != "A")
                {
                    AutoItX.ControlClick("登入畫面", "", "[NAME:txtHospitalExtensionCode]", "LEFT", 2);
                    AutoItX.ControlSend("登入畫面", "", "[NAME:txtHospitalExtensionCode]", "A");
                }
                AutoItX.ControlSend("登入畫面", "", "[NAME:txtPassword]", "IlovePierce4926");
                AutoItX.Sleep(500);
                // [NAME:btnLogin]
                //AutoItX.ControlClick("登入畫面", "", "[NAME:picLogin]");
                AutoItX.ControlClick("登入畫面", "", "[NAME:btnLogin]");

                AutoItX.WinActivate("杏雲雲端醫療服務");
                AutoItX.WinWaitActive("杏雲雲端醫療服務");
                AutoItX.Sleep(2000);
            }
            else
            {
                AutoItX.WinActivate("杏雲雲端醫療服務");
            }
        }
    }
}
