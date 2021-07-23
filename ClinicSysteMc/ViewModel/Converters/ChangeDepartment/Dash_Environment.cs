using AutoIt;
using System;

namespace ClinicSysteMc.ViewModel.Converters
{
    public partial class Dash
    {
        public static void Environment()
        {
            #region Environment
            try
            {
                // 營造環境
                if (AutoItX.WinExists("看診清單") == 1) //如果直接存在就直接叫用
                {
                    // 有問診畫面就關掉
                    if (AutoItX.WinExists("問診畫面") == 1)
                    {
                        AutoItX.WinClose("問診畫面");
                    }
                    //AutoItX.WinWaitActive("看診清單");
                    //AutoItX.WinActivate("看診清單");
                    //Async之後就永遠等不到?
                }
                else
                {
                    Thesis.LogIN();
                    // 打開"看診清單"
                    AutoItX.Run(@"C:\Program Files (x86)\THESE\杏雲醫療資訊系統\THCClinic.exe", @"C:\Program Files (x86)\THESE\杏雲醫療資訊系統\");
                    //AutoItX.WinWaitActive("看診清單");
                    //AutoItX.WinActivate("看診清單");
                    //Async之後就永遠等不到?
                }
            }
            catch (Exception ex)
            {
                string o = $"Something wrong in Environment: {ex.Message}";
                Logging.Record_error(o);
                log.Error(o);
                return;
            }
            #endregion Environment
        }
    }
}
