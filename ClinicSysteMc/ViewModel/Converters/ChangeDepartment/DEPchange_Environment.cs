using AutoIt;
using System;

namespace ClinicSysteMc.ViewModel.Converters
{
    internal partial class DEPchange
    {
        private void Environment()
        {
            #region Environment
            try
            {
                // 營造環境
                if (AutoItX.WinExists("看診清單") == 1) //如果直接存在就直接叫用
                {
                    AutoItX.WinActivate("看診清單");
                }
                else
                {
                    Thesis.LogIN();
                    // 打開"看診清單"
                    AutoItX.Run(@"C:\Program Files (x86)\THESE\杏雲醫療資訊系統\THCClinic.exe", @"C:\Program Files (x86)\THESE\杏雲醫療資訊系統\");
                    AutoItX.WinWaitActive("看診清單");
                    AutoItX.WinActivate("看診清單");
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
