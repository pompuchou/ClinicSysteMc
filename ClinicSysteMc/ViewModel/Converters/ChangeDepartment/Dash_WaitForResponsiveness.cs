using AutoIt;
using System;
using System.Threading;

namespace ClinicSysteMc.ViewModel.Converters
{
    public partial class Dash
    {
        internal void WaitForResponsiveness(string WinName, string CtrlName)
        {
            // 找到方法, 分辨responsiveness, active不代表responsiveness
            IntPtr w = AutoItX.WinGetHandle(WinName);
            int n = 0;
            log.Info($"Now in {WinName}.");

            while (AutoItX.ControlGetHandle(w, CtrlName) == IntPtr.Zero)
            {
                Thread.Sleep(100);
                n++;
            };
            log.Debug($"Wait for {n * 100} msec in {WinName}.");
            // 能夠找到control的pointer那就代表有responsiveness
        }
    }
}
