using AutoIt;

namespace ClinicSysteMc.ViewModel.Converters
{
    public partial class Dash
    {
        private void Finishing()
        {
            AutoItX.WinActivate("問診畫面");

            int idx = 0;
            // 先問「此病患已批價, 是否繼續?」, THCClinic, 還可能問重大傷病, 超過8種藥物, 可能有重複用藥畫面
            do
            {
                AutoItX.Send("{F9}");
                AutoItX.Sleep(100);
                idx++; // time out for 10 sec at most
            } while (AutoItX.WinExists("THCClinic") == 0 && idx < 100);

            // 下一個畫面「確定要重複開立收據」,
            // 一定會有「已經批價」, 可能會有「重大傷病身分」, 可能會有「八種以上藥物]
            // 可能會有「跨院重複開立醫囑提示」
            idx = 0;
            do
            {
                if (AutoItX.WinExists("THCClinic") == 1) AutoItX.ControlClick("THCClinic", "", "[CLASSNN:Button1]");
                if (AutoItX.WinExists("診間完成檢核") == 1) AutoItX.ControlClick("診間完成檢核", "", "[NAME:Button_2]");
                if (AutoItX.WinExists("跨院重複開立醫囑提示") == 1) AutoItX.ControlClick("跨院重複開立醫囑提示", "", "[NAME:OK_Button]");
                AutoItX.Sleep(100);
                idx++; // time out for 10 sec at most
            } while (AutoItX.WinExists("These.CludCln.Accounting") == 0 && idx < 100);

            // 是否重印收據
            AutoItX.WinWaitActive("These.CludCln.Accounting");
            AutoItX.ControlClick("These.CludCln.Accounting", "", "[CLASSNN:Button2]");
            AutoItX.Sleep(1000);
        }
    }
}
