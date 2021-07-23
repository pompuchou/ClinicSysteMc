using Hardcodet.Wpf.TaskbarNotification;
using System.Threading;

namespace ClinicSysteMc.ViewModel.Converters
{
    public partial class Dash
    {
        private void Summerize(int idx)
        {
            Logging.Record_admin("change department", $"修改了{idx}筆");
            tb.ShowBalloonTip("完成", $"修改了{idx}筆, 請匯入門診資料.", BalloonIcon.Info);
        }
    }
}
