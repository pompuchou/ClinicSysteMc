using Hardcodet.Wpf.TaskbarNotification;
using System.Threading;
using System.Windows;

namespace ClinicSysteMc.ViewModel.Converters
{
    /// <summary>
    /// Dash.xaml 的互動邏輯
    /// 非採MVVM設計
    /// 20210722創立
    /// </summary>
    public partial class Dash : Window
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private readonly TaskbarIcon tb = new TaskbarIcon();
        private readonly string _strYM;
        private bool _stopflag;

        public Dash(string YM)
        {
            InitializeComponent();
            _strYM = YM;
            _stopflag = false;
        }

        private async void Button_Click(object sender, RoutedEventArgs e)
        {
            log.Info("1. Start button pressed.");
            await StoryBoard();
            log.Info("10. Ended.");
            Thread.Sleep(3000);
            this.Close();
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            // 沒有Async根本不可能中斷
            // 有Async為何不能winactivate
            _stopflag = true;
        }
    }
}
