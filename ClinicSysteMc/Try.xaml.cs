using ClinicSysteMc.ViewModel.Converters;
using System.Windows;

namespace ClinicSysteMc
{
    /// <summary>
    /// Try.xaml 的互動邏輯
    /// </summary>
    public partial class Try : Window
    {
        public Try()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            // 測試Dashboard
            Dash d = new Dash("11007");
            d.Show();
        }
    }
}
