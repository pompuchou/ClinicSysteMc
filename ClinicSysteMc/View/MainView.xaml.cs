using System.Deployment.Application;
using System.Windows;

[assembly: log4net.Config.XmlConfigurator(Watch = true)]

namespace ClinicSysteMc.View
{
    /// <summary>
    /// MainWindow.xaml 的互動邏輯
    /// </summary>

    public partial class MainView : Window
    {
        public MainView()
        {
            InitializeComponent();
            string version;
            try
            {
                //// get deployment version
                version = ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString();
            }
            catch (InvalidDeploymentException)
            {
                //// you cannot read publish version when app isn't installed
                //// (e.g. during debug)
                version = "debugging, not installed";
            }
            this.Title += $" v.{version}";
        }
    }
}