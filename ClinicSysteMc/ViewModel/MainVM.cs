using ClinicSysteMc.Model;
using ClinicSysteMc.ViewModel.Commands;
using Hardcodet.Wpf.TaskbarNotification;
using System.ComponentModel;
using System.Deployment.Application;
using System.Linq;

namespace ClinicSysteMc.ViewModel
{
    /// <summary>
    /// 20200510 created
    /// </summary>
    internal class MainVM : INotifyPropertyChanged
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private readonly TaskbarIcon tb = new TaskbarIcon();

        public MainVM() //constructor
        {
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

            log.Info($"Clinic System log in, version: {version}.");

            BTN_File = new FILEinput(this);
            BTN_YM = new YMinput(this);
            BTN_SE = new BEinput(this);
            BTN_ACT = new Plain(this);
            BTN_RFR = new DATArefresh(this);
            Refresh_Data();
        }

        #region Command Properties

        public FILEinput BTN_File { get; set; }

        public YMinput BTN_YM { get; set; }

        public BEinput BTN_SE { get; set; }

        public Plain BTN_ACT { get; set; }

        public DATArefresh BTN_RFR { get; set; }

        #endregion

        #region Data Properties

        private int _progressvalue;
        public int ProgressValue 
        {
            get { return _progressvalue; }
            set
            {
                _progressvalue = value;
                OnPropertyChanged("ProgressValue");
            } 
        }

        private object _loginout;

        public object LogInOut
        {
            get { return _loginout; }
            set
            {
                _loginout = value;
                OnPropertyChanged("LogInOut");
            }
        }

        private object _opd;

        public object OPD
        {
            get { return _opd; }
            set
            {
                _opd = value;
                OnPropertyChanged("OPD");
            }
        }

        private object _pt;

        public object PT
        {
            get { return _pt; }
            set
            {
                _pt = value;
                OnPropertyChanged("PT");
            }
        }

        private object _order;

        public object Order
        {
            get { return _order; }
            set
            {
                _order = value;
                OnPropertyChanged("Order");
            }
        }

        private object _upload;

        public object Upload
        {
            get { return _upload; }
            set
            {
                _upload = value;
                OnPropertyChanged("Upload");
            }
        }

        private object _pijia;

        public object Pijia
        {
            get { return _pijia; }
            set
            {
                _pijia = value;
                OnPropertyChanged("Pijia");
            }
        }

        private object _changeDepartment;

        public object ChangeDepartment
        {
            get { return _changeDepartment; }
            set
            {
                _changeDepartment = value;
                OnPropertyChanged("ChangeDepartment");
            }
        }

        private object _lab;

        public object Lab
        {
            get { return _lab; }
            set
            {
                _lab = value;
                OnPropertyChanged("Lab");
            }
        }

        #endregion Data Properties

        public void Refresh_Data()
        {
            CSDataContext dc = new CSDataContext();

            #region Function Page

            LogInOut = (from p in dc.log_Adm
                        where p.operation_name == "Log in" || p.operation_name == "Log out"
                        orderby p.regdate descending
                        select new { p.regdate, p.operation_name }).Take(100);
            // 20200522 add opd_import, for correct display
            OPD = (from p in dc.log_Adm
                   where p.operation_name == "add opd" || p.operation_name == "opd_import"
                   orderby p.regdate descending
                   select new { p.regdate }).Take(100);
            PT = (from p in dc.log_Adm
                  where p.operation_name == "病患檔案格式"
                  orderby p.regdate descending
                  select new { p.regdate }).Take(100);
            Order = (from p in dc.log_Adm
                     where p.operation_name == "計價檔格式"
                     orderby p.regdate descending
                     select new { p.regdate }).Take(100);
            Upload = (from p in dc.log_Adm
                      where p.operation_name == "健保上傳XML檔配對"
                      orderby p.regdate descending
                      select new { p.regdate }).Take(100);
            // 20200522 add PIJIA add/change, for correct display
            Pijia = (from p in dc.log_Adm
                     where p.operation_name == "新增批價檔: " || p.operation_name == "PIJIA add/change"
                     orderby p.regdate descending
                     select new { p.regdate }).Take(100);
            ChangeDepartment = (from p in dc.log_Adm
                                where p.operation_name == "change department"
                                orderby p.regdate descending
                                select new { p.regdate }).Take(100);
            Lab = (from p in dc.log_Adm
                   where p.operation_name == "Lab file format"
                   orderby p.regdate descending
                   select new { p.regdate }).Take(100);

            #endregion Function Page

            tb.ShowBalloonTip("完成", "主頁資料已更新", BalloonIcon.Info);

        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}