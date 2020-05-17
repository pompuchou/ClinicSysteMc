using ClinicSysteMc.Model;
using ClinicSysteMc.ViewModel.Commands;
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
            Refresh_Data();
        }

        #region Command Properties

        public FILEinput BTN_File { get; set; }

        public YMinput BTN_YM { get; set; }

        public BEinput BTN_SE { get; set; }

        public Plain BTN_ACT { get; set; }

        #endregion

        #region Data Properties

        private object _loginout;

        public object LogInOut
        {
            get { return _loginout; }
            set
            {
                OnPropertyChanged("LogInOut");
                _loginout = value;
            }
        }

        private object _opd;

        public object OPD
        {
            get { return _opd; }
            set
            {
                OnPropertyChanged("OPD");
                _opd = value;
            }
        }

        private object _pt;

        public object PT
        {
            get { return _pt; }
            set
            {
                OnPropertyChanged("PT");
                _pt = value;
            }
        }

        private object _order;

        public object Order
        {
            get { return _order; }
            set
            {
                OnPropertyChanged("Order");
                _order = value;
            }
        }

        private object _upload;

        public object Upload
        {
            get { return _upload; }
            set
            {
                OnPropertyChanged("Upload");
                _upload = value;
            }
        }

        private object _pijia;

        public object Pijia
        {
            get { return _pijia; }
            set
            {
                OnPropertyChanged("Pijia");
                _pijia = value;
            }
        }

        private object _changeDepartment;

        public object ChangeDepartment
        {
            get { return _changeDepartment; }
            set
            {
                OnPropertyChanged("ChangeDepartment");
                _changeDepartment = value;
            }
        }

        private object _lab;

        public object Lab
        {
            get { return _lab; }
            set
            {
                OnPropertyChanged("Lab");
                _lab = value;
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
            OPD = (from p in dc.log_Adm
                   where p.operation_name == "add opd"
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
            Pijia = (from p in dc.log_Adm
                     where p.operation_name == "新增批價檔: "
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
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}