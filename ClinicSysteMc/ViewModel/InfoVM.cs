using ClinicSysteMc.Model;
using ClinicSysteMc.ViewModel.Commands;
using Hardcodet.Wpf.TaskbarNotification;
using System.ComponentModel;
using System.Linq;

namespace ClinicSysteMc.ViewModel
{
    /// <summary>
    /// 20200510 created
    /// </summary>
    internal class InfoVM : INotifyPropertyChanged
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private readonly TaskbarIcon tb = new TaskbarIcon();

        public InfoVM() //constructor
        {
            log.Info("Execute InfoVM.");
            BTN_RFR = new DATArefreshi(this);
            Refresh_Data();
        }

        #region Command Properties

        public DATArefreshi BTN_RFR { get; set; }

        #endregion

        #region Data Properties

        private object _adm;

        public object Admin
        {
            get { return _adm; }
            set
            {
                OnPropertyChanged("Admin");
                _adm = value;
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

        private object _err;

        public object Err
        {
            get { return _err; }
            set
            {
                OnPropertyChanged("Err");
                _err = value;
            }
        }

        #endregion Data Properties

        public void Refresh_Data()
        {
            CSDataContext dc = new CSDataContext();
            // 20200522 add Companion Log in and out
            Admin = (from p in dc.log_Adm
                     where p.operation_name != "Log in" && p.operation_name != "Log out" && p.operation_name != "update opd" &&
                           p.operation_name != "update opd order" && p.operation_name != "OPD file format" &&
                           p.operation_name != "Change order data" && p.operation_name != "Add a new order" &&
                           p.operation_name != "Lab file format" && p.operation_name != "Change patient data" &&
                           p.operation_name != "Add a new patient" && p.operation_name != "Companion Log out" &&
                           p.operation_name != "Companion Log in"
                     orderby p.regdate descending
                     select new { p.regdate, p.operation_name, p.description }).Take(100);
            Err = (from p in dc.log_Err
                   orderby p.error_date descending
                   select new { p.error_date, p.error_message }).Take(100);
            OPD = (from p in dc.log_Adm
                   where p.operation_name == "update opd" || p.operation_name == "update opd order"
                   orderby p.regdate descending
                   select new { p.regdate, p.description }).Take(100);
            Order = (from p in dc.log_Adm
                     where p.operation_name == "Change order data" || p.operation_name == "Add a new order"
                     orderby p.regdate descending
                     select new { p.regdate, p.description }).Take(100);
            PT = (from p in dc.log_Adm
                  where p.operation_name == "Change patient data" || p.operation_name == "Add a new patient"
                  orderby p.regdate descending
                  select new { p.regdate, p.description }).Take(100);

            tb.ShowBalloonTip("完成", "訊息頁資料已更新", BalloonIcon.Info);

        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}