using ClinicSysteMc.Model;
using System.ComponentModel;
using System.Linq;


namespace ClinicSysteMc.ViewModel
{
    internal class LabMatchVM : INotifyPropertyChanged
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public LabMatchVM()
        {
            log.Info("Execute LabMatchVM.");
            Refresh_Data();
        }

        #region Data Properties

        private object _dataNOorder;

        public object DataNoOrder
        {
            get { return _dataNOorder; }
            set
            {
                OnPropertyChanged("DataNoOrder");
                _dataNOorder = value;
            }
        }

        private object _orderNOdata;

        public object OrderNoData
        {
            get { return _orderNOdata; }
            set
            {
                OnPropertyChanged("OrderNoData");
                _orderNOdata = value;
            }
        }

        #endregion Data Properties

        private void Refresh_Data()
        {
            CSDataContext dc = new CSDataContext();
            DataNoOrder = from p in dc.v_labdata_not_match_with_opd_order
                          orderby p.uid, p.l05
                          select p;
            OrderNoData = from p in dc.v_opdorder_not_match_with_lab_record
                          orderby p.uid, p.SDATE
                          select p;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}