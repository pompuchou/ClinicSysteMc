using System;
using System.ComponentModel;
using System.Windows;

namespace ClinicSysteMc.ViewModel.Dialog
{
    /// <summary>
    /// BEdialog.xaml 的互動邏輯
    /// </summary>
    public partial class BEdialog : Window, INotifyPropertyChanged
    {
        public BEdialog(DateTime begindate, DateTime enddate)
        {
            InitializeComponent();
            _begindate = begindate;
            _enddate = enddate;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            // Dialog box accepted
            DialogResult = true;
        }

        #region INotifyPropertyChanged Members

        /// <summary>
        /// Raises the PropertyChange event for the property specified
        /// </summary>
        /// <param name="propertyName">Property name to update. Is case-sensitive.</param>
        public virtual void RaisePropertyChanged(string propertyName)
        {
            OnPropertyChanged(propertyName);
        }

        private DateTime _begindate;

        public DateTime BeginDate
        {
            get { return _begindate; }
            set 
            { 
                _begindate = value;
                // begindate and enddate must be in the same month
                // begindate must be earlier than enddate
                if ((_begindate.Year != _enddate.Year || _begindate.Month != _enddate.Month) || (_begindate.CompareTo(_enddate) > 0))
                {
                    // set enddate to the end date of the same month
                    EndDate = DateTime.Parse($"{_begindate.Year}/{_begindate.Month}/1").AddMonths(1).AddSeconds(-1);
                }
                RaisePropertyChanged("BeginDate");
            }
        }

        private DateTime _enddate;

        public DateTime EndDate
        {
            get { return _enddate; }
            set 
            { 
                _enddate = value;
                // begindate and enddate must be in the same month
                // begindate must be earlier than enddate
                if ((_begindate.Year != _enddate.Year || _begindate.Month != _enddate.Month) || (_begindate.CompareTo(_enddate) > 0))
                {
                    // set begindate to the end date of the same month
                    BeginDate = DateTime.Parse($"{_enddate.Year}/{_enddate.Month}/1");
                }
                RaisePropertyChanged("EndDate");
            }
        }


        /// <summary>
        /// Raised when a property on this object has a new value.
        /// </summary>
        public event PropertyChangedEventHandler PropertyChanged;

        /// <summary>
        /// Raises this object's PropertyChanged event.
        /// </summary>
        /// <param name="propertyName">The property that has a new value.</param>
        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChangedEventHandler handler = this.PropertyChanged;
            if (handler != null)
            {
                var e = new PropertyChangedEventArgs(propertyName);
                handler(this, e);
            }
        }

        #endregion INotifyPropertyChanged Members
    }
}