using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace ClinicSysteMc.ViewModel.Dialog
{
    /// <summary>
    /// YMdialog.xaml 的互動邏輯
    /// </summary>
    public partial class YMdialog : Window, INotifyPropertyChanged
    {
        public YMdialog()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            // Don't accept the dialog box if there is invalid data
            if (!IsValid(this)) return;

            // Dialog box accepted
            DialogResult = true;
        }

        public string StrYM { get; set; }

        // Validate all dependency objects in a window
        private bool IsValid(DependencyObject node)
        {
            // Check if dependency object was passed
            if (node != null)
            {
                // Check if dependency object is valid.
                // NOTE: Validation.GetHasError works for controls that have validation rules attached 
                var isValid = !Validation.GetHasError(node);
                if (!isValid)
                {
                    // If the dependency object is invalid, and it can receive the focus,
                    // set the focus
                    if (node is IInputElement) Keyboard.Focus((IInputElement)node);
                    return false;
                }
            }

            // If this dependency object is valid, check all child dependency objects
            return LogicalTreeHelper.GetChildren(node).OfType<DependencyObject>().All(IsValid);

            // All dependency objects are valid
        }

        private void UPButton_Click(object sender, RoutedEventArgs e)
        {
            string strY = StrYM.Substring(0, 3);
            string strM = StrYM.Substring(3);
            if (strM == "12")
            {
                strM = "01";
                strY = (int.Parse(strY) + 1).ToString();
            }
            else
            {
                strM = (int.Parse(strM) + 101).ToString().Substring(1);
            }
            StrYM = strY + strM;

            RaisePropertyChanged("strYM");
        }

        private void DWButton_Click(object sender, RoutedEventArgs e)
        {
            string strY = StrYM.Substring(0, 3);
            string strM = StrYM.Substring(3);
            if (strM == "01")
            {
                strM = "12";
                strY = (int.Parse(strY) - 1).ToString();
            }
            else
            {
                strM = (int.Parse(strM) + 99).ToString().Substring(1);
            }
            StrYM = strY + strM;

            RaisePropertyChanged("strYM");
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

        #endregion // INotifyPropertyChanged Members
    }
}
