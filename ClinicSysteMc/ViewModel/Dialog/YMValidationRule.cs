using System.Globalization;
using System.Windows.Controls;

namespace ClinicSysteMc.ViewModel.Dialog
{
    internal class YMValidationRule : ValidationRule
    {
        public override ValidationResult Validate(object value, CultureInfo cultureInfo)
        {
            string sYM = (string)value;
            // Is a number?
            if (!int.TryParse(sYM, out int iYM)) return new ValidationResult(false, "Not a number.");

            // Is in range?
            // the clinic began since MK105 September.
            // in MK166, I'm 104 y/o. I must be retired.
            if (iYM < 10509 || iYM > 16612) return new ValidationResult(false, "YM must be between 10509 and 16612.");

            // iYM, last 2 digits must between 1 and 12
            int iM = int.Parse(sYM.Substring(sYM.Length - 2));
            if (iM < 1 || iM > 12) return new ValidationResult(false, "Month must be between 1 and 12.");

            // Number is valid
            return new ValidationResult(true, null);
        }
    }
}