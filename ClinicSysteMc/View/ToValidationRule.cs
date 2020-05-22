using System.Globalization;
using System.Windows.Controls;

namespace ClinicSysteMc.View
{
    internal class ToValidationRule : ValidationRule
    {
        public override ValidationResult Validate(object value, CultureInfo cultureInfo)
        {
            string sN = (string)value;
            // Is a number?
            if (!int.TryParse(sN, out int iN)) return new ValidationResult(false, "Not a number.");

            // Is in range?
            // the clinic began since MK105 September.
            // in MK166, I'm 104 y/o. I must be retired.
            if (iN < 0) return new ValidationResult(false, "must be greater than 0.");

            // Number is valid
            return new ValidationResult(true, null);
        }
    }
}
