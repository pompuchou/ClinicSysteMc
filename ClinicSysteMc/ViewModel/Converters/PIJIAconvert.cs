using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClinicSysteMc.ViewModel.Converters
{
    internal class PIJIAconvert
    {
        private readonly DateTime _begindate;
        private readonly DateTime _enddate;
        public PIJIAconvert(DateTime begindate, DateTime enddate)
        {
            _begindate = begindate;
            _enddate = enddate;
        }

        public void Convert()
        {

        }
    }
}
