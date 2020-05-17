using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClinicSysteMc.ViewModel.Converters
{
    internal class OPDconvert_auto
    {
        private readonly DateTime _begindate;
        private readonly DateTime _enddate;
        public OPDconvert_auto(DateTime begindate, DateTime enddate)
        {
            _begindate = begindate;
            _enddate = enddate;
        }

        public void Convert()
        {

        }
    }
}
