using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ObligacionesFinan.Clases
{
    class ObligacionesFinancieras
    {
        public int DocEntry { get; set; }
        public List<ObligacionesFinancierasLineas> ObligacionesFinancierasLineas { get; set; }
    }
}
