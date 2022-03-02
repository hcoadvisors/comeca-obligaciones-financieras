using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ObligacionesFinan.Clases
{
    class ObligacionesFinancierasLineas
    {
        public int LineId { get; set; }
        public int RDREntry { get; set; }
        public int TransId { get; set; }
        public double PayAmt { get; set; }
        public double Capita { get; set; }
        public double Interes { get; set; }
    }
}
