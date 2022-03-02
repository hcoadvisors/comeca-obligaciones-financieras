using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ObligacionesFinan.Clases
{
    class Poliza
    {
        public int DocEntry { get; set; }
        public List<PolizaLinea> PolizaLineas { get; set; }
    }
}
