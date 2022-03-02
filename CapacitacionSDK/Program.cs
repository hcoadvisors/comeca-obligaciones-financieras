using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace CapacitacionSDK
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            ConexionAddOn oConexionAddon = new ConexionAddOn(); 

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run();
        }
    }
}
