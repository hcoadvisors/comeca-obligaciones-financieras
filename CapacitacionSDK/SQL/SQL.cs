using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace CapacitacionSDK.SQL
{
    class SQL
    {
        private string query;
        private string recurso;

        public SQL(string recurso)
        {
            this.recurso = recurso;
            obtenerQuery();
        }

        private void obtenerQuery()
        {

            var assembly = Assembly.GetExecutingAssembly();

            using (Stream stream = assembly.GetManifestResourceStream(recurso))
            using (StreamReader reader = new StreamReader(stream))
            {
                query = reader.ReadToEnd();
            }
        }

        public string getQuery()
        {
            return this.query;
        }
    }
}
