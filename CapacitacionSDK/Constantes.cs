using CapacitacionSDK.SQL;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ObligacionesFinan
{
    static class Constantes
    {
        public const string
            Colombia = "CO",
            Guatemala = "GL";

        public static string consultarCampo(string campo, string tabla, string condicionCampoo, string condicionValor, ref SAPbobsCOM.Company oCompany)
        {
            string valor = "";
            SAPbobsCOM.Recordset oRecordset = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SQL sql = new SQL("ObligacionesFinan.SQL.GetField.sql");
            oRecordset.DoQuery(string.Format(sql.getQuery(), campo, tabla, condicionCampoo, condicionValor));
            if (oRecordset.RecordCount > 0)
            {
                valor = oRecordset.Fields.Item(0).Value.ToString();
            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset);
            oRecordset = null;
            GC.Collect();

            return valor;
        }
    }
}
