using CapacitacionSDK.SQL;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml;

namespace ObligacionesFinan.Formularios
{
    class PrestamoRecibido : Prestamo
    {
        public PrestamoRecibido(Application sboaApplication, SAPbobsCOM.Company sboCompany, string pais)
        {
            this.SBO_Application = sboaApplication;
            this.oCompany = sboCompany;
            this.DATASOURCE = "@HCO_OOFR";
            this.DATASOURCELINE = "@HCO_OFR1";
            this.Pais = pais;
        }

        public void CrearFormulario()
        {
            CargarFormulario();
            inicializar();
        }        

        public void CargarFormulario()
        {
            try
            {
                bool blnFormOpen = false;

                if (!blnFormOpen)
                {
                    FormCreationParams oFormCreationParams;
                    XmlDocument oXmlDataDocument = new XmlDocument();
                    oXmlDataDocument.Load(System.Windows.Forms.Application.StartupPath + @"/FormulariosXml/PrestamoRecibido.xml");
                    oFormCreationParams = (SAPbouiCOM.FormCreationParams)SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                    oFormCreationParams.XmlData = oXmlDataDocument.InnerXml;
                    oForm = SBO_Application.Forms.AddEx(oFormCreationParams);
                    oForm.Visible = true;

                }
            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }           

        }
        
        protected override void cargarDocNum()
        {
            SAPbobsCOM.Recordset oRecordset = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SQL sql = new SQL("ObligacionesFinan.SQL.GetMaxOOFRDocNum.sql");
            oRecordset.DoQuery(string.Format(sql.getQuery()));

            if (oRecordset.RecordCount > 0)
            {
                oForm.DataSources.DBDataSources.Item(DATASOURCE).SetValue("DocNum", 0, oRecordset.Fields.Item("DocNum").Value.ToString());
            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset);
            oRecordset = null;
            GC.Collect();
        }

        protected override void cargarCuentaTransitoria()
        {
            var cuenta = Constantes.consultarCampo("U_HCO_OFRAccTran", "OADM", "\'1\'", "1", ref this.oCompany);
            if (!string.IsNullOrEmpty(cuenta))
            {
                oForm.DataSources.DBDataSources.Item(DATASOURCE).SetValue("U_HCO_AccTran", 0, cuenta);
            }
        }
    }
}
