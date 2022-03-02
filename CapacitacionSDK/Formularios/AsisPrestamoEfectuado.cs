using CapacitacionSDK.SQL;
using ObligacionesFinan.Clases;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace ObligacionesFinan.Formularios
{
    class AsisPrestamoEfectuado : AsisPrestamo
    {
        public AsisPrestamoEfectuado(Application sboaApplication, SAPbobsCOM.Company sboCompany, string pais)
        {
            this.SBO_Application = sboaApplication;
            this.oCompany = sboCompany;
            this.DATASOURCE = "HCO_OOFE";
            this.Pais = pais;
        }

        public void CrearFormulario()
        {
            CargarFormulario();
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
                    oXmlDataDocument.Load(System.Windows.Forms.Application.StartupPath + @"/FormulariosXml/AsisPrestamoEfectuado.xml");
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

        protected override void actualizarUDO(int i, int rDREntry, int transId)
        {
            var dtCuotas = oForm.DataSources.DataTables.Item("dtCuotas");

            UDOObligacionesFinancieras uDOOFRecibida = new UDOObligacionesFinancieras("HCO_OOFE", "HCO_OFE1");
            uDOOFRecibida.DocEntry = (int)dtCuotas.GetValue("DocEntry", i);
            uDOOFRecibida.ObligacionesFinancierasLineas = new List<ObligacionesFinancierasLineas>();

            ObligacionesFinancierasLineas obligacionesFinancierasLineas = new ObligacionesFinancierasLineas();
            obligacionesFinancierasLineas.LineId = (int)dtCuotas.GetValue("LineId", i);
            obligacionesFinancierasLineas.RDREntry = rDREntry;
            obligacionesFinancierasLineas.TransId = transId;
            obligacionesFinancierasLineas.Interes = (double)dtCuotas.GetValue("U_HCO_Interes", i);
            obligacionesFinancierasLineas.Capita = (double)dtCuotas.GetValue("U_HCO_Capita", i);
            obligacionesFinancierasLineas.PayAmt = (double)dtCuotas.GetValue("U_HCO_PayAmt", i);

            uDOOFRecibida.ObligacionesFinancierasLineas.Add(obligacionesFinancierasLineas);

            uDOOFRecibida.actualizarUDO(ref oCompany);
        }

        protected override void consultaGrid()
        {
            SQL sql = new SQL("ObligacionesFinan.SQL.GetPayOFE.sql");
            oForm.DataSources.DataTables.Item("dtCuotas").ExecuteQuery(string.Format(sql.getQuery(), oForm.DataSources.UserDataSources.Item("udCCodeD").Value, oForm.DataSources.UserDataSources.Item("udCCodeH").Value));
        }
    }
}
