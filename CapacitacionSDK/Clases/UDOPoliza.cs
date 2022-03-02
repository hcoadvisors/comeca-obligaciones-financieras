using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ObligacionesFinan.Clases
{
    class UDOPoliza : Poliza
    {
        protected string DATASOURCE;
        protected string DATASOURCELINE;

        public UDOPoliza(string dataSource, string dataSourceLine)
        {
            this.DATASOURCE = dataSource;
            this.DATASOURCELINE = dataSourceLine;
        }

        public void actualizarUDO(ref SAPbobsCOM.Company oComany)
        {
            SAPbobsCOM.GeneralService oGeneralService;
            SAPbobsCOM.GeneralData oGeneralData;
            SAPbobsCOM.GeneralDataParams oGeneralParams;
            SAPbobsCOM.GeneralDataCollection oChildren;

            oGeneralService = oComany.GetCompanyService().GetGeneralService(DATASOURCE);
            oGeneralParams = (SAPbobsCOM.GeneralDataParams)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);

            oGeneralParams.SetProperty("DocEntry", this.DocEntry);

            oGeneralData = oGeneralService.GetByParams(oGeneralParams);

            oChildren = oGeneralData.Child(DATASOURCELINE);

            for (int i = 0; i < oChildren.Count; i++)
            {
                for (int j = 0; j < this.PolizaLineas.Count; j++)
                {
                    if (oChildren.Item(i).GetProperty("LineId").ToString() == this.PolizaLineas[j].LineId.ToString())
                    {
                        oChildren.Item(i).SetProperty("U_HCO_TransId", this.PolizaLineas[j].TransId);
                    }
                }
            }

            oGeneralService.Update(oGeneralData);

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oGeneralService);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oGeneralData);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oGeneralParams);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oChildren);
            oGeneralService = null;
            oGeneralData = null;
            oGeneralParams = null;
            oChildren = null;
            GC.Collect();
        }
    }
}
