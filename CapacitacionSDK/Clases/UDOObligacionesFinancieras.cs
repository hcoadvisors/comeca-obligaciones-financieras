using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ObligacionesFinan.Clases
{
    class UDOObligacionesFinancieras : ObligacionesFinancieras
    {
        protected string DATASOURCE;
        protected string DATASOURCELINE;

        public UDOObligacionesFinancieras(string dataSource, string dataSourceLine)
        {
            this.DATASOURCE = dataSource;
            this.DATASOURCELINE = dataSourceLine;
        }
        public void actualizarUDO(ref SAPbobsCOM.Company oCompany)
        {
            SAPbobsCOM.GeneralService oGeneralService;
            SAPbobsCOM.GeneralData oGeneralData;
            SAPbobsCOM.GeneralDataParams oGeneralParams;
            SAPbobsCOM.GeneralDataCollection oChildren;

            oGeneralService = oCompany.GetCompanyService().GetGeneralService(DATASOURCE);
            oGeneralParams = (SAPbobsCOM.GeneralDataParams)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);

            oGeneralParams.SetProperty("DocEntry", this.DocEntry);

            oGeneralData = oGeneralService.GetByParams(oGeneralParams);

            oChildren = oGeneralData.Child(DATASOURCELINE);

            bool recalcularCuotas = false;
            double saldoInicial = 0;
            double capitalPagado = 0;
            double totalPagado = 0;
            int cuotasPagadas = 0;
            double importe = (double)oGeneralData.GetProperty("U_HCO_Amount");
            int plazo = oChildren.Count;
            string tipoCalculo = oGeneralData.GetProperty("U_HCO_CalcType").ToString();
            string periodicidad = oGeneralData.GetProperty("U_HCO_Periodi").ToString();
            double tasa = (double)oGeneralData.GetProperty("U_HCO_YearRate");
            bool calculoInicial = true;
            double cuota = 0;       
            double capital = 0;
            double valorTotal = 0;

            for (int i = 0; i < oChildren.Count; i++)
            {
                if ((int)oChildren.Item(i).GetProperty("U_HCO_TransId") != 0)
                {
                    capitalPagado += (double)oChildren.Item(i).GetProperty("U_HCO_Capita");
                    totalPagado += (double)oChildren.Item(i).GetProperty("U_HCO_PayAmt");
                    cuotasPagadas++;
                }
                if (oChildren.Item(i).GetProperty("LineId").ToString() == this.ObligacionesFinancierasLineas[0].LineId.ToString())
                {
                    oChildren.Item(i).SetProperty("U_HCO_RDREntry", this.ObligacionesFinancierasLineas[0].RDREntry);
                    oChildren.Item(i).SetProperty("U_HCO_TransId", this.ObligacionesFinancierasLineas[0].TransId);
                    oChildren.Item(i).SetProperty("U_HCO_Interes", this.ObligacionesFinancierasLineas[0].Interes);
                    capitalPagado += this.ObligacionesFinancierasLineas[0].Capita;
                    totalPagado += this.ObligacionesFinancierasLineas[0].PayAmt;
                    if ((double)oChildren.Item(i).GetProperty("U_HCO_Capita") != this.ObligacionesFinancierasLineas[0].Capita)
                    {
                        oChildren.Item(i).SetProperty("U_HCO_Capita", this.ObligacionesFinancierasLineas[0].Capita);
                        oChildren.Item(i).SetProperty("U_HCO_PayAmt", this.ObligacionesFinancierasLineas[0].PayAmt);
                        oChildren.Item(i).SetProperty("U_HCO_FinalAmt", (double)oChildren.Item(i).GetProperty("U_HCO_InitalAmt") - this.ObligacionesFinancierasLineas[0].Capita);
                        saldoInicial = (double)oChildren.Item(i).GetProperty("U_HCO_InitalAmt") - this.ObligacionesFinancierasLineas[0].Capita;
                        recalcularCuotas = true;
                        cuotasPagadas++;
                        continue;
                    }
                }
                
                if(recalcularCuotas)
                {
                    if (calculoInicial)
                    {
                        importe = importe - capitalPagado;
                        plazo = plazo - cuotasPagadas;
                                              
                        string baseTasa = oGeneralData.GetProperty("U_HCO_RateBase").ToString();

                        if (periodicidad.Equals("D"))   
                            tasa = tasa / 100 / (baseTasa == "1" ? 360 : 365);                        
                        else if (periodicidad.Equals("S"))
                            tasa = tasa / 100 / (baseTasa == "1" ? 360 : 365) * 7;
                        if (periodicidad.Equals("T"))                        
                            tasa = tasa / 100 / 12 * 3;
                        else if (periodicidad.Equals("SE"))                        
                            tasa = tasa / 100 / 12 * 6;
                        else if (periodicidad.Equals("A"))                                                   
                            tasa = tasa / 100;                        
                        else if (periodicidad.Equals("M"))                        
                            tasa = tasa / 100 / 12;                        

                        if (tipoCalculo.Equals("2"))
                        {
                            cuota = 1 + tasa;
                            cuota = Math.Pow(cuota, plazo);
                            cuota = tasa * cuota;
                            cuota = cuota * importe;
                            cuota = cuota / (Math.Pow((1 + tasa), plazo) - 1);

                            oGeneralData.SetProperty("U_HCO_MnthPay", cuota);
                            oGeneralData.SetProperty("U_HCO_DocTotal", (cuota * plazo) + totalPagado);
                            oGeneralData.SetProperty("U_HCO_OpenBal", (cuota * plazo));
                            oGeneralData.SetProperty("U_HCO_PaidToDate", totalPagado);
                        }
                        else                        
                            capital = importe / plazo;

                        calculoInicial = false;
                    }

                    if (tipoCalculo.Equals("2"))
                    {
                        oChildren.Item(i).SetProperty("U_HCO_InitalAmt", saldoInicial);
                        oChildren.Item(i).SetProperty("U_HCO_PayAmt", cuota);
                        oChildren.Item(i).SetProperty("U_HCO_Capita", cuota - (saldoInicial * tasa));
                        oChildren.Item(i).SetProperty("U_HCO_Interes", saldoInicial * tasa);
                        oChildren.Item(i).SetProperty("U_HCO_FinalAmt", saldoInicial - (cuota - (saldoInicial * tasa)));
                        saldoInicial = (saldoInicial - (cuota - (saldoInicial * tasa)));
                    }
                    else
                    {
                        oChildren.Item(i).SetProperty("U_HCO_InitalAmt", saldoInicial);
                        oChildren.Item(i).SetProperty("U_HCO_PayAmt", capital + (saldoInicial * tasa));
                        oChildren.Item(i).SetProperty("U_HCO_Capita", capital);
                        oChildren.Item(i).SetProperty("U_HCO_Interes", saldoInicial * tasa);
                        oChildren.Item(i).SetProperty("U_HCO_FinalAmt", saldoInicial - capital);
                        saldoInicial = (saldoInicial - capital);
                        valorTotal += (capital + (saldoInicial * tasa));
                    }

                    
                }
            }

            if(recalcularCuotas)
            {
                oChildren.Item(oChildren.Count - 1).SetProperty("U_HCO_FinalAmt", 0);
                if (tipoCalculo.Equals("1"))
                {
                    oGeneralData.SetProperty("U_HCO_MnthPay", capital);
                    oGeneralData.SetProperty("U_HCO_DocTotal", valorTotal + totalPagado);
                    oGeneralData.SetProperty("U_HCO_OpenBal", valorTotal);
                    oGeneralData.SetProperty("U_HCO_PaidToDate", totalPagado);
                }
            }
                
            else
            {
                double openBal = (double)oGeneralData.GetProperty("U_HCO_OpenBal");
                double paiToDate = (double)oGeneralData.GetProperty("U_HCO_PaidToDate");
                oGeneralData.SetProperty("U_HCO_OpenBal", openBal - this.ObligacionesFinancierasLineas[0].PayAmt);
                oGeneralData.SetProperty("U_HCO_PaidToDate", paiToDate + this.ObligacionesFinancierasLineas[0].PayAmt);
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
