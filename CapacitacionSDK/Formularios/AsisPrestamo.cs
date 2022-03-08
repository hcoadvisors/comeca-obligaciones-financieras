using CapacitacionSDK.SQL;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ObligacionesFinan.Formularios
{
    abstract class AsisPrestamo
    {
        protected Application SBO_Application;
        protected Form oForm;
        protected SAPbobsCOM.Company oCompany;
        protected string DATASOURCE;
        protected string Pais;

        private double valorPago;

        public void ManejarEventosItem(ref ItemEvent pVal, ref bool BubbelEvent)
        {
            if (pVal.BeforeAction)
            {
                this.oForm = SBO_Application.Forms.Item(pVal.FormUID);
                if(pVal.EventType == BoEventTypes.et_VALIDATE)
                {
                    if (pVal.ItemUID.Equals("grCuotas"))
                    {
                        if (pVal.ColUID.Equals("U_HCO_PayAmt"))
                        {
                            if (!pVal.InnerEvent && pVal.ItemChanged)
                            {
                                var dtCuotas = oForm.DataSources.DataTables.Item("dtCuotas");
                                var pago = (double)dtCuotas.GetValue("U_HCO_PayAmt", pVal.Row);
                                var interes = (double)dtCuotas.GetValue("U_HCO_Interes", pVal.Row);
                                var comision = (double)dtCuotas.GetValue("Comision", pVal.Row);
                                var seguros = (double)dtCuotas.GetValue("U_HCO_Insuran", pVal.Row);
                                var otros = (double)dtCuotas.GetValue("U_HCO_Other", pVal.Row);
                                if (pago < (interes + comision + seguros + otros))
                                {
                                    SBO_Application.StatusBar.SetText("El valor del pago debe ser al menos el interes, la comisión, seguros y otros.");
                                    dtCuotas.SetValue("U_HCO_PayAmt", pVal.Row, this.valorPago);
                                }
                                else
                                {
                                    var capital = pago - interes - comision - seguros - otros;
                                    dtCuotas.SetValue("U_HCO_Capita", pVal.Row, capital);
                                }
                            }
                        }
                        else if (pVal.ColUID.Equals("U_HCO_Interes"))
                        {
                            if (!pVal.InnerEvent && pVal.ItemChanged)
                            {
                                var dtCuotas = oForm.DataSources.DataTables.Item("dtCuotas");
                                var capital = (double)dtCuotas.GetValue("U_HCO_Capita", pVal.Row);
                                var interes = (double)dtCuotas.GetValue("U_HCO_Interes", pVal.Row);
                                var comision = (double)dtCuotas.GetValue("Comision", pVal.Row);
                                var seguros = (double)dtCuotas.GetValue("U_HCO_Insuran", pVal.Row);
                                var otros = (double)dtCuotas.GetValue("U_HCO_Other", pVal.Row);
                                var pago = capital + interes + comision + seguros + otros;
                                dtCuotas.SetValue("U_HCO_PayAmt", pVal.Row, pago);
                            }
                        }
                        else if (pVal.ColUID.Equals("FechaContabilizacion"))
                        {
                            if (!pVal.InnerEvent && pVal.ItemChanged)
                            {
                                var dtCuotas = oForm.DataSources.DataTables.Item("dtCuotas");
                                if ((double)dtCuotas.GetValue("PorcIntMora", pVal.Row) > 0)
                                {
                                    var valorInteres = (((DateTime)dtCuotas.GetValue("FechaContabilizacion", pVal.Row) - (DateTime)dtCuotas.GetValue("U_HCO_Date", pVal.Row)).TotalDays)
                                        * ((double)dtCuotas.GetValue("U_HCO_Capita", pVal.Row) * (double)dtCuotas.GetValue("PorcIntMora", pVal.Row) / 100);
                                    if (valorInteres > 0)
                                        dtCuotas.SetValue("InteresMora", pVal.Row, valorInteres);
                                    else
                                        dtCuotas.SetValue("InteresMora", pVal.Row, 0);
                                }
                            }
                        }
                        else if(pVal.ColUID.Equals("InteresMora"))
                        {
                            if (!pVal.InnerEvent && pVal.ItemChanged)
                            {
                                var dtCuotas = oForm.DataSources.DataTables.Item("dtCuotas");
                                var porInteres = (double)dtCuotas.GetValue("InteresMora", pVal.Row) / (((DateTime)dtCuotas.GetValue("FechaContabilizacion", pVal.Row) - (DateTime)dtCuotas.GetValue("U_HCO_Date", pVal.Row)).TotalDays)
                                    / (double)dtCuotas.GetValue("U_HCO_Capita", pVal.Row) * 100;
                                dtCuotas.SetValue("PorcIntMora", pVal.Row, porInteres);
                            }
                        }
                        else if (pVal.ColUID.Equals("PorcIntMora"))
                        {
                            if (!pVal.InnerEvent && pVal.ItemChanged)
                            {
                                var dtCuotas = oForm.DataSources.DataTables.Item("dtCuotas");
                                var valorInteres = (((DateTime)dtCuotas.GetValue("FechaContabilizacion", pVal.Row) - (DateTime)dtCuotas.GetValue("U_HCO_Date", pVal.Row)).TotalDays)
                                       * ((double)dtCuotas.GetValue("U_HCO_Capita", pVal.Row) * (double)dtCuotas.GetValue("PorcIntMora", pVal.Row) / 100);
                                if (valorInteres > 0)
                                    dtCuotas.SetValue("InteresMora", pVal.Row, valorInteres);
                                else
                                    dtCuotas.SetValue("InteresMora", pVal.Row, 0);
                            }
                        }
                    }                    
                }
            }
            else
            {
                if (pVal.EventType == BoEventTypes.et_CHOOSE_FROM_LIST)
                {
                    if (pVal.ItemUID.Equals("txtCCodeD"))
                    {
                        IChooseFromListEvent chflarg = (IChooseFromListEvent)pVal;
                        if (chflarg.SelectedObjects != null)
                        {
                            DataTable dt = chflarg.SelectedObjects;
                            this.oForm.DataSources.UserDataSources.Item("udCCodeD").Value = Convert.ToString(dt.GetValue("CardCode", 0));
                        }
                    }
                    else if (pVal.ItemUID.Equals("txtCCodeH"))
                    {
                        IChooseFromListEvent chflarg = (IChooseFromListEvent)pVal;
                        if (chflarg.SelectedObjects != null)
                        {
                            DataTable dt = chflarg.SelectedObjects;
                            this.oForm.DataSources.UserDataSources.Item("udCCodeH").Value = Convert.ToString(dt.GetValue("CardCode", 0));
                        }
                    }
                }
                else if (pVal.EventType == BoEventTypes.et_ITEM_PRESSED)
                {
                    if (pVal.ItemUID.Equals("btnSig"))
                    {
                        switch (oForm.PaneLevel)
                        {
                            case 1:
                                if (validarCampos())
                                {
                                    cargarGrid();
                                    ((Button)oForm.Items.Item("btnSig").Specific).Caption = "Procesar";
                                    oForm.Items.Item("btnAnt").Visible = true;
                                    oForm.PaneLevel = 2;
                                }
                                break;
                            case 2:
                                if (SBO_Application.MessageBox("Este proceso es irrebersible, desa continuar", 1, "Continuar", "Cancelar") == 1)
                                {
                                    contabilizarCuotas();
                                    oForm.Items.Item("2").Visible = false;
                                    oForm.Items.Item("btnAnt").Visible = false;
                                    ((Button)oForm.Items.Item("btnSig").Specific).Caption = "Finalizar";
                                    oForm.PaneLevel = 3;
                                }
                                break;
                            case 3:
                                oForm.Close();
                                break;
                        }
                    }
                    else if (pVal.ItemUID.Equals("btnAnt"))
                    {
                        if (oForm.PaneLevel == 2)
                        {
                            oForm.Items.Item("btnAnt").Visible = false;
                            ((Button)oForm.Items.Item("btnSig").Specific).Caption = "Siguiente";
                            oForm.PaneLevel = 1;
                        }
                    }
                }
                else if (pVal.EventType == BoEventTypes.et_GOT_FOCUS)
                {
                    if (pVal.ItemUID.Equals("grCuotas"))
                    {
                        if (pVal.ColUID.Equals("U_HCO_PayAmt"))
                        {
                            var dtCuotas = oForm.DataSources.DataTables.Item("dtCuotas");
                            this.valorPago = (double)dtCuotas.GetValue("U_HCO_PayAmt", pVal.Row);
                        }
                    }
                }
            }
        }

        private bool validarCampos()
        {
            bool resultado = true;
            if (string.IsNullOrEmpty(oForm.DataSources.UserDataSources.Item("udCCodeD").Value))
            {
                SBO_Application.StatusBar.SetText("Debe seleccionar el socio de negocios desde", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                oForm.Items.Item("txtCCodeD").Click(BoCellClickType.ct_Regular);
                resultado = false;
            }
            else if (string.IsNullOrEmpty(oForm.DataSources.UserDataSources.Item("udCCodeH").Value))
            {
                SBO_Application.StatusBar.SetText("Debe seleccionar el socio de negocios hasta", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                oForm.Items.Item("txtCCodeH").Click(BoCellClickType.ct_Regular);
                resultado = false;
            }

            return resultado;
        }

        private void cargarGrid()
        {
            consultaGrid();            

            var grCuotas = (Grid)oForm.Items.Item("grCuotas").Specific;

            grCuotas.Columns.Item("DocNum").Editable = false;
            grCuotas.Columns.Item("DocEntry").Editable = false;
            grCuotas.Columns.Item("U_HCO_Date").Editable = false;
            grCuotas.Columns.Item("U_HCO_CardCode").Editable = false;
            grCuotas.Columns.Item("U_HCO_InitalAmt").Editable = false;
            grCuotas.Columns.Item("U_HCO_Capita").Editable = false;
            grCuotas.Columns.Item("LineId").Editable = false;
            grCuotas.Columns.Item("LineId").Visible = false;
            grCuotas.Columns.Item("U_HCO_AcctCode").Visible = false;
            grCuotas.Columns.Item("Comision").Editable = false;
            grCuotas.Columns.Item("U_HCO_OcrCode").Visible = false;
            grCuotas.Columns.Item("U_HCO_OcrCode").Editable = false;
            grCuotas.Columns.Item("U_HCO_Currency").Visible = false;
            grCuotas.Columns.Item("U_HCO_Currency").Editable = false;
            grCuotas.Columns.Item("MainCurncy").Visible = false;
            grCuotas.Columns.Item("MainCurncy").Editable = false;
            grCuotas.Columns.Item("U_HCO_AccBank").Visible = false;
            grCuotas.Columns.Item("U_HCO_AccBank").Editable = false;
            grCuotas.Columns.Item("#").Visible = false;
            grCuotas.Columns.Item("#").Editable = false;
            grCuotas.Columns.Item("U_HCO_MnthPay").Visible = false;
            grCuotas.Columns.Item("U_HCO_MnthPay").Editable = false;
            grCuotas.Columns.Item("U_HCO_AccComm").Visible = false;
            grCuotas.Columns.Item("U_HCO_AccComm").Editable = false;
            grCuotas.Columns.Item("U_HCO_Insuran").Editable = false;
            grCuotas.Columns.Item("U_HCO_AccIns").Editable = false;
            grCuotas.Columns.Item("U_HCO_AccIns").Visible = false;
            grCuotas.Columns.Item("U_HCO_Other").Editable = false;
            grCuotas.Columns.Item("U_HCO_AccOthe").Visible = false;
            grCuotas.Columns.Item("U_HCO_AccOthe").Editable = false;

            ((EditTextColumn)grCuotas.Columns.Item("DocEntry")).LinkedObjectType = DATASOURCE;
            ((EditTextColumn)grCuotas.Columns.Item("U_HCO_CardCode")).LinkedObjectType = "2";

            ((EditTextColumn)grCuotas.Columns.Item("Check")).TitleObject.Caption = "Seleccionar";
            ((EditTextColumn)grCuotas.Columns.Item("DocNum")).TitleObject.Caption = "Numero obliación";
            ((EditTextColumn)grCuotas.Columns.Item("DocEntry")).TitleObject.Caption = "Numero int. obligación";
            ((EditTextColumn)grCuotas.Columns.Item("U_HCO_CardCode")).TitleObject.Caption = "Código SN";
            ((EditTextColumn)grCuotas.Columns.Item("U_HCO_Date")).TitleObject.Caption = "Fecha de cuota";
            ((EditTextColumn)grCuotas.Columns.Item("FechaContabilizacion")).TitleObject.Caption = "Fecha de contabilizacion";
            ((EditTextColumn)grCuotas.Columns.Item("U_HCO_InitalAmt")).TitleObject.Caption = "Monto inicial";
            ((EditTextColumn)grCuotas.Columns.Item("U_HCO_PayAmt")).TitleObject.Caption = "Valor cuota";
            ((EditTextColumn)grCuotas.Columns.Item("U_HCO_Capita")).TitleObject.Caption = "Valor abono a capital";
            ((EditTextColumn)grCuotas.Columns.Item("U_HCO_Interes")).TitleObject.Caption = "Valor intereses";
            ((EditTextColumn)grCuotas.Columns.Item("InteresMora")).TitleObject.Caption = "Valor intereses mora";
            ((EditTextColumn)grCuotas.Columns.Item("PorcIntMora")).TitleObject.Caption = "Porcentaje intereses mora";
            ((EditTextColumn)grCuotas.Columns.Item("U_HCO_Other")).TitleObject.Caption = "Otros";
            ((EditTextColumn)grCuotas.Columns.Item("U_HCO_Insuran")).TitleObject.Caption = "Seguros";

            grCuotas.Columns.Item("Check").Type = BoGridColumnType.gct_CheckBox;
        }

        private void contabilizarCuotas()
        {
            var dtCuotas = oForm.DataSources.DataTables.Item("dtCuotas");

            var dtResult = oForm.DataSources.DataTables.Item("dtResult");
            dtResult.Clear();            
            dtResult.Columns.Add("DocNum", BoFieldsType.ft_AlphaNumeric);
            dtResult.Columns.Add("RDREntry", BoFieldsType.ft_AlphaNumeric);
            dtResult.Columns.Add("TransId", BoFieldsType.ft_AlphaNumeric);
            dtResult.Columns.Add("Descripcion", BoFieldsType.ft_AlphaNumeric);

            int j = 0;
            for (int i = 0; i < dtCuotas.Rows.Count; i++)
            {
                if (dtCuotas.GetValue("Check", i).ToString().Equals("Y"))
                {
                    dtResult.Rows.Add();
                    dtResult.SetValue("DocNum", j, dtCuotas.GetValue("DocNum", i));
                    try
                    {
                        SBO_Application.StatusBar.SetText("Procesando cuota " + (i + 1) + " de " + dtCuotas.Rows.Count, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                        int rDREntry = 0;
                        int transId = 0;

                        if (!oCompany.InTransaction)
                            oCompany.StartTransaction();

                        var resultado = abonarCapital(i);
                        if (resultado.Item1)
                        {
                            dtResult.SetValue("TransId", j, resultado.Item2);
                            int.TryParse(resultado.Item2, out transId);                            
                        }
                        else
                        {
                            dtResult.SetValue("Descripcion", j, "Error pagando capital - " + resultado.Item2);
                            if (oCompany.InTransaction)
                                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                            continue;
                        }
                        if ((double)dtCuotas.GetValue("U_HCO_Interes", i) > 0)
                        {
                            var resultadoOrden = crearOrden(i);
                            if (resultadoOrden.Item1)
                            {
                                dtResult.SetValue("RDREntry", j, resultadoOrden.Item2);
                                rDREntry = int.Parse(resultadoOrden.Item2);
                            }
                            else
                            {
                                dtResult.SetValue("TransId", j, "");
                                dtResult.SetValue("Descripcion", j, "Error creando pedido de intereses - " + resultadoOrden.Item2);
                                if (oCompany.InTransaction)
                                    oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                                continue;
                            }
                        }
                        
                        if(rDREntry != 0 || transId != 0)
                            actualizarUDO(i, rDREntry, transId);

                        dtResult.SetValue("Descripcion", j, "Cuota contabilizada correctamente");

                        if (oCompany.InTransaction)
                            oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                    }
                    catch (Exception ex)
                    {
                        if (oCompany.InTransaction)
                            oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                        dtResult.SetValue("TransId", j, "");
                        dtResult.SetValue("RDREntry", j, "");
                        dtResult.SetValue("Descripcion", j, ex.Message);
                    }
                    j++;
                }
            }

            SBO_Application.StatusBar.SetText("Proceso finalizado", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);

            var grResult = (Grid)oForm.Items.Item("grResult").Specific;

            grResult.Columns.Item("RDREntry").Editable = false;
            grResult.Columns.Item("TransId").Editable = false;
            grResult.Columns.Item("Descripcion").Editable = false;
            grResult.Columns.Item("DocNum").Editable = false;

            ((EditTextColumn)grResult.Columns.Item("RDREntry")).LinkedObjectType = DATASOURCE.Equals("HCO_OOFR") ? "22" : "17";
            ((EditTextColumn)grResult.Columns.Item("TransId")).LinkedObjectType = DATASOURCE.Equals("HCO_OOFR") ? "46" : "24";

            ((EditTextColumn)grResult.Columns.Item("DocNum")).TitleObject.Caption = "Prestamo";
            ((EditTextColumn)grResult.Columns.Item("RDREntry")).TitleObject.Caption = "Pedido de intereses";
            ((EditTextColumn)grResult.Columns.Item("TransId")).TitleObject.Caption = "Pago de capital";
            ((EditTextColumn)grResult.Columns.Item("Descripcion")).TitleObject.Caption = "Resultado";

            grResult.AutoResizeColumns();
        }

        protected (bool, string) abonarCapital(int i)
        {
            var dtCuotas = oForm.DataSources.DataTables.Item("dtCuotas");
            bool crearPago = false;

            double total = 0;

            SAPbobsCOM.Payments oPayment = (SAPbobsCOM.Payments)oCompany.GetBusinessObject(DATASOURCE.Equals("HCO_OOFR") ? SAPbobsCOM.BoObjectTypes.oVendorPayments : SAPbobsCOM.BoObjectTypes.oIncomingPayments);

            oPayment.DocDate = (DateTime)dtCuotas.GetValue("FechaContabilizacion", i);
            oPayment.CounterReference = dtCuotas.GetValue("DocNum", i).ToString();
            oPayment.Remarks = "Cuota de obl. rec. Linea " + dtCuotas.GetValue("LineId", i).ToString();
            oPayment.DocType = SAPbobsCOM.BoRcptTypes.rAccount;
            oPayment.DocCurrency = dtCuotas.GetValue("U_HCO_Currency", i).ToString();

            if((double)dtCuotas.GetValue("U_HCO_Capita", i) > 0)
            {
                oPayment.AccountPayments.AccountCode = dtCuotas.GetValue("U_HCO_AcctCode", i).ToString();
                oPayment.AccountPayments.SumPaid = (double)dtCuotas.GetValue("U_HCO_Capita", i);
                total += (double)dtCuotas.GetValue("U_HCO_Capita", i);
                oPayment.AccountPayments.ProfitCenter = dtCuotas.GetValue("U_HCO_OcrCode", i).ToString();
                if (Pais.Equals(Constantes.Colombia))
                {
                    oPayment.AccountPayments.UserFields.Fields.Item("U_HBT_Tercero").Value = dtCuotas.GetValue("U_HCO_CardCode", i).ToString();
                }
                oPayment.AccountPayments.Add();
                crearPago = true;
            }                       
            if((int)dtCuotas.GetValue("LineId", i) == 1)
            {
                if((double)dtCuotas.GetValue("Comision", i) > 0)
                {
                    oPayment.AccountPayments.AccountCode = dtCuotas.GetValue("U_HCO_AccComm", i).ToString();
                    oPayment.AccountPayments.SumPaid = (double)dtCuotas.GetValue("Comision", i);
                    total += (double)dtCuotas.GetValue("Comision", i);
                    oPayment.AccountPayments.ProfitCenter = dtCuotas.GetValue("U_HCO_OcrCode", i).ToString();
                    if (Pais.Equals(Constantes.Colombia))
                    {
                        oPayment.AccountPayments.UserFields.Fields.Item("U_HBT_Tercero").Value = dtCuotas.GetValue("U_HCO_CardCode", i).ToString();
                    }
                    oPayment.AccountPayments.Add();
                    crearPago = true;
                }
                if ((double)dtCuotas.GetValue("U_HCO_Insuran", i) > 0)
                {
                    oPayment.AccountPayments.AccountCode = dtCuotas.GetValue("U_HCO_AccIns", i).ToString();
                    oPayment.AccountPayments.SumPaid = (double)dtCuotas.GetValue("U_HCO_Insuran", i);
                    total += (double)dtCuotas.GetValue("U_HCO_Insuran", i);
                    oPayment.AccountPayments.ProfitCenter = dtCuotas.GetValue("U_HCO_OcrCode", i).ToString();
                    if (Pais.Equals(Constantes.Colombia))
                    {
                        oPayment.AccountPayments.UserFields.Fields.Item("U_HBT_Tercero").Value = dtCuotas.GetValue("U_HCO_CardCode", i).ToString();
                    }
                    oPayment.AccountPayments.Add();
                    crearPago = true;
                }
                if ((double)dtCuotas.GetValue("U_HCO_Other", i) > 0)
                {
                    oPayment.AccountPayments.AccountCode = dtCuotas.GetValue("U_HCO_AccOthe", i).ToString();
                    oPayment.AccountPayments.SumPaid = (double)dtCuotas.GetValue("U_HCO_Other", i);
                    total += (double)dtCuotas.GetValue("U_HCO_Other", i);
                    oPayment.AccountPayments.ProfitCenter = dtCuotas.GetValue("U_HCO_OcrCode", i).ToString();
                    if (Pais.Equals(Constantes.Colombia))
                    {
                        oPayment.AccountPayments.UserFields.Fields.Item("U_HBT_Tercero").Value = dtCuotas.GetValue("U_HCO_CardCode", i).ToString();
                    }
                    oPayment.AccountPayments.Add();
                    crearPago = true;
                }
            }

            if (crearPago)
            {
                oPayment.TransferAccount = dtCuotas.GetValue("U_HCO_AccBank", i).ToString();
                oPayment.TransferSum = total;
                oPayment.TransferDate = (DateTime)dtCuotas.GetValue("FechaContabilizacion", i);
                oPayment.TransferReference = dtCuotas.GetValue("Referencia", i).ToString();

                if (oPayment.Add() != 0)
                {
                    return (false, oCompany.GetLastErrorDescription());
                }
                else
                {
                    return (true, oCompany.GetNewObjectKey());
                }
            }
            else
                return (true, "No hay abonos a capital, comisión, seguros u otros que pagar");
        }

        protected (bool, string) crearOrden(int i)
        {
            var dtCuotas = oForm.DataSources.DataTables.Item("dtCuotas");

            SAPbobsCOM.Documents oOrder = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(DATASOURCE.Equals("HCO_OOFR") ? SAPbobsCOM.BoObjectTypes.oPurchaseOrders : SAPbobsCOM.BoObjectTypes.oOrders);

            oOrder.CardCode = dtCuotas.GetValue("U_HCO_CardCode", i).ToString();
            oOrder.DocDate = (DateTime)dtCuotas.GetValue("FechaContabilizacion", i);
            oOrder.DocDueDate = (DateTime)dtCuotas.GetValue("FechaContabilizacion", i);
            oOrder.DocCurrency = dtCuotas.GetValue("U_HCO_Currency", i).ToString();

            oOrder.Lines.ItemCode = "INTERESES";
            oOrder.Lines.UnitPrice = (double)dtCuotas.GetValue("U_HCO_Interes", i) ;
            oOrder.Lines.CostingCode = dtCuotas.GetValue("U_HCO_OcrCode", i).ToString();
            oOrder.Lines.Add();

            if((double)dtCuotas.GetValue("InteresMora", i) > 0)
            {
                oOrder.Lines.ItemCode = "INTERESESMORA";
                oOrder.Lines.UnitPrice = (double)dtCuotas.GetValue("InteresMora", i);
                oOrder.Lines.CostingCode = dtCuotas.GetValue("U_HCO_OcrCode", i).ToString();
                oOrder.Lines.Add();
            }

            if (oOrder.Add() != 0)
            {
                return (false, oCompany.GetLastErrorDescription());
            }
            else
            {
                return (true, oCompany.GetNewObjectKey());
            }
        }
        protected abstract void consultaGrid();
        protected abstract void actualizarUDO(int i, int rDREntry, int transId);

    }
}
