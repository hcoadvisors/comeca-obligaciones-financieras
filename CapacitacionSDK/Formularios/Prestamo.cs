using CapacitacionSDK.SQL;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace ObligacionesFinan.Formularios
{
    abstract class Prestamo
    {
        protected Application SBO_Application;
        protected Form oForm;
        protected SAPbobsCOM.Company oCompany;
        protected string DATASOURCE;
        protected string DATASOURCELINE;
        protected string Pais;

        private double capital;
        private int plazo;
        private double interes;

        public void ManejarEventosMenus(ref MenuEvent pVal, ref bool pBubbleEvent)
        {
            if (pVal.BeforeAction)
                this.oForm = SBO_Application.Forms.ActiveForm;
            else
            {

                if (new string[] { "1282", "1281" }.Contains(pVal.MenuUID))
                {
                    activarCampos(true);
                    if (pVal.MenuUID == "1281")
                    {
                        this.oForm.Items.Item("txtDocNum").Enabled = true;
                        this.oForm.Items.Item("txtTransId").Enabled = true;
                        this.oForm.Items.Item("btnCalCuo").Enabled = false;
                    }
                    else if (pVal.MenuUID == "1282")
                    {
                        cargarDocNum();
                        cargarCuentaTransitoria();
                    }
                }

            }
        }

        public void ManejarEventosItem(ref ItemEvent pVal, ref bool BubbelEvent)
        {
            if (pVal.BeforeAction)
            {
                this.oForm = SBO_Application.Forms.Item(pVal.FormUID);
                if (pVal.EventType == BoEventTypes.et_ITEM_PRESSED)
                {
                    if (pVal.ItemUID.Equals("1"))
                    {
                        if (oForm.Mode == BoFormMode.fm_ADD_MODE || oForm.Mode == BoFormMode.fm_UPDATE_MODE)
                        {
                            BubbelEvent = validarCampos();
                            if (BubbelEvent)
                            {
                                Matrix matrixCuotas = (Matrix)oForm.Items.Item("mtCoutas").Specific;

                                if (matrixCuotas.RowCount == 0)
                                {
                                    SBO_Application.StatusBar.SetText("Debe generar las cuotas para crear este documento", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                    BubbelEvent = false;
                                }
                            }
                            if (BubbelEvent && oForm.Mode == BoFormMode.fm_ADD_MODE)
                                BubbelEvent = crearAsiento();
                        }
                    }                    
                }
                else if(pVal.EventType == BoEventTypes.et_VALIDATE)
                {
                    if (pVal.ItemUID.Equals("mtCoutas"))
                    {
                        Matrix matrixCuotas = (Matrix)oForm.Items.Item("mtCoutas").Specific;
                        if (this.capital != double.Parse(((EditText)matrixCuotas.GetCellSpecific(pVal.ColUID, pVal.Row)).Value, CultureInfo.InvariantCulture))
                        {
                            if (SBO_Application.MessageBox("Se recalcularán las cuotas, de sea continuar?", 1, "Continuar", "Cancelar") == 1)
                            {
                                
                                double interes = double.Parse(oForm.DataSources.DBDataSources.Item(DATASOURCELINE).GetValue("U_HCO_Interes", pVal.Row - 1), CultureInfo.InvariantCulture);
                                double comision = double.Parse(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_Commission", 0), CultureInfo.InvariantCulture);
                                double seguros = double.Parse(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_Insuran", 0), CultureInfo.InvariantCulture);
                                double otros = double.Parse(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_Other", 0), CultureInfo.InvariantCulture);
                                double saldoInicial = double.Parse(oForm.DataSources.DBDataSources.Item(DATASOURCELINE).GetValue("U_HCO_InitalAmt", pVal.Row - 1), CultureInfo.InvariantCulture);
                                double pago = (double.Parse(((EditText)matrixCuotas.GetCellSpecific(pVal.ColUID, pVal.Row)).Value, CultureInfo.InvariantCulture)) + interes + (pVal.Row == 1 ? (saldoInicial * comision / 100) + seguros + otros : 0);
                                if((double.Parse(((EditText)matrixCuotas.GetCellSpecific(pVal.ColUID, pVal.Row)).Value, CultureInfo.InvariantCulture)) > saldoInicial)
                                {
                                    SBO_Application.StatusBar.SetText("El valor supera el saldo pendiente, ingrese un valor menor", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                    matrixCuotas.LoadFromDataSource();
                                    if (oForm.Mode != BoFormMode.fm_UPDATE_MODE)
                                        oForm.Mode = BoFormMode.fm_OK_MODE;
                                }
                                else
                                {
                                    matrixCuotas.FlushToDataSource();
                                    double capital = double.Parse(oForm.DataSources.DBDataSources.Item(DATASOURCELINE).GetValue("U_HCO_Capita", pVal.Row - 1), CultureInfo.InvariantCulture);
                                    oForm.DataSources.DBDataSources.Item(DATASOURCELINE).SetValue("U_HCO_PayAmt", pVal.Row - 1, pago.ToString(CultureInfo.InvariantCulture));
                                    oForm.DataSources.DBDataSources.Item(DATASOURCELINE).SetValue("U_HCO_FinalAmt", pVal.Row - 1, (saldoInicial - capital).ToString(CultureInfo.InvariantCulture));
                                    matrixCuotas.LoadFromDataSourceEx();
                                    if (oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_CalcType", 0).Equals("2"))
                                        calcularCoutas(true, pVal.Row);
                                    else
                                        calcularCoutasFijo(true, pVal.Row);
                                }
                            }
                            else
                            {
                                matrixCuotas.LoadFromDataSource();
                                if (oForm.Mode != BoFormMode.fm_UPDATE_MODE)
                                    oForm.Mode = BoFormMode.fm_OK_MODE;
                            }
                        }
                    }
                    else if (new string[] { "txtAmount", "txtIniDate", "txtMonths", "txtYRate", "txtRBase", "txtCom", "txtInsuran", "txtOther", "txtGraMth", "txtPayDay" }.Contains(pVal.ItemUID))
                    {
                        if(!pVal.InnerEvent && pVal.ItemChanged)
                        {
                            if(oForm.Mode == BoFormMode.fm_ADD_MODE)
                            {
                                Matrix matrixCuotas = (Matrix)oForm.Items.Item("mtCoutas").Specific;
                                if(matrixCuotas.RowCount > 0)
                                {
                                    matrixCuotas.Clear();
                                    matrixCuotas.FlushToDataSource();
                                }
                                if (pVal.ItemUID.Equals("txtIniDate"))
                                    cargarTasaCambio();
                            }
                            else if(oForm.Mode == BoFormMode.fm_UPDATE_MODE)
                            {
                                if (SBO_Application.MessageBox("Se recalcularán las cuotas, de sea continuar?", 1, "Continuar", "Cancelar") == 1)
                                {
                                    Matrix matrixCuotas = (Matrix)oForm.Items.Item("mtCoutas").Specific;
                                    int j = 0;
                                    for(int i = 0; i < matrixCuotas.RowCount; i++)
                                    {
                                        if (!string.IsNullOrEmpty(oForm.DataSources.DBDataSources.Item(DATASOURCELINE).GetValue("U_HCO_TransId", i)))
                                            j++;
                                    }
                                    if (int.Parse(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_Months", 0)) <= j)
                                    {
                                        SBO_Application.StatusBar.SetText("El nuevo plazo debe ser mayor a las cuotas ya pagadas");
                                        this.oForm.DataSources.DBDataSources.Item(DATASOURCE).SetValue("U_HCO_Months", 0, this.plazo.ToString(CultureInfo.InvariantCulture));
                                        this.oForm.DataSources.DBDataSources.Item(DATASOURCE).SetValue("U_HCO_YearRate", 0, this.interes.ToString(CultureInfo.InvariantCulture));
                                        oForm.Mode = BoFormMode.fm_OK_MODE;
                                    }
                                    else
                                    {
                                        if (oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_CalcType", 0).Equals("2"))
                                            calcularCoutas(j== 0 ? false : true, j++);
                                        else
                                            calcularCoutasFijo(j == 0 ? false : true, j++);
                                    }                                   
                                }
                                else
                                {
                                    this.oForm.DataSources.DBDataSources.Item(DATASOURCE).SetValue("U_HCO_Months", 0, this.plazo.ToString(CultureInfo.InvariantCulture));
                                    this.oForm.DataSources.DBDataSources.Item(DATASOURCE).SetValue("U_HCO_YearRate", 0, this.interes.ToString(CultureInfo.InvariantCulture));
                                    oForm.Mode = BoFormMode.fm_OK_MODE;
                                }
                            }
                        }
                    }
                }
            }
            else
            {
                if (pVal.EventType == BoEventTypes.et_CHOOSE_FROM_LIST)
                {
                    if (pVal.ItemUID.Equals("txtCCode"))
                    {
                        IChooseFromListEvent chflarg = (IChooseFromListEvent)pVal;
                        if (chflarg.SelectedObjects != null)
                        {
                            DataTable dt = chflarg.SelectedObjects;
                            this.oForm.DataSources.DBDataSources.Item(DATASOURCE).SetValue("U_HCO_CardCode", 0, Convert.ToString(dt.GetValue("CardCode", 0)));
                            this.oForm.DataSources.DBDataSources.Item(DATASOURCE).SetValue("U_HCO_CardName", 0, Convert.ToString(dt.GetValue("CardName", 0)));                                                     
                        }
                    }
                    else if (pVal.ItemUID.Equals("txtCName"))
                    {
                        IChooseFromListEvent chflarg = (IChooseFromListEvent)pVal;
                        if (chflarg.SelectedObjects != null)
                        {
                            DataTable dt = chflarg.SelectedObjects;
                            this.oForm.DataSources.DBDataSources.Item(DATASOURCE).SetValue("U_HCO_CardCode", 0, Convert.ToString(dt.GetValue("CardCode", 0)));
                            this.oForm.DataSources.DBDataSources.Item(DATASOURCE).SetValue("U_HCO_CardName", 0, Convert.ToString(dt.GetValue("CardName", 0)));                                                     
                        }
                    }
                    else if (pVal.ItemUID.Equals("txtCurr"))
                    {
                        IChooseFromListEvent chflarg = (IChooseFromListEvent)pVal;
                        if (chflarg.SelectedObjects != null)
                        {
                            DataTable dt = chflarg.SelectedObjects;
                            this.oForm.DataSources.DBDataSources.Item(DATASOURCE).SetValue("U_HCO_Currency", 0, Convert.ToString(dt.GetValue("CurrCode", 0)));
                            cargarTasaCambio();
                        }
                    }
                    else if (pVal.ItemUID.Equals("txtAccount"))
                    {
                        IChooseFromListEvent chflarg = (IChooseFromListEvent)pVal;
                        if (chflarg.SelectedObjects != null)
                        {
                            DataTable dt = chflarg.SelectedObjects;
                            if (validarCuentas(Convert.ToString(dt.GetValue("AcctCode", 0)), "txtAccount"))
                            {
                                this.oForm.DataSources.DBDataSources.Item(DATASOURCE).SetValue("U_HCO_AcctCode", 0, Convert.ToString(dt.GetValue("AcctCode", 0)));
                                ((StaticText)this.oForm.Items.Item("lblAccount").Specific).Caption = Convert.ToString(dt.GetValue("AcctName", 0));
                                var moneda = Constantes.consultarCampo("ActCurr", "OACT", "\"AcctCode\"", Convert.ToString(dt.GetValue("AcctCode", 0)), ref this.oCompany);
                                if (!string.IsNullOrEmpty(moneda))
                                    this.oForm.DataSources.DBDataSources.Item(DATASOURCE).SetValue("U_HCO_Currency", 0, moneda);                                
                            }
                            else
                            {
                                SBO_Application.StatusBar.SetText("Esta cuenta ya fue usada en otro campo, escoja otra por favor.");
                            }
                        }
                    }
                    else if (pVal.ItemUID.Equals("txtAccBank"))
                    {
                        IChooseFromListEvent chflarg = (IChooseFromListEvent)pVal;
                        if (chflarg.SelectedObjects != null)
                        {
                            DataTable dt = chflarg.SelectedObjects;
                            if (validarCuentas(Convert.ToString(dt.GetValue("AcctCode", 0)), "txtAccBank"))
                            {
                                this.oForm.DataSources.DBDataSources.Item(DATASOURCE).SetValue("U_HCO_AccBank", 0, Convert.ToString(dt.GetValue("AcctCode", 0)));
                                ((StaticText)this.oForm.Items.Item("lblAccBank").Specific).Caption = Convert.ToString(dt.GetValue("AcctName", 0));                             
                            }
                            else
                            {
                                SBO_Application.StatusBar.SetText("Esta cuenta ya fue usada en otro campo, escoja otra por favor.");
                            }
                        }
                    }
                    else if (pVal.ItemUID.Equals("txtAccIns"))
                    {
                        IChooseFromListEvent chflarg = (IChooseFromListEvent)pVal;
                        if (chflarg.SelectedObjects != null)
                        {
                            DataTable dt = chflarg.SelectedObjects;
                            if (validarCuentas(Convert.ToString(dt.GetValue("AcctCode", 0)), "txtAccIns"))
                            {
                                this.oForm.DataSources.DBDataSources.Item(DATASOURCE).SetValue("U_HCO_AccIns", 0, Convert.ToString(dt.GetValue("AcctCode", 0)));
                                ((StaticText)this.oForm.Items.Item("lblAccIns").Specific).Caption = Convert.ToString(dt.GetValue("AcctName", 0));                                                              
                            }
                            else
                            {
                                SBO_Application.StatusBar.SetText("Esta cuenta ya fue usada en otro campo, escoja otra por favor.");
                            }
                        }
                    }
                    else if (pVal.ItemUID.Equals("txtAccComm"))
                    {
                        IChooseFromListEvent chflarg = (IChooseFromListEvent)pVal;
                        if (chflarg.SelectedObjects != null)
                        {
                            DataTable dt = chflarg.SelectedObjects;
                            if (validarCuentas(Convert.ToString(dt.GetValue("AcctCode", 0)), "txtAccComm"))
                            {
                                this.oForm.DataSources.DBDataSources.Item(DATASOURCE).SetValue("U_HCO_AccComm", 0, Convert.ToString(dt.GetValue("AcctCode", 0)));
                                ((StaticText)this.oForm.Items.Item("lblAccComm").Specific).Caption = Convert.ToString(dt.GetValue("AcctName", 0));                                
                                
                            }
                            else
                            {
                                SBO_Application.StatusBar.SetText("Esta cuenta ya fue usada en otro campo, escoja otra por favor.");
                            }
                        }
                    }
                    else if (pVal.ItemUID.Equals("txtAccTran"))
                    {
                        IChooseFromListEvent chflarg = (IChooseFromListEvent)pVal;
                        if (chflarg.SelectedObjects != null)
                        {
                            DataTable dt = chflarg.SelectedObjects;
                            if (validarCuentas(Convert.ToString(dt.GetValue("AcctCode", 0)), "txtAccTran"))
                            {
                                this.oForm.DataSources.DBDataSources.Item(DATASOURCE).SetValue("U_HCO_AccTran", 0, Convert.ToString(dt.GetValue("AcctCode", 0)));
                                ((StaticText)this.oForm.Items.Item("lblAccTran").Specific).Caption = Convert.ToString(dt.GetValue("AcctName", 0));                                                               
                            }
                            else
                            {
                                SBO_Application.StatusBar.SetText("Esta cuenta ya fue usada en otro campo, escoja otra por favor.");
                            }
                        }
                    }
                    else if (pVal.ItemUID.Equals("txtAccOthe"))
                    {
                        IChooseFromListEvent chflarg = (IChooseFromListEvent)pVal;
                        if (chflarg.SelectedObjects != null)
                        {
                            DataTable dt = chflarg.SelectedObjects;
                            if (validarCuentas(Convert.ToString(dt.GetValue("AcctCode", 0)), "txtAccOthe"))
                            {
                                this.oForm.DataSources.DBDataSources.Item(DATASOURCE).SetValue("U_HCO_AccOthe", 0, Convert.ToString(dt.GetValue("AcctCode", 0)));
                                ((StaticText)this.oForm.Items.Item("lblAccOthe").Specific).Caption = Convert.ToString(dt.GetValue("AcctName", 0));                               
                            }
                            else
                            {
                                SBO_Application.StatusBar.SetText("Esta cuenta ya fue usada en otro campo, escoja otra por favor.");
                            }
                        }
                    }
                    else if (pVal.ItemUID.Equals("txtOcrCode"))
                    {
                        IChooseFromListEvent chflarg = (IChooseFromListEvent)pVal;
                        if (chflarg.SelectedObjects != null)
                        {
                            DataTable dt = chflarg.SelectedObjects;
                            this.oForm.DataSources.DBDataSources.Item(DATASOURCE).SetValue("U_HCO_OcrCode", 0, Convert.ToString(dt.GetValue("PrcCode", 0)));
                        }
                    }
                }
                else if (pVal.EventType == BoEventTypes.et_ITEM_PRESSED)
                {
                    if (pVal.ItemUID.Equals("btnCanc"))
                    {
                        this.oForm.Close();
                    }
                    else if (pVal.ItemUID.Equals("btnCalCuo"))
                    {
                        if (validarCampos())
                        {
                            if(string.IsNullOrEmpty(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_CalcType", 0)))                            
                                oForm.DataSources.DBDataSources.Item(DATASOURCE).SetValue("U_HCO_CalcType", 0, "2");
                            if (string.IsNullOrEmpty(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_RateBase", 0)))
                                oForm.DataSources.DBDataSources.Item(DATASOURCE).SetValue("U_HCO_RateBase", 0, "1");
                            if (string.IsNullOrEmpty(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_GraMth", 0)))
                                oForm.DataSources.DBDataSources.Item(DATASOURCE).SetValue("U_HCO_GraMth", 0, "0");
                            if (string.IsNullOrEmpty(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_PayDay", 0)))
                                oForm.DataSources.DBDataSources.Item(DATASOURCE).SetValue("U_HCO_PayDay", 0, DateTime.ParseExact(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_IniDate", 0), "yyyyMMdd", CultureInfo.InvariantCulture).Day.ToString());

                            //Se valida si es capital fijo o capital variable.
                            if (oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_CalcType", 0).Equals("2"))
                                calcularCoutas(false, 0);
                            else
                                //Capital variable.
                                calcularCoutasFijo(false, 0);
                            if (this.oForm.Mode != BoFormMode.fm_ADD_MODE)
                                this.oForm.Mode = BoFormMode.fm_UPDATE_MODE;
                        }
                    }
                    else if (pVal.ItemUID.Equals("1"))
                    {
                        if (oForm.Mode == BoFormMode.fm_ADD_MODE)
                        {
                            cargarDocNum();
                            cargarCuentaTransitoria();
                        }
                    }
                    else if (pVal.ItemUID.Equals("tabCuotas"))
                    {
                        this.oForm.PaneLevel = 1;
                    }
                    else if (pVal.ItemUID.Equals("tabFinan"))
                    {
                        this.oForm.PaneLevel = 2;
                    }
                }
                else if(pVal.EventType == BoEventTypes.et_GOT_FOCUS)
                {
                    if (pVal.ItemUID.Equals("mtCoutas"))
                    {
                        if (pVal.ColUID.Equals("txtCapita"))
                        {
                            Matrix matrixCuotas = (Matrix)oForm.Items.Item("mtCoutas").Specific;
                            this.capital = double.Parse(((EditText)matrixCuotas.GetCellSpecific(pVal.ColUID, pVal.Row)).Value, CultureInfo.InvariantCulture);
                        }
                    }  
                    else if(new string[] { "txtMonths", "txtYearRate" }.Contains(pVal.ItemUID))
                    {
                        if(!string.IsNullOrEmpty(this.oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_Months", 0)))
                            this.plazo = int.Parse(this.oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_Months", 0), CultureInfo.InvariantCulture);
                        this.interes = double.Parse(this.oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_YearRate", 0), CultureInfo.InvariantCulture);                        
                    }
                }
                else if(pVal.EventType == BoEventTypes.et_COMBO_SELECT)
                {
                    if(new string[] { "cbCalType", "cbPeriodi", "txtRBase" }.Contains(pVal.ItemUID))
                    {
                        if (!pVal.InnerEvent && pVal.ItemChanged)
                        {
                            if (oForm.Mode == BoFormMode.fm_ADD_MODE)
                            {
                                //Validación del campo periodicidad, se compara con fecha y plazo en meses para confirmar que sea coherente con los meses ingresados
                                if(pVal.ItemUID.Equals("cbPeriodi"))
                                {
                                    if(string.IsNullOrEmpty(this.oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_Months", 0)))
                                    {
                                        SBO_Application.StatusBar.SetText("Debe ingresar el plazo en meses primero", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                        oForm.Items.Item("txtMonths").Click(BoCellClickType.ct_Regular);
                                    }
                                    else if (string.IsNullOrEmpty(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_IniDate", 0)))
                                    {
                                        SBO_Application.StatusBar.SetText("Debe ingresar la fecha de inicio primero", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                        oForm.Items.Item("txtIniDate").Click(BoCellClickType.ct_Regular);
                                    }
                                    else
                                    {
                                        this.plazo = int.Parse(this.oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_Months", 0), CultureInfo.InvariantCulture);
                                        DateTime fechaInicial = DateTime.ParseExact(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_IniDate", 0), "yyyyMMdd", CultureInfo.InvariantCulture);
                                        string periodicidad = oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_Periodi", 0);
                                        if (periodicidad.Equals("T"))
                                        {
                                            var calculo = (int)(((fechaInicial.AddMonths(plazo).Year * 12 + fechaInicial.AddMonths(plazo).Month) - (fechaInicial.Year * 12 + fechaInicial.Month)) / 3);
                                            if( calculo == 0)
                                            {
                                                SBO_Application.StatusBar.SetText("El plazo en meses debe ser igual o mayor a un trimestre", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                                oForm.DataSources.DBDataSources.Item(DATASOURCE).SetValue("U_HCO_Periodi", 0, "M");
                                            }
                                        }
                                        else if (periodicidad.Equals("SE"))
                                        {
                                            var calculo = (int)(((fechaInicial.AddMonths(plazo).Year * 12 + fechaInicial.AddMonths(plazo).Month) - (fechaInicial.Year * 12 + fechaInicial.Month)) / 6);
                                            if (calculo == 0)
                                            {
                                                SBO_Application.StatusBar.SetText("El plazo en meses debe ser igual o mayor a un semestre", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                                oForm.DataSources.DBDataSources.Item(DATASOURCE).SetValue("U_HCO_Periodi", 0, "M");
                                            }
                                        }
                                        else if (periodicidad.Equals("A"))
                                        {
                                            var calculo = (int)(((fechaInicial.AddMonths(plazo).Year * 12 + fechaInicial.AddMonths(plazo).Month) - (fechaInicial.Year * 12 + fechaInicial.Month)) / 12);
                                            if (calculo == 0)
                                            {
                                                SBO_Application.StatusBar.SetText("El plazo en meses debe ser igual o mayor a un año", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                                oForm.DataSources.DBDataSources.Item(DATASOURCE).SetValue("U_HCO_Periodi", 0, "M");
                                            }
                                        }
                                    }
                                }
                                //Limpia la matriz para que se haga el calculo de nuevo.
                                Matrix matrixCuotas = (Matrix)oForm.Items.Item("mtCoutas").Specific;
                                if (matrixCuotas.RowCount > 0)
                                {
                                    matrixCuotas.Clear();
                                    matrixCuotas.FlushToDataSource();
                                }
                            }
                        }
                    }
                }
            }
        }

        public void ManejarEventosFormData(ref BusinessObjectInfo oBusinessInfo, ref bool pBubbleEvent)
        {
            if (oBusinessInfo.ActionSuccess)
            {
                this.oForm = SBO_Application.Forms.ActiveForm;
                if (!oBusinessInfo.BeforeAction)
                {
                    if (oBusinessInfo.EventType == BoEventTypes.et_FORM_DATA_LOAD)
                    {
                        Matrix matrix = (Matrix)oForm.Items.Item("mtCoutas").Specific;
                        int j = 0;
                        for (int i = 1; i <= matrix.RowCount; i++)
                        {
                            Cell oItem = matrix.Columns.Item("txtTransId").Cells.Item(i);
                            EditText editText = (EditText)oItem.Specific;
                            Cell oItemRdr = matrix.Columns.Item("txtREntry").Cells.Item(i);
                            EditText editTextRdr = (EditText)oItemRdr.Specific;
                            var commonSetting = matrix.CommonSetting;
                            commonSetting.SetCellEditable(i, 4, string.IsNullOrEmpty(editText.Value) && string.IsNullOrEmpty(editTextRdr.Value) ? true : false);
                            if (!string.IsNullOrEmpty(oForm.DataSources.DBDataSources.Item(DATASOURCELINE).GetValue("U_HCO_TransId", i - 1)) || !string.IsNullOrEmpty(oForm.DataSources.DBDataSources.Item(DATASOURCELINE).GetValue("U_HCO_RDREntry", i - 1)))
                                j++;
                        }

                        if (j == matrix.RowCount)
                        {
                            oForm.Items.Item("txtYRate").Enabled = false;
                            oForm.Items.Item("txtMonths").Enabled = false;
                        }
                        else
                        {
                            oForm.Items.Item("txtYRate").Enabled = true;
                            oForm.Items.Item("txtMonths").Enabled = true;
                        }

                        var nombre = Constantes.consultarCampo("AcctName", "OACT", "\"AcctCode\"", oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_AcctCode", 0), ref this.oCompany);
                        ((StaticText)this.oForm.Items.Item("lblAccount").Specific).Caption = nombre;

                        nombre = Constantes.consultarCampo("AcctName", "OACT", "\"AcctCode\"", oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_AccBank", 0), ref this.oCompany);
                        ((StaticText)this.oForm.Items.Item("lblAccBank").Specific).Caption = nombre;

                        nombre = Constantes.consultarCampo("AcctName", "OACT", "\"AcctCode\"", oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_AccTran", 0), ref this.oCompany);
                        ((StaticText)this.oForm.Items.Item("lblAccTran").Specific).Caption = nombre;

                        nombre = Constantes.consultarCampo("AcctName", "OACT", "\"AcctCode\"", oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_AccIns", 0), ref this.oCompany);
                        ((StaticText)this.oForm.Items.Item("lblAccIns").Specific).Caption = nombre;

                        nombre = Constantes.consultarCampo("AcctName", "OACT", "\"AcctCode\"", oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_AccComm", 0), ref this.oCompany);
                        ((StaticText)this.oForm.Items.Item("lblAccComm").Specific).Caption = nombre;

                        nombre = Constantes.consultarCampo("AcctName", "OACT", "\"AcctCode\"", oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_AccOthe", 0), ref this.oCompany);
                        ((StaticText)this.oForm.Items.Item("lblAccOthe").Specific).Caption = nombre;
                    }
                }
            }
        }

        private bool validarCampos()
        {
            bool BubbelEvent = true;
            if (string.IsNullOrEmpty(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_CardCode", 0)))
            {
                SBO_Application.StatusBar.SetText("Debe seleccionar el socio de negocios", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                oForm.Items.Item("txtCCode").Click(BoCellClickType.ct_Regular);
                BubbelEvent = false;
            }
            else if (double.Parse(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_Amount", 0), CultureInfo.InvariantCulture) == 0)
            {
                SBO_Application.StatusBar.SetText("Debe ingresar el importe", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                oForm.Items.Item("txtAmount").Click(BoCellClickType.ct_Regular);
                BubbelEvent = false;
            }
            else if (double.Parse(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_YearRate", 0), CultureInfo.InvariantCulture) == 0)
            {
                SBO_Application.StatusBar.SetText("Debe ingresar el interes anual", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                oForm.Items.Item("txtYRate").Click(BoCellClickType.ct_Regular);
                BubbelEvent = false;
            }            
            else if (string.IsNullOrEmpty(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_IniDate", 0)))
            {
                SBO_Application.StatusBar.SetText("Debe ingresar la fecha de inicio", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                oForm.Items.Item("txtIniDate").Click(BoCellClickType.ct_Regular);
                BubbelEvent = false;
            }
            else if (string.IsNullOrEmpty(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_Months", 0)))
            {
                SBO_Application.StatusBar.SetText("Debe ingresar el plazo en meses ", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                oForm.Items.Item("txtMonths").Click(BoCellClickType.ct_Regular);
                BubbelEvent = false;
            }
            else if (string.IsNullOrEmpty(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_Currency", 0)))
            {
                SBO_Application.StatusBar.SetText("Debe seleccionar la moneda", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                oForm.Items.Item("txtCurr").Click(BoCellClickType.ct_Regular);
                BubbelEvent = false;
            }
            else if (string.IsNullOrEmpty(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_OcrCode", 0)))
            {
                SBO_Application.StatusBar.SetText("Debe seleccionar el centro de costos", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                oForm.Items.Item("txtOcrCode").Click(BoCellClickType.ct_Regular);
                BubbelEvent = false;
            }
            else if (string.IsNullOrEmpty(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_AcctCode", 0)))
            {
                SBO_Application.StatusBar.SetText("Debe seleccionar la cuenta contable", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                this.oForm.PaneLevel = 2;
                oForm.Items.Item("txtAccount").Click(BoCellClickType.ct_Regular);
                BubbelEvent = false;
            }
            else if (string.IsNullOrEmpty(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_AccBank", 0)))
            {
                SBO_Application.StatusBar.SetText("Debe seleccionar la cuenta del banco", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                this.oForm.PaneLevel = 2;
                oForm.Items.Item("txtAccBank").Click(BoCellClickType.ct_Regular);
                BubbelEvent = false;
            }
            else if (string.IsNullOrEmpty(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_AccTran", 0)))
            {
                SBO_Application.StatusBar.SetText("Debe seleccionar la cuenta transitoria", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                this.oForm.PaneLevel = 2;
                oForm.Items.Item("txtAccTran").Click(BoCellClickType.ct_Regular);
                BubbelEvent = false;
            }
            else if (string.IsNullOrEmpty(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_AccIns", 0)) && double.Parse(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_Insuran", 0), CultureInfo.InvariantCulture) > 0)
            {
                SBO_Application.StatusBar.SetText("Debe seleccionar la cuenta de seguros", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                this.oForm.PaneLevel = 2;
                oForm.Items.Item("txtAccIns").Click(BoCellClickType.ct_Regular);
                BubbelEvent = false;
            }
            else if (string.IsNullOrEmpty(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_AccComm", 0)) && double.Parse(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_Commission", 0), CultureInfo.InvariantCulture) > 0)
            {
                SBO_Application.StatusBar.SetText("Debe seleccionar la cuenta de comisión", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                this.oForm.PaneLevel = 2;
                oForm.Items.Item("txtAccComm").Click(BoCellClickType.ct_Regular);
                BubbelEvent = false;
            }
            else if (string.IsNullOrEmpty(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_AccOthe", 0)) && double.Parse(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_Other", 0), CultureInfo.InvariantCulture) > 0)
            {
                SBO_Application.StatusBar.SetText("Debe seleccionar la cuenta de otros", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                this.oForm.PaneLevel = 2;
                oForm.Items.Item("txtAccOthe").Click(BoCellClickType.ct_Regular);
                BubbelEvent = false;
            }

            return BubbelEvent;
        }

        private void calcularCoutas(bool recalculo, int linea)
        {
            Matrix matrixCuotas = (Matrix)this.oForm.Items.Item("mtCoutas").Specific;

            int plazoReal = int.Parse(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_Months", 0));
            int plazo = int.Parse(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_Months", 0), CultureInfo.InvariantCulture);
            int mesesGracia = 0;
            if(!string.IsNullOrEmpty(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_GraMth", 0)))
                mesesGracia = int.Parse(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_GraMth", 0));
            DateTime fechaOriginal = DateTime.ParseExact(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_IniDate", 0), "yyyyMMdd", CultureInfo.InvariantCulture);
            DateTime fechaInicial = DateTime.ParseExact(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_IniDate", 0), "yyyyMMdd", CultureInfo.InvariantCulture);
            int diaPago = fechaOriginal.Day;
            if (!string.IsNullOrEmpty(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_PayDay", 0)))
                diaPago = int.Parse(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_PayDay", 0));
            fechaInicial = new DateTime(fechaInicial.Year, fechaInicial.Month, diaPago);
            fechaInicial = fechaInicial.AddMonths(mesesGracia);
            double importe = double.Parse(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_Amount", 0), CultureInfo.InvariantCulture);
            double tasa = double.Parse(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_YearRate", 0), CultureInfo.InvariantCulture);
            double tasaOriginal = tasa;
            double tasaInteresDiario = 0;
            string baseTasa = oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_RateBase", 0);
            double comision = double.Parse(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_Commission", 0), CultureInfo.InvariantCulture);
            double seguros = double.Parse(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_Insuran", 0), CultureInfo.InvariantCulture);
            double otros = double.Parse(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_Other", 0), CultureInfo.InvariantCulture);
            string periodicidad = oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_Periodi", 0);
            double valorPagado = 0;
            double realPagado = 0;
            double valorComision = 0;
            int diasInteres = 0;
            double totalPrestamo = 0;

            if (periodicidad.Equals("D"))
            {
                fechaInicial = fechaOriginal.AddDays(1);
                if (baseTasa == "1")
                    plazo = 30 * plazo;
                else
                    plazo = (int)(fechaInicial.AddMonths(plazo) - fechaInicial).TotalDays;                
                tasa = tasa / 100 / (baseTasa == "1" ? 360 : 365);
                diasInteres = 1;
            }
            else if (periodicidad.Equals("S"))
            {
                fechaInicial = fechaInicial.AddDays(7);                
                if (baseTasa == "1")
                    plazo = 30 * plazo / 7;                
                else                
                    plazo = (int)((fechaInicial.AddMonths(plazo) - fechaInicial).TotalDays / 7);                
                tasa = tasa / 100 / (baseTasa == "1" ? 360 : 365) * 7;
            }
            if (periodicidad.Equals("T"))
            {
                fechaInicial = fechaInicial.AddMonths(3);
                plazo = (int)(((fechaInicial.AddMonths(plazo).Year * 12 + fechaInicial.AddMonths(plazo).Month) - (fechaInicial.Year * 12 + fechaInicial.Month)) / 3);
                tasa = tasa / 100 / 12 * 3;
                if (plazo == 0)
                {
                    SBO_Application.StatusBar.SetText("El plazo en meses debe ser igual o mayor a un trimestre", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oForm.DataSources.DBDataSources.Item(DATASOURCE).SetValue("U_HCO_Periodi", 0, "M");
                    return;
                }                
            }
            else if (periodicidad.Equals("SE"))
            {
                fechaInicial = fechaInicial.AddMonths(6);
                plazo = (int)(((fechaInicial.AddMonths(plazo).Year * 12 + fechaInicial.AddMonths(plazo).Month) - (fechaInicial.Year * 12 + fechaInicial.Month)) / 6);
                tasa = tasa / 100 / 12 * 6;
                if (plazo == 0)
                {
                    SBO_Application.StatusBar.SetText("El plazo en meses debe ser igual o mayor a un semestre", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oForm.DataSources.DBDataSources.Item(DATASOURCE).SetValue("U_HCO_Periodi", 0, "M");
                    return;
                }
            }
            else if (periodicidad.Equals("A"))
            {
                fechaInicial = fechaInicial.AddMonths(12);
                plazo = (int)(((fechaInicial.AddMonths(plazo).Year * 12 + fechaInicial.AddMonths(plazo).Month) - (fechaInicial.Year * 12 + fechaInicial.Month)) / 12);
                if (plazo == 0)
                {
                    SBO_Application.StatusBar.SetText("El plazo en meses debe ser igual o mayor a un año", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oForm.DataSources.DBDataSources.Item(DATASOURCE).SetValue("U_HCO_Periodi", 0, "M");
                    return;
                }
                tasa = tasa / 100;
            }
            else if (periodicidad.Equals("M"))
            {
                fechaInicial = fechaInicial.AddMonths(1);
                tasa = tasa / 100 / 12;          
            }

            if (baseTasa == "1")
            {
                plazo = 30 * Math.Abs((fechaInicial.Month - fechaOriginal.Month) + 12 * (fechaInicial.Year - fechaOriginal.Year));
                plazo += (fechaOriginal - new DateTime(fechaOriginal.Year, fechaOriginal.Month, diaPago)).Days;
            }
            else
                diasInteres = (fechaInicial - fechaOriginal).Days;

            

            if (recalculo)
            {
                for (int i = 0; i < linea; i++)
                {
                    importe = importe - double.Parse(oForm.DataSources.DBDataSources.Item(DATASOURCELINE).GetValue("U_HCO_Capita", i), CultureInfo.InvariantCulture);
                    valorPagado += double.Parse(oForm.DataSources.DBDataSources.Item(DATASOURCELINE).GetValue("U_HCO_PayAmt", i), CultureInfo.InvariantCulture);
                    if (!string.IsNullOrEmpty(oForm.DataSources.DBDataSources.Item(DATASOURCELINE).GetValue("U_HCO_TransId", i)) || !string.IsNullOrEmpty(oForm.DataSources.DBDataSources.Item(DATASOURCELINE).GetValue("U_HCO_RDREntry", i)))
                        realPagado += double.Parse(oForm.DataSources.DBDataSources.Item(DATASOURCELINE).GetValue("U_HCO_PayAmt", i), CultureInfo.InvariantCulture);
                }
                plazo = plazo - linea;
            }

            if (linea == 0)
                valorComision = double.Parse(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_Amount", 0), CultureInfo.InvariantCulture) * comision / 100;

            tasaInteresDiario = tasaOriginal / (baseTasa == "1" ? 360 : 365)  / 100;

            double cuota = 1 + tasa;
            cuota = Math.Pow(cuota, plazo);
            cuota = tasa * cuota;
            cuota = cuota * importe;
            cuota = cuota / (Math.Pow((1 + tasa), plazo) - 1);

            oForm.DataSources.DBDataSources.Item(DATASOURCE).SetValue("U_HCO_MnthPay", 0, cuota.ToString(CultureInfo.InvariantCulture));

            if (!recalculo)
            {
                matrixCuotas.Clear();
                matrixCuotas.AddRow(plazo);
                matrixCuotas.FlushToDataSource();
            }
            else
            {
                if(plazoReal - matrixCuotas.RowCount > 0)
                {
                    matrixCuotas.AddRow(plazoReal - matrixCuotas.RowCount);
                    matrixCuotas.FlushToDataSource();
                }
                fechaInicial = DateTime.ParseExact(oForm.DataSources.DBDataSources.Item(DATASOURCELINE).GetValue("U_HCO_Date", linea == 0 ? 0 : linea - 1), "yyyyMMdd", CultureInfo.InvariantCulture);
                if (periodicidad.Equals("D"))
                {
                    fechaInicial = fechaInicial.AddDays(1);
                    diasInteres = 1;
                }
                else if (periodicidad.Equals("S"))
                {
                    fechaInicial = fechaInicial.AddDays(7);
                    diasInteres = 7;
                }
                else if (periodicidad.Equals("M"))
                {
                    fechaInicial = fechaInicial.AddMonths(1);
                    diasInteres = (baseTasa == "1") ? 30 : (fechaInicial - fechaOriginal).Days;
                }
                else if (periodicidad.Equals("T"))
                {
                    fechaInicial = fechaInicial.AddMonths(3);
                    diasInteres = (baseTasa == "1") ? 90 : (fechaInicial - fechaOriginal).Days;
                }
                else if (periodicidad.Equals("SE"))
                {
                    fechaInicial = fechaInicial.AddMonths(6);
                    diasInteres = (baseTasa == "1") ? 180 : (fechaInicial - fechaOriginal).Days;
                }
                else if (periodicidad.Equals("A"))
                {
                    fechaInicial = fechaInicial.AddYears(1);
                    diasInteres = (baseTasa == "1") ? 360 : (fechaInicial - fechaOriginal).Days;
                }
            }

            double saldoInicial = importe;

            var commonSetting = matrixCuotas.CommonSetting;

            for (int i = linea + 1; i <= plazo + linea; i++)
            {                                
                double intereses = tasaInteresDiario * saldoInicial * diasInteres;

                commonSetting.SetCellEditable(i, 4, true);
                oForm.DataSources.DBDataSources.Item(DATASOURCELINE).SetValue("LineId", i - 1, i.ToString());
                oForm.DataSources.DBDataSources.Item(DATASOURCELINE).SetValue("U_HCO_Date", i - 1, fechaInicial.ToString("yyyyMMdd"));
                oForm.DataSources.DBDataSources.Item(DATASOURCELINE).SetValue("U_HCO_InitalAmt", i - 1, saldoInicial.ToString(CultureInfo.InvariantCulture));
                oForm.DataSources.DBDataSources.Item(DATASOURCELINE).SetValue("U_HCO_PayAmt", i - 1, i == 1 ? (cuota + valorComision + seguros + otros).ToString(CultureInfo.InvariantCulture) : cuota.ToString(CultureInfo.InvariantCulture));
                oForm.DataSources.DBDataSources.Item(DATASOURCELINE).SetValue("U_HCO_Capita", i - 1, (cuota - (intereses)).ToString(CultureInfo.InvariantCulture));
                oForm.DataSources.DBDataSources.Item(DATASOURCELINE).SetValue("U_HCO_Interes", i - 1, intereses.ToString(CultureInfo.InvariantCulture));
                oForm.DataSources.DBDataSources.Item(DATASOURCELINE).SetValue("U_HCO_FinalAmt", i - 1, (saldoInicial - (cuota - intereses)).ToString(CultureInfo.InvariantCulture));

                fechaOriginal = fechaInicial;

                if (periodicidad.Equals("D"))
                {
                    fechaInicial = fechaInicial.AddDays(1);
                    diasInteres = 1;
                }
                else if (periodicidad.Equals("S"))
                {
                    fechaInicial = fechaInicial.AddDays(7);
                    diasInteres = 7;
                }
                else if (periodicidad.Equals("M"))
                {
                    fechaInicial = fechaInicial.AddMonths(1);
                    diasInteres = (baseTasa == "1") ? 30 : (fechaInicial - fechaOriginal).Days;
                }
                else if (periodicidad.Equals("T"))
                {
                    fechaInicial = fechaInicial.AddMonths(3);
                    diasInteres = (baseTasa == "1") ? 90 : (fechaInicial - fechaOriginal).Days;
                }
                else if (periodicidad.Equals("SE"))
                {
                    fechaInicial = fechaInicial.AddMonths(6);
                    diasInteres = (baseTasa == "1") ? 180 : (fechaInicial - fechaOriginal).Days;
                }
                else if (periodicidad.Equals("A"))
                {
                    fechaInicial = fechaInicial.AddYears(1);
                    diasInteres = (baseTasa == "1") ? 360 : (fechaInicial - fechaOriginal).Days;
                }

                saldoInicial = (saldoInicial - (cuota - intereses));
                totalPrestamo += i == 1 ? (cuota + valorComision + seguros + otros) : cuota;
            }

            matrixCuotas.LoadFromDataSourceEx();

            if (recalculo)
            {
                while(plazo + linea < matrixCuotas.RowCount)
                {
                    matrixCuotas.DeleteRow(plazo + linea + 1);
                }
                matrixCuotas.FlushToDataSource();
            }
            
            double interesUlC = double.Parse(oForm.DataSources.DBDataSources.Item(DATASOURCELINE).GetValue("U_HCO_Interes", matrixCuotas.RowCount - 1), CultureInfo.InvariantCulture);
            saldoInicial = double.Parse(oForm.DataSources.DBDataSources.Item(DATASOURCELINE).GetValue("U_HCO_InitalAmt", matrixCuotas.RowCount - 1), CultureInfo.InvariantCulture);
            oForm.DataSources.DBDataSources.Item(DATASOURCELINE).SetValue("U_HCO_PayAmt", matrixCuotas.RowCount - 1, (saldoInicial + interesUlC).ToString(CultureInfo.InvariantCulture));
            oForm.DataSources.DBDataSources.Item(DATASOURCELINE).SetValue("U_HCO_Capita", matrixCuotas.RowCount - 1, saldoInicial.ToString(CultureInfo.InvariantCulture));
            oForm.DataSources.DBDataSources.Item(DATASOURCELINE).SetValue("U_HCO_FinalAmt", matrixCuotas.RowCount - 1, 0.ToString(CultureInfo.InvariantCulture));
            matrixCuotas.LoadFromDataSourceEx();
            totalPrestamo -= cuota;
            totalPrestamo += saldoInicial + interesUlC;
            oForm.DataSources.DBDataSources.Item(DATASOURCE).SetValue("U_HCO_DocTotal", 0, (totalPrestamo + valorPagado).ToString(CultureInfo.InvariantCulture));
            oForm.DataSources.DBDataSources.Item(DATASOURCE).SetValue("U_HCO_OpenBal", 0, (totalPrestamo + valorPagado - realPagado).ToString(CultureInfo.InvariantCulture));
            oForm.DataSources.DBDataSources.Item(DATASOURCE).SetValue("U_HCO_PaidToDate", 0, realPagado.ToString(CultureInfo.InvariantCulture));

        }

        private void calcularCoutasFijo(bool recalculo, int linea)
        {
            Matrix matrixCuotas = (Matrix)this.oForm.Items.Item("mtCoutas").Specific;

            int plazoReal = int.Parse(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_Months", 0));
            int plazo = int.Parse(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_Months", 0), CultureInfo.InvariantCulture);
            int diaPago = int.Parse(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_PayDay", 0));
            int mesesGracia = 0;
            if (!string.IsNullOrEmpty(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_GraMth", 0)))
                mesesGracia = int.Parse(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_GraMth", 0));
            DateTime fechaOriginal = DateTime.ParseExact(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_IniDate", 0), "yyyyMMdd", CultureInfo.InvariantCulture);
            DateTime fechaInicial = DateTime.ParseExact(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_IniDate", 0), "yyyyMMdd", CultureInfo.InvariantCulture);
            fechaInicial = new DateTime(fechaInicial.Year, fechaInicial.Month, diaPago);
            fechaInicial = fechaInicial.AddMonths(mesesGracia);
            double importe = double.Parse(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_Amount", 0), CultureInfo.InvariantCulture);
            double tasa = double.Parse(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_YearRate", 0), CultureInfo.InvariantCulture);
            double tasaOriginal = tasa;
            double tasaInteresDiario = 0;
            string baseTasa = oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_RateBase", 0);
            double comision = double.Parse(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_Commission", 0), CultureInfo.InvariantCulture);
            double seguros = double.Parse(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_Insuran", 0), CultureInfo.InvariantCulture);
            double otros = double.Parse(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_Other", 0), CultureInfo.InvariantCulture);
            string periodicidad = oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_Periodi", 0);
            double valorPagado = 0;
            double realPagado = 0;
            double valorComision = 0;
            int diasInteres = 0;

            if (periodicidad.Equals("D"))
            {
                fechaInicial = fechaOriginal.AddDays(1);
                if (baseTasa == "1")
                    plazo = 30 * plazo;
                else
                    plazo = (int)(fechaInicial.AddMonths(plazo) - fechaInicial).TotalDays;
                tasa = tasa / 100 / (baseTasa == "1" ? 360 : 365);
                diasInteres = 1;
            }
            else if (periodicidad.Equals("S"))
            {
                fechaInicial = fechaInicial.AddDays(7);
                if (baseTasa == "1")
                    plazo = 30 * plazo / 7;
                else
                    plazo = (int)((fechaInicial.AddMonths(plazo) - fechaInicial).TotalDays / 7);
                tasa = tasa / 100 / (baseTasa == "1" ? 360 : 365) * 7;
            }
            if (periodicidad.Equals("T"))
            {
                fechaInicial = fechaInicial.AddMonths(3);
                plazo = (int)(((fechaInicial.AddMonths(plazo).Year * 12 + fechaInicial.AddMonths(plazo).Month) - (fechaInicial.Year * 12 + fechaInicial.Month)) / 3);
                tasa = tasa / 100 / 12 * 3;
                if (plazo == 0)
                {
                    SBO_Application.StatusBar.SetText("El plazo en meses debe ser igual o mayor a un trimestre", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oForm.DataSources.DBDataSources.Item(DATASOURCE).SetValue("U_HCO_Periodi", 0, "M");
                    return;
                }
            }
            else if (periodicidad.Equals("SE"))
            {
                fechaInicial = fechaInicial.AddMonths(6);
                plazo = (int)(((fechaInicial.AddMonths(plazo).Year * 12 + fechaInicial.AddMonths(plazo).Month) - (fechaInicial.Year * 12 + fechaInicial.Month)) / 6);
                tasa = tasa / 100 / 12 * 6;
                if (plazo == 0)
                {
                    SBO_Application.StatusBar.SetText("El plazo en meses debe ser igual o mayor a un semestre", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oForm.DataSources.DBDataSources.Item(DATASOURCE).SetValue("U_HCO_Periodi", 0, "M");
                    return;
                }
            }
            else if (periodicidad.Equals("A"))
            {
                fechaInicial = fechaInicial.AddMonths(12);
                plazo = (int)(((fechaInicial.AddMonths(plazo).Year * 12 + fechaInicial.AddMonths(plazo).Month) - (fechaInicial.Year * 12 + fechaInicial.Month)) / 12);
                if (plazo == 0)
                {
                    SBO_Application.StatusBar.SetText("El plazo en meses debe ser igual o mayor a un año", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oForm.DataSources.DBDataSources.Item(DATASOURCE).SetValue("U_HCO_Periodi", 0, "M");
                    return;
                }
                tasa = tasa / 100;
            }
            else if (periodicidad.Equals("M"))
            {
                fechaInicial = fechaInicial.AddMonths(1);
                tasa = tasa / 100 / 12;
            }

            if (baseTasa == "1")
            {
                diasInteres = 30 * Math.Abs((fechaInicial.Month - fechaOriginal.Month) + 12 * (fechaInicial.Year - fechaOriginal.Year));
                diasInteres += (new DateTime(fechaOriginal.Year, fechaOriginal.Month, diaPago) - fechaOriginal).Days + 1;
            }
            else
                diasInteres = (fechaInicial - fechaOriginal).Days;

            if (recalculo)
            {
                for (int i = 0; i < linea; i++)
                {
                    importe = importe - double.Parse(oForm.DataSources.DBDataSources.Item(DATASOURCELINE).GetValue("U_HCO_Capita", i), CultureInfo.InvariantCulture);
                    valorPagado += double.Parse(oForm.DataSources.DBDataSources.Item(DATASOURCELINE).GetValue("U_HCO_PayAmt", i), CultureInfo.InvariantCulture);
                    if (!string.IsNullOrEmpty(oForm.DataSources.DBDataSources.Item(DATASOURCELINE).GetValue("U_HCO_TransId", i)) || !string.IsNullOrEmpty(oForm.DataSources.DBDataSources.Item(DATASOURCELINE).GetValue("U_HCO_RDREntry", i)))
                        realPagado += double.Parse(oForm.DataSources.DBDataSources.Item(DATASOURCELINE).GetValue("U_HCO_PayAmt", i), CultureInfo.InvariantCulture);
                }
                plazo = plazo - linea;
            }
            
            if(linea == 0)                
                valorComision = double.Parse(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_Amount", 0), CultureInfo.InvariantCulture) * comision / 100;

            tasaInteresDiario = tasaOriginal / (baseTasa == "1" ? 360 : 365) / 100;            

            double capital = importe / plazo;
            double valorTotal = 0;

            oForm.DataSources.DBDataSources.Item(DATASOURCE).SetValue("U_HCO_MnthPay", 0, capital.ToString(CultureInfo.InvariantCulture));

            if (!recalculo)
            {
                matrixCuotas.Clear();
                matrixCuotas.AddRow(plazo);
                matrixCuotas.FlushToDataSource();
            }
            else
            {
                if (plazoReal - matrixCuotas.RowCount > 0)
                {
                    matrixCuotas.AddRow(plazoReal - matrixCuotas.RowCount);
                    matrixCuotas.FlushToDataSource();
                }
                fechaInicial = DateTime.ParseExact(oForm.DataSources.DBDataSources.Item(DATASOURCELINE).GetValue("U_HCO_Date", linea == 0 ? 0 : linea - 1), "yyyyMMdd", CultureInfo.InvariantCulture);
                if (periodicidad.Equals("D"))
                {
                    fechaInicial = fechaInicial.AddDays(1);
                    diasInteres = 1;
                }
                else if (periodicidad.Equals("S"))
                {
                    fechaInicial = fechaInicial.AddDays(7);
                    diasInteres = 7;
                }
                else if (periodicidad.Equals("M"))
                {
                    fechaInicial = fechaInicial.AddMonths(1);
                    diasInteres = (baseTasa == "1") ? 30 : (fechaInicial - fechaOriginal).Days;
                }
                else if (periodicidad.Equals("T"))
                {
                    fechaInicial = fechaInicial.AddMonths(3);
                    diasInteres = (baseTasa == "1") ? 90 : (fechaInicial - fechaOriginal).Days;
                }
                else if (periodicidad.Equals("SE"))
                {
                    fechaInicial = fechaInicial.AddMonths(6);
                    diasInteres = (baseTasa == "1") ? 180 : (fechaInicial - fechaOriginal).Days;
                }
                else if (periodicidad.Equals("A"))
                {
                    fechaInicial = fechaInicial.AddYears(1);
                    diasInteres = (baseTasa == "1") ? 360 : (fechaInicial - fechaOriginal).Days;
                }
            }

            double saldoInicial = importe;

            var commonSetting = matrixCuotas.CommonSetting;

            for (int i = linea + 1; i <= plazo + linea; i++)
            {
                double intereses = tasaInteresDiario * saldoInicial * diasInteres;

                commonSetting.SetCellEditable(i, 4, true);

                oForm.DataSources.DBDataSources.Item(DATASOURCELINE).SetValue("LineId", i - 1, i.ToString());
                oForm.DataSources.DBDataSources.Item(DATASOURCELINE).SetValue("U_HCO_Date", i - 1, fechaInicial.ToString("yyyyMMdd"));
                oForm.DataSources.DBDataSources.Item(DATASOURCELINE).SetValue("U_HCO_InitalAmt", i - 1, saldoInicial.ToString(CultureInfo.InvariantCulture));
                oForm.DataSources.DBDataSources.Item(DATASOURCELINE).SetValue("U_HCO_PayAmt", i - 1, i == 1 ? (capital + intereses + valorComision + seguros + otros).ToString(CultureInfo.InvariantCulture) : (capital + intereses).ToString(CultureInfo.InvariantCulture));
                oForm.DataSources.DBDataSources.Item(DATASOURCELINE).SetValue("U_HCO_Capita", i - 1, capital.ToString(CultureInfo.InvariantCulture));
                oForm.DataSources.DBDataSources.Item(DATASOURCELINE).SetValue("U_HCO_Interes", i - 1, (intereses).ToString(CultureInfo.InvariantCulture));
                oForm.DataSources.DBDataSources.Item(DATASOURCELINE).SetValue("U_HCO_FinalAmt", i - 1, (saldoInicial - capital).ToString(CultureInfo.InvariantCulture));

                fechaOriginal = fechaInicial;

                if (periodicidad.Equals("D"))
                {
                    fechaInicial = fechaInicial.AddDays(1);
                    diasInteres = 1;
                }
                else if (periodicidad.Equals("S"))
                {
                    fechaInicial = fechaInicial.AddDays(7);
                    diasInteres = 7;
                }
                else if (periodicidad.Equals("M"))
                {
                    fechaInicial = fechaInicial.AddMonths(1);
                    diasInteres = (baseTasa == "1") ? 30 : (fechaInicial - fechaOriginal).Days;
                }
                else if (periodicidad.Equals("T"))
                {
                    fechaInicial = fechaInicial.AddMonths(3);
                    diasInteres = (baseTasa == "1") ? 90 : (fechaInicial - fechaOriginal).Days;
                }
                else if (periodicidad.Equals("SE"))
                {
                    fechaInicial = fechaInicial.AddMonths(6);
                    diasInteres = (baseTasa == "1") ? 180 : (fechaInicial - fechaOriginal).Days;
                }
                else if (periodicidad.Equals("A"))
                {
                    fechaInicial = fechaInicial.AddYears(1);
                    diasInteres = (baseTasa == "1") ? 360 : (fechaInicial - fechaOriginal).Days;
                }

                valorTotal += i == 1 ? (capital + intereses + valorComision + seguros + otros) : (capital + intereses);
                saldoInicial = (saldoInicial - capital);
            }

            matrixCuotas.LoadFromDataSourceEx();

            if (recalculo)
            {
                while (plazo + linea < matrixCuotas.RowCount)
                {
                    matrixCuotas.DeleteRow(plazo + linea + 1);
                }
                matrixCuotas.FlushToDataSource();
            }

            oForm.DataSources.DBDataSources.Item(DATASOURCELINE).SetValue("U_HCO_FinalAmt", matrixCuotas.RowCount - 1, 0.ToString(CultureInfo.InvariantCulture));
            matrixCuotas.LoadFromDataSourceEx();
            oForm.DataSources.DBDataSources.Item(DATASOURCE).SetValue("U_HCO_DocTotal", 0, (valorTotal + valorPagado).ToString(CultureInfo.InvariantCulture));
            oForm.DataSources.DBDataSources.Item(DATASOURCE).SetValue("U_HCO_OpenBal", 0, (valorTotal + valorPagado - realPagado).ToString(CultureInfo.InvariantCulture));
            oForm.DataSources.DBDataSources.Item(DATASOURCE).SetValue("U_HCO_PaidToDate", 0, realPagado.ToString(CultureInfo.InvariantCulture));
        }

        protected void inicializar()
        {
            cargarDocNum();
            cargarCuentaTransitoria();

            ChooseFromList objChooseFromList = oForm.ChooseFromLists.Item("cflAccount");
            Conditions conditions = objChooseFromList.GetConditions();
            Condition condition = conditions.Add();
            condition.BracketOpenNum = 1;
            condition.Alias = "Postable";
            condition.Operation = BoConditionOperation.co_EQUAL;
            condition.CondVal = "Y";
            condition.Relationship = BoConditionRelationship.cr_AND;
            condition = conditions.Add();
            condition.Alias = "LocManTran";
            condition.Operation = BoConditionOperation.co_EQUAL;
            condition.CondVal = "N";
            condition.Relationship = BoConditionRelationship.cr_AND;
            condition = conditions.Add();
            condition.Alias = DATASOURCE.Equals("@HCO_OOFR") ? "U_HCO_PasLoanAcct" : "U_HCO_ActLoanAcct";
            condition.Operation = BoConditionOperation.co_START;
            condition.CondVal = "Y";
            condition.BracketCloseNum = 1;
            objChooseFromList.SetConditions(conditions);

            objChooseFromList = oForm.ChooseFromLists.Item("cflAccBank");
            conditions = objChooseFromList.GetConditions();
            condition = conditions.Add();
            condition.BracketOpenNum = 1;
            condition.Alias = "Postable";
            condition.Operation = BoConditionOperation.co_EQUAL;
            condition.CondVal = "Y";
            condition.Relationship = BoConditionRelationship.cr_AND;
            condition = conditions.Add();
            condition.Alias = "LocManTran";
            condition.Operation = BoConditionOperation.co_EQUAL;
            condition.CondVal = "N";
            condition.BracketCloseNum = 1;
            objChooseFromList.SetConditions(conditions);

            objChooseFromList = oForm.ChooseFromLists.Item("cflAccTran");
            conditions = objChooseFromList.GetConditions();
            condition = conditions.Add();
            condition.BracketOpenNum = 1;
            condition.Alias = "Postable";
            condition.Operation = BoConditionOperation.co_EQUAL;
            condition.CondVal = "Y";
            condition.Relationship = BoConditionRelationship.cr_AND;
            condition = conditions.Add();
            condition.Alias = "LocManTran";
            condition.Operation = BoConditionOperation.co_EQUAL;
            condition.CondVal = "N";
            condition.BracketCloseNum = 1;
            objChooseFromList.SetConditions(conditions);

            objChooseFromList = oForm.ChooseFromLists.Item("cflCardCode");
            conditions = objChooseFromList.GetConditions();
            condition = conditions.Add();
            condition.BracketOpenNum = 1;
            condition.Alias = "CardType";
            condition.Operation = BoConditionOperation.co_EQUAL;
            condition.CondVal = DATASOURCE.Equals("@HCO_OOFR") ? "S" : "C";
            condition.BracketCloseNum = 1;
            objChooseFromList.SetConditions(conditions);

            objChooseFromList = oForm.ChooseFromLists.Item("cflCardName");
            conditions = objChooseFromList.GetConditions();
            condition = conditions.Add();
            condition.BracketOpenNum = 1;
            condition.Alias = "CardType";
            condition.Operation = BoConditionOperation.co_EQUAL;
            condition.CondVal = DATASOURCE.Equals("@HCO_OOFR") ? "S" : "C";
            condition.BracketCloseNum = 1;
            objChooseFromList.SetConditions(conditions);

            objChooseFromList = oForm.ChooseFromLists.Item("cflOcrCode");
            conditions = objChooseFromList.GetConditions();
            condition = conditions.Add();
            condition.BracketOpenNum = 1;
            condition.Alias = "DimCode";
            condition.Operation = BoConditionOperation.co_EQUAL;
            condition.CondVal = "1";
            condition.BracketCloseNum = 1;
            objChooseFromList.SetConditions(conditions);

            objChooseFromList = oForm.ChooseFromLists.Item("cflAccIns");
            conditions = objChooseFromList.GetConditions();
            condition = conditions.Add();
            condition.BracketOpenNum = 1;
            condition.Alias = "Postable";
            condition.Operation = BoConditionOperation.co_EQUAL;
            condition.CondVal = "Y";
            condition.Relationship = BoConditionRelationship.cr_AND;
            condition = conditions.Add();
            condition.Alias = "LocManTran";
            condition.Operation = BoConditionOperation.co_EQUAL;
            condition.CondVal = "N";
            condition.BracketCloseNum = 1;
            objChooseFromList.SetConditions(conditions);

            objChooseFromList = oForm.ChooseFromLists.Item("cflAccComm");
            conditions = objChooseFromList.GetConditions();
            condition = conditions.Add();
            condition.BracketOpenNum = 1;
            condition.Alias = "Postable";
            condition.Operation = BoConditionOperation.co_EQUAL;
            condition.CondVal = "Y";
            condition.Relationship = BoConditionRelationship.cr_AND;
            condition = conditions.Add();
            condition.Alias = "LocManTran";
            condition.Operation = BoConditionOperation.co_EQUAL;
            condition.CondVal = "N";
            condition.BracketCloseNum = 1;
            objChooseFromList.SetConditions(conditions);

            objChooseFromList = oForm.ChooseFromLists.Item("cflAccOthe");
            conditions = objChooseFromList.GetConditions();
            condition = conditions.Add();
            condition.BracketOpenNum = 1;
            condition.Alias = "Postable";
            condition.Operation = BoConditionOperation.co_EQUAL;
            condition.CondVal = "Y";
            condition.Relationship = BoConditionRelationship.cr_AND;
            condition = conditions.Add();
            condition.Alias = "LocManTran";
            condition.Operation = BoConditionOperation.co_EQUAL;
            condition.CondVal = "N";
            condition.BracketCloseNum = 1;
            objChooseFromList.SetConditions(conditions);

            activarCampos(true);

            System.Runtime.InteropServices.Marshal.ReleaseComObject(objChooseFromList);
            objChooseFromList = null;
            GC.Collect();
        }

        protected void activarCampos(bool activar)
        {
            this.oForm.Items.Item("txtCCode").Enabled = activar;
            this.oForm.Items.Item("txtCName").Enabled = activar;
            this.oForm.Items.Item("txtAmount").Enabled = activar;
            this.oForm.Items.Item("txtRBase").Enabled = activar;
            this.oForm.Items.Item("txtCom").Enabled = activar;
            this.oForm.Items.Item("txtAccount").Enabled = activar;
            this.oForm.Items.Item("txtIniDate").Enabled = activar;
            this.oForm.Items.Item("txtCurr").Enabled = activar;
            this.oForm.Items.Item("cbCalType").Enabled = activar;
            this.oForm.Items.Item("btnCalCuo").Enabled = activar;
            this.oForm.Items.Item("txtOcrCode").Enabled = activar;
            this.oForm.Items.Item("txtAccBank").Enabled = activar;
            this.oForm.Items.Item("txtInsuran").Enabled = activar;
            this.oForm.Items.Item("txtOther").Enabled = activar;
            this.oForm.Items.Item("txtAccIns").Enabled = activar;
            this.oForm.Items.Item("txtAccComm").Enabled = activar;
            this.oForm.Items.Item("txtAccOthe").Enabled = activar;
            this.oForm.Items.Item("txtAccTran").Enabled = activar;
            this.oForm.Items.Item("cbPeriodi").Enabled = activar;
            this.oForm.Items.Item("txtRate").Enabled = activar;
            this.oForm.Items.Item("txtGraMth").Enabled = activar;
            this.oForm.Items.Item("txtPayDay").Enabled = activar;
            this.oForm.Items.Item("chbAccDisb").Enabled = activar;
        }

        protected string consultarMonedaLocal()
        {
            string moneda = "";
            SAPbobsCOM.Recordset oRecordset = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SQL sql = new SQL("ObligacionesFinan.SQL.GetMainCurrency.sql");
            oRecordset.DoQuery(string.Format(sql.getQuery()));

            if (oRecordset.RecordCount > 0)
            {
                moneda = oRecordset.Fields.Item("MainCurncy").Value.ToString();
            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset);
            oRecordset = null;
            GC.Collect();

            return moneda;
        }

        protected bool crearAsiento()
        {

            bool creado = true;

            SAPbobsCOM.JournalEntries oJournal = (SAPbobsCOM.JournalEntries)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);

            oJournal.ReferenceDate = DateTime.ParseExact(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_IniDate", 0), "yyyyMMdd", CultureInfo.InvariantCulture);
            oJournal.Reference = oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("DocNum", 0);
            oJournal.Reference2 = "Asiento de prestamo " + (DATASOURCE.Equals("@HCO_OOFE") ? "Activo" : "Pasivo");

            var desembolsoBanco = oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_AccDisb", 0);
            oJournal.Lines.AccountCode = desembolsoBanco == "Y" ? oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_AccBank", 0) : oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_AccTran", 0);
            oJournal.Lines.CostingCode = oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_OcrCode", 0);

            if (Pais.Equals(Constantes.Colombia))
            {
                oJournal.Lines.UserFields.Fields.Item("U_HBT_Tercero").Value = oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_CardCode", 0);
                oJournal.Lines.UserFields.Fields.Item("U_HBT_ConcepMM").Value = "5004";
            }

            if (oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_Currency", 0).Equals(consultarMonedaLocal()))
            {
                oJournal.Lines.Debit = DATASOURCE.Equals("@HCO_OOFR") ? double.Parse(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_Amount", 0), CultureInfo.InvariantCulture) : 0;
                oJournal.Lines.Credit = DATASOURCE.Equals("@HCO_OOFR") ? 0 : double.Parse(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_Amount", 0), CultureInfo.InvariantCulture);
                oJournal.Lines.Add();

                oJournal.Lines.Credit = DATASOURCE.Equals("@HCO_OOFR") ? double.Parse(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_Amount", 0), CultureInfo.InvariantCulture) : 0;
                oJournal.Lines.Debit = DATASOURCE.Equals("@HCO_OOFR") ? 0 : double.Parse(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_Amount", 0), CultureInfo.InvariantCulture);
            }
            else
            {
                oJournal.Lines.FCCurrency = oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_Currency", 0);
                oJournal.Lines.FCDebit = DATASOURCE.Equals("@HCO_OOFR") ? double.Parse(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_Amount", 0), CultureInfo.InvariantCulture) : 0;
                oJournal.Lines.FCCredit = DATASOURCE.Equals("@HCO_OOFR") ? 0 : double.Parse(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_Amount", 0), CultureInfo.InvariantCulture);
                oJournal.Lines.Add();

                oJournal.Lines.FCCurrency = oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_Currency", 0);
                oJournal.Lines.FCCredit = DATASOURCE.Equals("@HCO_OOFR") ? double.Parse(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_Amount", 0), CultureInfo.InvariantCulture) : 0;
                oJournal.Lines.FCDebit = DATASOURCE.Equals("@HCO_OOFR") ? 0 : double.Parse(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_Amount", 0), CultureInfo.InvariantCulture);
            }

            oJournal.Lines.AccountCode = oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_AcctCode", 0);
            oJournal.Lines.CostingCode = oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_OcrCode", 0);

            if (Pais.Equals(Constantes.Colombia))
            {
                oJournal.Lines.UserFields.Fields.Item("U_HBT_Tercero").Value = oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_CardCode", 0);
                oJournal.Lines.UserFields.Fields.Item("U_HBT_ConcepMM").Value = "5004";
            }

            oJournal.Lines.Add();

            if (oJournal.Add() != 0)
            {
                creado = false;
                SBO_Application.StatusBar.SetText("Error creando asiento: " + oCompany.GetLastErrorDescription(), BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            else
            {
                oForm.DataSources.DBDataSources.Item(DATASOURCE).SetValue("U_HCO_TransId", 0, oCompany.GetNewObjectKey());
            }

            return creado;
        }

        private bool validarCuentas(string cuenta, string item)
        {
            if(!item.Equals("txtAccount"))
            {
                if (cuenta.Equals(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_AcctCode", 0)))
                    return false;
            }
            if (!item.Equals("txtAccBank"))
            {
                if (cuenta.Equals(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_AccBank", 0)))
                    return false;
            }
            if (!item.Equals("txtAccTran"))
            {
                if (cuenta.Equals(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_AccTran", 0)))
                    return false;
            }
            if (!item.Equals("txtAccIns"))
            {
                if (cuenta.Equals(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_AccIns", 0)))
                    return false;
            }
            if (!item.Equals("txtAccComm"))
            {
                if (cuenta.Equals(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_AccComm", 0)))
                    return false;
            }
            if (!item.Equals("txtAccOthe"))
            {
                if (cuenta.Equals(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_AccOthe", 0)))
                    return false;
            }

            return true;
        }

        private void cargarTasaCambio()
        {
            SAPbobsCOM.Recordset oRecordset = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SQL sql = new SQL("ObligacionesFinan.SQL.GetRate.sql");
            DateTime fecha = DateTime.Today;
            if(!string.IsNullOrEmpty(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_IniDate", 0)))
                fecha = DateTime.ParseExact(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_IniDate", 0), "yyyyMMdd", CultureInfo.InvariantCulture);
            oRecordset.DoQuery(string.Format(sql.getQuery(), oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_Currency", 0), fecha.ToString("yyyyMMdd")));

            if (oRecordset.RecordCount > 0)
            {
                oForm.DataSources.DBDataSources.Item(DATASOURCE).SetValue("U_HCO_Rate", 0, oRecordset.Fields.Item("Rate").Value.ToString());
            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset);
            oRecordset = null;
            GC.Collect();
        }

        protected abstract void cargarDocNum();
        protected abstract void cargarCuentaTransitoria();

    }
}