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
    class Poliza
    {
        protected Application SBO_Application;
        protected Form oForm;
        protected SAPbobsCOM.Company oCompany;
        protected string DATASOURCE;
        protected string DATASOURCELINE;
        protected string Pais;
        private double importe;

        public Poliza(Application sboaApplication, SAPbobsCOM.Company sboCompany, string pais)
        {
            this.SBO_Application = sboaApplication;
            this.oCompany = sboCompany;            
            this.DATASOURCE = "@HCO_OPOL";
            this.DATASOURCELINE = "@HCO_POL1";
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
                    oXmlDataDocument.Load(System.Windows.Forms.Application.StartupPath + @"/FormulariosXml/Poliza.xml");
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

        protected void inicializar()
        {
            cargarDocNum();

            ChooseFromList objChooseFromList = oForm.ChooseFromLists.Item("cflAcctCode");
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
            condition.CondVal = "C";
            condition.BracketCloseNum = 1;
            objChooseFromList.SetConditions(conditions);

            objChooseFromList = oForm.ChooseFromLists.Item("cflCardName");
            conditions = objChooseFromList.GetConditions();
            condition = conditions.Add();
            condition.BracketOpenNum = 1;
            condition.Alias = "CardType";
            condition.Operation = BoConditionOperation.co_EQUAL;
            condition.CondVal = "C";
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

            activarCampos(true);

            System.Runtime.InteropServices.Marshal.ReleaseComObject(objChooseFromList);
            objChooseFromList = null;
            GC.Collect();
        }

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
                        this.oForm.Items.Item("btnCalculo").Enabled = false;
                    }
                    else if (pVal.MenuUID == "1282")
                    {
                        cargarDocNum();
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
                                Matrix matrixCuotas = (Matrix)oForm.Items.Item("mtCuotas").Specific;

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
                else if (pVal.EventType == BoEventTypes.et_VALIDATE)
                {
                    if (new string[] { "txtAmount", "txtIniDate", "txtMonths", "txtEndDate" }.Contains(pVal.ItemUID))
                    {
                        if (!pVal.InnerEvent && pVal.ItemChanged)
                        {
                            Matrix matrixCuotas = (Matrix)oForm.Items.Item("mtCuotas").Specific;
                            if (oForm.Mode == BoFormMode.fm_ADD_MODE)
                            {
                                matrixCuotas.Clear();
                                matrixCuotas.FlushToDataSource();
                            }
                            else if (oForm.Mode == BoFormMode.fm_UPDATE_MODE)
                            {
                                if (SBO_Application.MessageBox("Se recalcularán las cuotas, de sea continuar?", 1, "Continuar", "Cancelar") == 1)
                                {
                                    int j = 0;
                                    for (int i = 0; i < matrixCuotas.RowCount; i++)
                                    {
                                        if (!string.IsNullOrEmpty(oForm.DataSources.DBDataSources.Item(DATASOURCELINE).GetValue("U_HCO_TransId", i)))
                                            j++;
                                    }
                                    recalcularCuotas(j);
                                }
                            }
                        }
                        if (pVal.ItemUID.Equals("txtIniDate"))
                        {
                            if(!string.IsNullOrEmpty(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_IniDate", 0)) && !string.IsNullOrEmpty(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_EndDate", 0)))
                            {
                                DateTime fechaInicial = DateTime.ParseExact(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_IniDate", 0), "yyyyMMdd", CultureInfo.InvariantCulture);
                                DateTime fechaFinal = DateTime.ParseExact(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_EndDate", 0), "yyyyMMdd", CultureInfo.InvariantCulture);
                                var meses = Math.Abs((fechaFinal.Month - fechaInicial.Month) + 12 * (fechaFinal.Year - fechaInicial.Year));
                                oForm.DataSources.DBDataSources.Item(DATASOURCE).SetValue("U_HCO_Months", 0, meses.ToString(CultureInfo.InvariantCulture));
                            }                            
                        }
                        else if (pVal.ItemUID.Equals("txtEndDate"))
                        {
                            if (!string.IsNullOrEmpty(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_IniDate", 0)) && !string.IsNullOrEmpty(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_EndDate", 0)))
                            {
                                DateTime fechaInicial = DateTime.ParseExact(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_IniDate", 0), "yyyyMMdd", CultureInfo.InvariantCulture);
                                DateTime fechaFinal = DateTime.ParseExact(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_EndDate", 0), "yyyyMMdd", CultureInfo.InvariantCulture);
                                var meses = Math.Abs((fechaFinal.Month - fechaInicial.Month) + 12 * (fechaFinal.Year - fechaInicial.Year));
                                oForm.DataSources.DBDataSources.Item(DATASOURCE).SetValue("U_HCO_Months", 0, meses.ToString(CultureInfo.InvariantCulture));
                            }
                        }
                        else if (pVal.ItemUID.Equals("txtMonths"))
                        {
                            if (!string.IsNullOrEmpty(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_IniDate", 0)) && !string.IsNullOrEmpty(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_Months", 0)))
                            {
                                DateTime fechaInicial = DateTime.ParseExact(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_IniDate", 0), "yyyyMMdd", CultureInfo.InvariantCulture);
                                int meses = int.Parse(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_Months", 0), CultureInfo.InvariantCulture);
                                var fechaFinal = fechaInicial.AddMonths(meses);
                                oForm.DataSources.DBDataSources.Item(DATASOURCE).SetValue("U_HCO_EndDate", 0, fechaFinal.ToString("yyyyMMdd"));
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
                        }
                    }
                    else if (pVal.ItemUID.Equals("txtAccount"))
                    {
                        IChooseFromListEvent chflarg = (IChooseFromListEvent)pVal;
                        if (chflarg.SelectedObjects != null)
                        {
                            DataTable dt = chflarg.SelectedObjects;
                            this.oForm.DataSources.DBDataSources.Item(DATASOURCE).SetValue("U_HCO_AcctCode", 0, Convert.ToString(dt.GetValue("AcctCode", 0)));
                        }
                    }
                    else if (pVal.ItemUID.Equals("txtAccBank"))
                    {
                        IChooseFromListEvent chflarg = (IChooseFromListEvent)pVal;
                        if (chflarg.SelectedObjects != null)
                        {
                            DataTable dt = chflarg.SelectedObjects;
                            this.oForm.DataSources.DBDataSources.Item(DATASOURCE).SetValue("U_HCO_AccBank", 0, Convert.ToString(dt.GetValue("AcctCode", 0)));
                        }
                    }
                    else if (pVal.ItemUID.Equals("txtAccTran"))
                    {
                        IChooseFromListEvent chflarg = (IChooseFromListEvent)pVal;
                        if (chflarg.SelectedObjects != null)
                        {
                            DataTable dt = chflarg.SelectedObjects;
                            this.oForm.DataSources.DBDataSources.Item(DATASOURCE).SetValue("U_HCO_AccTran", 0, Convert.ToString(dt.GetValue("AcctCode", 0)));
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
                    if (pVal.ItemUID.Equals("btnCalculo"))
                    {
                        if (validarCampos())
                        {
                            calcularCoutas();                                
                        }
                    }
                    else if (pVal.ItemUID.Equals("1"))
                    {
                        if (oForm.Mode == BoFormMode.fm_ADD_MODE)
                            cargarDocNum();
                    }
                }
                else if (pVal.EventType == BoEventTypes.et_GOT_FOCUS)
                {                    
                    if (pVal.ItemUID.Equals("txtAmount"))
                    {                        
                        this.importe = double.Parse(this.oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_Amount", 0), CultureInfo.InvariantCulture);
                    }                    
                }
            }
        }

        public void ManejarEventosFormData(ref SAPbouiCOM.BusinessObjectInfo oBusinessInfo, ref bool pBubbleEvent)
        {
            if (oBusinessInfo.ActionSuccess)
            {                
                if (!oBusinessInfo.BeforeAction)
                {
                    this.oForm = SBO_Application.Forms.ActiveForm;
                    if (oBusinessInfo.EventType == BoEventTypes.et_FORM_DATA_LOAD)
                    {
                        //Si todas las cuotas estan pagas se bloquea el campo de monto cuando se cargan los datos.
                        Matrix matrix = (Matrix)oForm.Items.Item("mtCuotas").Specific;
                        int j = 0;
                        for (int i = 0; i < matrix.RowCount; i++)
                        {
                            if (!string.IsNullOrEmpty(oForm.DataSources.DBDataSources.Item(DATASOURCELINE).GetValue("U_HCO_TransId", i)))
                                j++;
                            else
                                break;
                        }
                        if(j == matrix.RowCount)
                        {
                            oForm.Items.Item("txtAmount").Enabled = false;
                        }
                    }
                }
            }
        }

        //Carga el docnum siguiente dependiendo de datasource.
        private void cargarDocNum()
        {
            SAPbobsCOM.Recordset oRecordset = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SQL sql = new SQL("ObligacionesFinan.SQL.GetMaxOPOLocNum.sql");
            oRecordset.DoQuery(string.Format(sql.getQuery()));

            if (oRecordset.RecordCount > 0)
            {
                oForm.DataSources.DBDataSources.Item(DATASOURCE).SetValue("DocNum", 0, oRecordset.Fields.Item("DocNum").Value.ToString());
            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset);
            oRecordset = null;
            GC.Collect();
        }

        //Activa los items dependiendo del parametro de entrada
        protected void activarCampos(bool activar)
        {
            this.oForm.Items.Item("txtCCode").Enabled = activar;
            this.oForm.Items.Item("txtCName").Enabled = activar;
            this.oForm.Items.Item("txtAmount").Enabled = activar;
            this.oForm.Items.Item("txtAccount").Enabled = activar;
            this.oForm.Items.Item("txtIniDate").Enabled = activar;
            this.oForm.Items.Item("txtCurr").Enabled = activar;
            this.oForm.Items.Item("txtEndDate").Enabled = activar;
            this.oForm.Items.Item("txtMonths").Enabled = activar;
            this.oForm.Items.Item("txtAccBank").Enabled = activar;
            this.oForm.Items.Item("txtAccTran").Enabled = activar;
            this.oForm.Items.Item("txtOcrCode").Enabled = activar;
            this.oForm.Items.Item("btnCalculo").Enabled = activar;
        }

        //Valida que los campos requeridos estén diligenciados
        private bool validarCampos()
        {
            bool BubbelEvent = true;
            if (string.IsNullOrEmpty(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_CardCode", 0)))
            {
                SBO_Application.StatusBar.SetText("Debe seleccionar el socio de negocios", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                oForm.Items.Item("txtCCode").Click(BoCellClickType.ct_Regular);
                BubbelEvent = false;
            }
            else if (string.IsNullOrEmpty(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_AcctCode", 0)))
            {
                SBO_Application.StatusBar.SetText("Debe seleccionar la cuenta contable", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                oForm.Items.Item("txtAccount").Click(BoCellClickType.ct_Regular);
                BubbelEvent = false;
            }
            else if (string.IsNullOrEmpty(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_AccBank", 0)))
            {
                SBO_Application.StatusBar.SetText("Debe seleccionar la cuenta del banco", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                oForm.Items.Item("txtAccBank").Click(BoCellClickType.ct_Regular);
                BubbelEvent = false;
            }
            else if (string.IsNullOrEmpty(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_AccTran", 0)))
            {
                SBO_Application.StatusBar.SetText("Debe seleccionar la cuenta transitoria", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                oForm.Items.Item("txtAccTran").Click(BoCellClickType.ct_Regular);
                BubbelEvent = false;
            }
            else if (double.Parse(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_Months", 0), CultureInfo.InvariantCulture) == 0)
            {
                SBO_Application.StatusBar.SetText("Debe ingresar el plazo en meses", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                oForm.Items.Item("txtMonths").Click(BoCellClickType.ct_Regular);
                BubbelEvent = false;
            }
            else if (double.Parse(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_Amount", 0), CultureInfo.InvariantCulture) == 0)
            {
                SBO_Application.StatusBar.SetText("Debe ingresar el monto", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                oForm.Items.Item("txtAmount").Click(BoCellClickType.ct_Regular);
                BubbelEvent = false;
            }
            else if (string.IsNullOrEmpty(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_Currency", 0)))
            {
                SBO_Application.StatusBar.SetText("Debe seleccionar la moneda", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                oForm.Items.Item("txtCurr").Click(BoCellClickType.ct_Regular);
                BubbelEvent = false;
            }
            else if (string.IsNullOrEmpty(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_IniDate", 0)))
            {
                SBO_Application.StatusBar.SetText("Debe ingresar la fecha de inicio", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                oForm.Items.Item("txtIniDate").Click(BoCellClickType.ct_Regular);
                BubbelEvent = false;
            }
            else if (string.IsNullOrEmpty(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_EndDate", 0)))
            {
                SBO_Application.StatusBar.SetText("Debe ingresar la fecha final ", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                oForm.Items.Item("txtEndDate").Click(BoCellClickType.ct_Regular);
                BubbelEvent = false;
            }
            else if (string.IsNullOrEmpty(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_OcrCode", 0)))
            {
                SBO_Application.StatusBar.SetText("Debe seleccionar el centro de costos", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                oForm.Items.Item("txtOcrCode").Click(BoCellClickType.ct_Regular);
                BubbelEvent = false;
            }


            return BubbelEvent;
        }

        private void calcularCoutas()
        {
            Matrix matrixCuotas = (Matrix)this.oForm.Items.Item("mtCuotas").Specific;

            int plazo = int.Parse(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_Months", 0));
            DateTime fechaInicial = DateTime.ParseExact(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_IniDate", 0), "yyyyMMdd", CultureInfo.InvariantCulture);
            double importe = double.Parse(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_Amount", 0), CultureInfo.InvariantCulture);

            double cuota = importe / plazo;

            oForm.DataSources.DBDataSources.Item(DATASOURCE).SetValue("U_HCO_MnthPay", 0, cuota.ToString(CultureInfo.InvariantCulture));

            matrixCuotas.Clear();
            matrixCuotas.AddRow(plazo);
            matrixCuotas.FlushToDataSource();

            double saldoInicial = importe;

            for (int i = 1; i <= plazo; i++)
            {
                oForm.DataSources.DBDataSources.Item(DATASOURCELINE).SetValue("LineId", i - 1, i.ToString());
                oForm.DataSources.DBDataSources.Item(DATASOURCELINE).SetValue("U_HCO_Date", i - 1, fechaInicial.AddMonths(i).ToString("yyyyMMdd"));
                oForm.DataSources.DBDataSources.Item(DATASOURCELINE).SetValue("U_HCO_PayAmt", i - 1, cuota.ToString(CultureInfo.InvariantCulture));
                oForm.DataSources.DBDataSources.Item(DATASOURCELINE).SetValue("U_HCO_FinalAmt", i - 1, (saldoInicial - cuota).ToString(CultureInfo.InvariantCulture));

                saldoInicial = saldoInicial - cuota;
            }

            oForm.DataSources.DBDataSources.Item(DATASOURCELINE).SetValue("U_HCO_FinalAmt", matrixCuotas.RowCount - 1, 0.ToString(CultureInfo.InvariantCulture));

            matrixCuotas.LoadFromDataSourceEx();
        }

        private void recalcularCuotas(int linea)
        {
            Matrix matrixCuotas = (Matrix)this.oForm.Items.Item("mtCuotas").Specific;

            double importe = double.Parse(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_Amount", 0), CultureInfo.InvariantCulture);
            double diferencia = importe - this.importe;

            double nuevaCuota = diferencia / (matrixCuotas.RowCount - linea);

            for (int i = linea; i < matrixCuotas.RowCount; i++)
            {
                if(string.IsNullOrEmpty(oForm.DataSources.DBDataSources.Item(DATASOURCELINE).GetValue("U_HCO_TransId", i)))
                {
                    double cuota = double.Parse(oForm.DataSources.DBDataSources.Item(DATASOURCELINE).GetValue("U_HCO_PayAmt", i), CultureInfo.InvariantCulture);
                    oForm.DataSources.DBDataSources.Item(DATASOURCELINE).SetValue("U_HCO_PayAmt", i, (cuota + nuevaCuota).ToString(CultureInfo.InvariantCulture));
                }
            }

            for (int i = 0; i < matrixCuotas.RowCount; i++)
            {                
                double cuota = double.Parse(oForm.DataSources.DBDataSources.Item(DATASOURCELINE).GetValue("U_HCO_PayAmt", i), CultureInfo.InvariantCulture);
                importe = importe - cuota;
                oForm.DataSources.DBDataSources.Item(DATASOURCELINE).SetValue("U_HCO_FinalAmt", i, importe.ToString(CultureInfo.InvariantCulture));                
            }

            oForm.DataSources.DBDataSources.Item(DATASOURCELINE).SetValue("U_HCO_FinalAmt", matrixCuotas.RowCount - 1, 0.ToString(CultureInfo.InvariantCulture));

            matrixCuotas.LoadFromDataSourceEx();

        }

        protected bool crearAsiento()
        {

            bool creado = true;

            SAPbobsCOM.JournalEntries oJournal = (SAPbobsCOM.JournalEntries)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);

            oJournal.ReferenceDate = DateTime.ParseExact(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_IniDate", 0), "yyyyMMdd", CultureInfo.InvariantCulture);
            oJournal.Reference = oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("DocNum", 0);
            oJournal.Reference2 = "Asiento de poliza ";

            oJournal.Lines.AccountCode = oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_AccTran", 0);
            oJournal.Lines.CostingCode = oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_OcrCode", 0);

            if (Pais.Equals(Constantes.Colombia))
            {
                oJournal.Lines.UserFields.Fields.Item("U_HBT_Tercero").Value = oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_CardCode", 0);
                oJournal.Lines.UserFields.Fields.Item("U_HBT_ConcepMM").Value = "5004";
            }

            if (oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_Currency", 0).Equals(consultarMonedaLocal()))
            {
                oJournal.Lines.Debit = 0;
                oJournal.Lines.Credit = double.Parse(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_Amount", 0), CultureInfo.InvariantCulture);
                oJournal.Lines.Add();

                oJournal.Lines.Credit = 0;
                oJournal.Lines.Debit = double.Parse(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_Amount", 0), CultureInfo.InvariantCulture);
            }
            else
            {
                oJournal.Lines.FCCurrency = oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_Currency", 0);
                oJournal.Lines.FCDebit = 0;
                oJournal.Lines.FCCredit = double.Parse(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_Amount", 0), CultureInfo.InvariantCulture);
                oJournal.Lines.Add();

                oJournal.Lines.FCCurrency = oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_Currency", 0);
                oJournal.Lines.FCCredit = 0;
                oJournal.Lines.FCDebit = double.Parse(oForm.DataSources.DBDataSources.Item(DATASOURCE).GetValue("U_HCO_Amount", 0), CultureInfo.InvariantCulture);
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
    }
}
