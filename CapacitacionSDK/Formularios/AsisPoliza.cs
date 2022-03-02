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
    class AsisPoliza
    {
        protected Application SBO_Application;
        protected Form oForm;
        protected SAPbobsCOM.Company oCompany;
        protected string DATASOURCE;
        protected string Pais;

        public AsisPoliza(Application sboaApplication, SAPbobsCOM.Company sboCompany, string pais)
        {
            this.SBO_Application = sboaApplication;
            this.oCompany = sboCompany;
            this.DATASOURCE = "HCO_OPOL";
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
                    oXmlDataDocument.Load(System.Windows.Forms.Application.StartupPath + @"/FormulariosXml/AsisPoliza.xml");
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

        public void ManejarEventosItem(ref ItemEvent pVal, ref bool BubbelEvent)
        {
            if (pVal.BeforeAction)
            {
                this.oForm = SBO_Application.Forms.Item(pVal.FormUID);
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
            SQL sql = new SQL("ObligacionesFinan.SQL.GetPayPOL.sql");
            oForm.DataSources.DataTables.Item("dtCuotas").ExecuteQuery(string.Format(sql.getQuery(), oForm.DataSources.UserDataSources.Item("udCCodeD").Value, oForm.DataSources.UserDataSources.Item("udCCodeH").Value));

            var grCuotas = (Grid)oForm.Items.Item("grCuotas").Specific;

            grCuotas.Columns.Item("DocNum").Editable = false;
            grCuotas.Columns.Item("DocEntry").Editable = false;
            grCuotas.Columns.Item("U_HCO_CardCode").Editable = false;
            grCuotas.Columns.Item("U_HCO_PayAmt").Editable = false;
            grCuotas.Columns.Item("LineId").Editable = false;
            grCuotas.Columns.Item("LineId").Visible = false;
            grCuotas.Columns.Item("U_HCO_AcctCode").Editable = false;
            grCuotas.Columns.Item("U_HCO_AcctCode").Visible = false;
            grCuotas.Columns.Item("U_HCO_Currency").Editable = false;
            grCuotas.Columns.Item("U_HCO_Currency").Visible = false;
            grCuotas.Columns.Item("MainCurncy").Editable = false;
            grCuotas.Columns.Item("MainCurncy").Visible = false;
            grCuotas.Columns.Item("#").Editable = false;
            grCuotas.Columns.Item("#").Visible = false;
            grCuotas.Columns.Item("U_HCO_OcrCode").Editable = false;
            grCuotas.Columns.Item("U_HCO_OcrCode").Visible = false;
            grCuotas.Columns.Item("U_HCO_AccBank").Editable = false;
            grCuotas.Columns.Item("U_HCO_AccBank").Visible = false;

            ((EditTextColumn)grCuotas.Columns.Item("DocEntry")).LinkedObjectType = DATASOURCE;
            ((EditTextColumn)grCuotas.Columns.Item("U_HCO_CardCode")).LinkedObjectType = "2";

            ((EditTextColumn)grCuotas.Columns.Item("Check")).TitleObject.Caption = "Seleccionar";
            ((EditTextColumn)grCuotas.Columns.Item("DocNum")).TitleObject.Caption = "Numero obliación";
            ((EditTextColumn)grCuotas.Columns.Item("DocEntry")).TitleObject.Caption = "Numero int. obligación";
            ((EditTextColumn)grCuotas.Columns.Item("U_HCO_CardCode")).TitleObject.Caption = "Código SN";
            ((EditTextColumn)grCuotas.Columns.Item("U_HCO_Date")).TitleObject.Caption = "Fecha contabilización";
            ((EditTextColumn)grCuotas.Columns.Item("U_HCO_PayAmt")).TitleObject.Caption = "Valor cuota";

            grCuotas.Columns.Item("Check").Type = BoGridColumnType.gct_CheckBox;
        }

        private void contabilizarCuotas()
        {
            var dtCuotas = oForm.DataSources.DataTables.Item("dtCuotas");

            var dtResult = oForm.DataSources.DataTables.Item("dtResult");
            dtResult.Clear();
            dtResult.Columns.Add("DocNum", BoFieldsType.ft_AlphaNumeric);
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
                        SBO_Application.StatusBar.SetText("Procesando cuota " + (i + 1) + " de " + dtCuotas.Rows.Count, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_None);
                        int transId = 0;

                        if (!oCompany.InTransaction)
                            oCompany.StartTransaction();

                        var resultado = crearAsiento(i);
                        if (resultado.Item1)
                        {
                            dtResult.SetValue("TransId", j, resultado.Item2);
                            transId = int.Parse(resultado.Item2);
                        }
                        else
                        {
                            dtResult.SetValue("Descripcion", j, "Error creando asiento - " + resultado.Item2);
                            if (oCompany.InTransaction)
                                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                            continue;
                        }                        

                        actualizarUDO(i, transId);

                        dtResult.SetValue("Descripcion", j, "Cuota contabilizada correctamente");

                        if (oCompany.InTransaction)
                            oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                    }
                    catch (Exception ex)
                    {
                        if (oCompany.InTransaction)
                            oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                        dtResult.SetValue("Descripcion", j, ex.Message);
                    }
                    j++;
                }
            }

            var grResult = (Grid)oForm.Items.Item("grResult").Specific;

            grResult.Columns.Item("TransId").Editable = false;
            grResult.Columns.Item("Descripcion").Editable = false;
            grResult.Columns.Item("DocNum").Editable = false;

            ((EditTextColumn)grResult.Columns.Item("TransId")).LinkedObjectType = "30";

            ((EditTextColumn)grResult.Columns.Item("DocNum")).TitleObject.Caption = "Poliza";
            ((EditTextColumn)grResult.Columns.Item("TransId")).TitleObject.Caption = "Asiento";
            ((EditTextColumn)grResult.Columns.Item("Descripcion")).TitleObject.Caption = "Resultado";

            grResult.AutoResizeColumns();
        }

        private (bool, string) crearAsiento(int i)
        {
            var dtCuotas = oForm.DataSources.DataTables.Item("dtCuotas");

            SAPbobsCOM.JournalEntries oJournal = (SAPbobsCOM.JournalEntries)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);

            oJournal.ReferenceDate = (DateTime)dtCuotas.GetValue("U_HCO_Date", i);
            oJournal.Reference = dtCuotas.GetValue("DocNum", i).ToString();
            oJournal.Reference2 = dtCuotas.GetValue("LineId", i).ToString();
            oJournal.Reference3 = "Cuota de Poliza";

            oJournal.Lines.AccountCode = dtCuotas.GetValue("U_HCO_AcctCode", i).ToString();
            oJournal.Lines.CostingCode = dtCuotas.GetValue("U_HCO_OcrCode", i).ToString();

            if (Pais.Equals(Constantes.Colombia))
            {
                oJournal.Lines.UserFields.Fields.Item("U_HBT_Tercero").Value = dtCuotas.GetValue("U_HCO_CardCode", i).ToString();
                oJournal.Lines.UserFields.Fields.Item("U_HBT_ConcepMM").Value = "5004";
            }

            if (dtCuotas.GetValue("U_HCO_Currency", i).ToString().Equals(dtCuotas.GetValue("MainCurncy", i).ToString()))
            {
                oJournal.Lines.Credit = (double)dtCuotas.GetValue("U_HCO_PayAmt", i);
                oJournal.Lines.Debit = 0;
                oJournal.Lines.Add();

                oJournal.Lines.Debit = (double)dtCuotas.GetValue("U_HCO_PayAmt", i);
                oJournal.Lines.Credit = 0;
            }
            else
            {
                oJournal.Lines.FCCurrency = dtCuotas.GetValue("U_HCO_Currency", i).ToString();
                oJournal.Lines.FCCredit = (double)dtCuotas.GetValue("U_HCO_PayAmt", i);
                oJournal.Lines.FCDebit = 0;
                oJournal.Lines.Add();

                oJournal.Lines.FCCurrency = dtCuotas.GetValue("U_HCO_Currency", i).ToString();
                oJournal.Lines.FCDebit = (double)dtCuotas.GetValue("U_HCO_PayAmt", i);
                oJournal.Lines.FCCredit = 0;

            }            

            oJournal.Lines.AccountCode = dtCuotas.GetValue("U_HCO_AccBank", i).ToString();
            oJournal.Lines.CostingCode = dtCuotas.GetValue("U_HCO_OcrCode", i).ToString();

            if (Pais.Equals(Constantes.Colombia))
            {
                oJournal.Lines.UserFields.Fields.Item("U_HBT_Tercero").Value = dtCuotas.GetValue("U_HCO_CardCode", i).ToString();
                oJournal.Lines.UserFields.Fields.Item("U_HBT_ConcepMM").Value = "5004";
            }
            oJournal.Lines.Add();

            if (oJournal.Add() != 0)
            {
                return (false, oCompany.GetLastErrorDescription());
            }
            else
            {
                return (true, oCompany.GetNewObjectKey());
            }
        }

        private void actualizarUDO(int i, int transId)
        {
            var dtCuotas = oForm.DataSources.DataTables.Item("dtCuotas");

            UDOPoliza UDOpoliza = new UDOPoliza("HCO_OPOL", "HCO_POL1");
            UDOpoliza.DocEntry = (int)dtCuotas.GetValue("DocEntry", i);
            UDOpoliza.PolizaLineas = new List<PolizaLinea>();

            PolizaLinea polizaLinea = new PolizaLinea();
            polizaLinea.LineId = (int)dtCuotas.GetValue("LineId", i);
            polizaLinea.TransId = transId;

            UDOpoliza.PolizaLineas.Add(polizaLinea);

            UDOpoliza.actualizarUDO(ref oCompany);
        }
    }
}
