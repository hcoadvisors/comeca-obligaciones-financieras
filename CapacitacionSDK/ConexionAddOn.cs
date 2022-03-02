using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ObligacionesFinan.Formularios;
using SAPbobsCOM;
using SAPbouiCOM;
using CompanyClass = SAPbobsCOM.CompanyClass;

namespace CapacitacionSDK
{

    // 0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056

    public class ConexionAddOn
    {

        private SAPbouiCOM.Application SBO_Application;
        private SAPbobsCOM.Company oCompany;
        private PrestamoRecibido PrestamoRecibido;
        private PrestamoEfectuado PrestamoEfectuado;
        private AsisPrestamoRecibido AsisPrestamoRecibido;
        private AsisPrestamoEfectuado asisPrestamoEfectuado;
        private Poliza poliza;
        private AsisPoliza AsisPoliza;
        private string Pais;

        public ConexionAddOn()
        {
            SetApplication();
            AgregarMenus();
        }

        private void SetApplication()
        {

            string strCookie;
            string strConnectionContext;
            int intError;
            string strError = "";

            SAPbouiCOM.SboGuiApi oSboGuiApi=new SboGuiApiClass();
            string strCon = Environment.GetCommandLineArgs().GetValue(1).ToString();

            oSboGuiApi.Connect(strCon);

            SBO_Application = oSboGuiApi.GetApplication();

            SBO_Application.ItemEvent+=new _IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);

            SBO_Application.MenuEvent += new _IApplicationEvents_MenuEventEventHandler(SBO_Application_MenuEvent);

            SBO_Application.AppEvent += new _IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);

            SBO_Application.FormDataEvent += new _IApplicationEvents_FormDataEventEventHandler(SBO_Application_FormDataEvent);

            oCompany = new CompanyClass();

            strCookie = oCompany.GetContextCookie();

            strConnectionContext = SBO_Application.Company.GetConnectionContext(strCookie);

            if (oCompany.Connected)
            {
                oCompany.Disconnect();
            }

            oCompany.SetSboLoginContext(strConnectionContext);

            intError = oCompany.Connect();

            if (intError != 0)
            {
                oCompany.GetLastError(out intError, out strError);

                SBO_Application.StatusBar.SetText(strError,BoMessageTime.bmt_Medium,BoStatusBarMessageType.smt_Error);
            }
            else
            {
                new Estructuras(SBO_Application, oCompany).crearEstructuras();

                SBO_Application.StatusBar.SetText("AddOn conectado correctamente: " + oCompany.CompanyName, BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Success);

                SetFilters();

                consultarPais();
            }
        }

        private void consultarPais()
        {
            Recordset oRecordset = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            SQL.SQL sql = new SQL.SQL("ObligacionesFinan.SQL.GetContrySociety.sql");
            oRecordset.DoQuery(string.Format(sql.getQuery()));

            if (oRecordset.RecordCount > 0)
            {
                this.Pais = oRecordset.Fields.Item("Country").Value.ToString();
            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset);
            oRecordset = null;
            GC.Collect();
        }

        private void SetFilters()
        {

            SAPbouiCOM.EventFilters oFilters=new EventFilters();
            SAPbouiCOM.EventFilter oFilter;

            oFilter = oFilters.Add(BoEventTypes.et_MENU_CLICK);
            oFilter.AddEx("HCO_MOFR");
            oFilter.AddEx("HCO_MOFE");
            oFilter.AddEx("HCO_MAOFE");
            oFilter.AddEx("HCO_MAOFR");
            oFilter.AddEx("HCO_MPOL");
            oFilter.AddEx("HCO_MAPOL");

            oFilter = oFilters.Add(BoEventTypes.et_ALL_EVENTS);
            oFilter.AddEx("HCO_OOFR");
            oFilter.AddEx("HCO_OOFE");
            oFilter.AddEx("HCO_AOFE");
            oFilter.AddEx("HCO_AOFR");
            oFilter.AddEx("HCO_OPOL");
            oFilter.AddEx("HCO_APOL");

            SBO_Application.SetFilter(oFilters);

        }

        private void AgregarMenus()
        {

            SAPbouiCOM.Menus oMenus;
            SAPbouiCOM.MenuItem oMenuItem;
            SAPbouiCOM.MenuCreationParams oMenuCreationParams;

            oMenuCreationParams = (MenuCreationParams) SBO_Application.CreateObject(BoCreatableObjectType.cot_MenuCreationParams);

            oMenuItem = SBO_Application.Menus.Item("1536");

            oMenuCreationParams.Type = BoMenuType.mt_STRING;
            oMenuCreationParams.UniqueID = "HCO_MOFR";
            oMenuCreationParams.String = "Prestamo pasivo";
            oMenuCreationParams.Position = 13;

            oMenus = oMenuItem.SubMenus;

            if (!oMenus.Exists("HCO_MOFR"))
            {
                oMenus.AddEx(oMenuCreationParams);
            }

            oMenuCreationParams = (MenuCreationParams)SBO_Application.CreateObject(BoCreatableObjectType.cot_MenuCreationParams);

            oMenuItem = SBO_Application.Menus.Item("1536");

            oMenuCreationParams.Type = BoMenuType.mt_STRING;
            oMenuCreationParams.UniqueID = "HCO_MOFE";
            oMenuCreationParams.String = "Prestamo activo";
            oMenuCreationParams.Position = 14;

            oMenus = oMenuItem.SubMenus;

            if (!oMenus.Exists("HCO_MOFE"))
            {
                oMenus.AddEx(oMenuCreationParams);
            }

            oMenuCreationParams = (MenuCreationParams)SBO_Application.CreateObject(BoCreatableObjectType.cot_MenuCreationParams);

            oMenuItem = SBO_Application.Menus.Item("1536");

            oMenuCreationParams.Type = BoMenuType.mt_STRING;
            oMenuCreationParams.UniqueID = "HCO_MAOFR";
            oMenuCreationParams.String = "Asistente de obligaciones pasivas";
            oMenuCreationParams.Position = 15;

            oMenus = oMenuItem.SubMenus;

            if (!oMenus.Exists("HCO_MAOFR"))
            {
                oMenus.AddEx(oMenuCreationParams);
            }

            oMenuCreationParams = (MenuCreationParams)SBO_Application.CreateObject(BoCreatableObjectType.cot_MenuCreationParams);

            oMenuItem = SBO_Application.Menus.Item("1536");

            oMenuCreationParams.Type = BoMenuType.mt_STRING;
            oMenuCreationParams.UniqueID = "HCO_MAOFE";
            oMenuCreationParams.String = "Asistente de obligaciones activas";
            oMenuCreationParams.Position = 16;

            oMenus = oMenuItem.SubMenus;

            if (!oMenus.Exists("HCO_MAOFE"))
            {
                oMenus.AddEx(oMenuCreationParams);
            }

            oMenuCreationParams = (MenuCreationParams)SBO_Application.CreateObject(BoCreatableObjectType.cot_MenuCreationParams);

            oMenuItem = SBO_Application.Menus.Item("1536");

            oMenuCreationParams.Type = BoMenuType.mt_STRING;
            oMenuCreationParams.UniqueID = "HCO_MPOL";
            oMenuCreationParams.String = "Poliza / Gasto prepagado";
            oMenuCreationParams.Position = 17;

            oMenus = oMenuItem.SubMenus;

            if (!oMenus.Exists("HCO_MPOL"))
            {
                oMenus.AddEx(oMenuCreationParams);
            }

            oMenuCreationParams = (MenuCreationParams)SBO_Application.CreateObject(BoCreatableObjectType.cot_MenuCreationParams);

            oMenuItem = SBO_Application.Menus.Item("1536");

            oMenuCreationParams.Type = BoMenuType.mt_STRING;
            oMenuCreationParams.UniqueID = "HCO_MAPOL";
            oMenuCreationParams.String = "Asistente de polizas / gastos prepagados";
            oMenuCreationParams.Position = 18;

            oMenus = oMenuItem.SubMenus;

            if (!oMenus.Exists("HCO_MAPOL"))
            {
                oMenus.AddEx(oMenuCreationParams);
            }
        }

        public void SBO_Application_ItemEvent(string strFormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (pVal.FormTypeEx == "HCO_OOFR")
                {
                    this.PrestamoRecibido.ManejarEventosItem(ref pVal, ref BubbleEvent);
                }
                else if (pVal.FormTypeEx == "HCO_OOFE")
                    this.PrestamoEfectuado.ManejarEventosItem(ref pVal, ref BubbleEvent);
                else if (pVal.FormTypeEx == "HCO_AOFR")
                    this.AsisPrestamoRecibido.ManejarEventosItem(ref pVal, ref BubbleEvent);
                else if (pVal.FormTypeEx == "HCO_AOFE")
                    this.asisPrestamoEfectuado.ManejarEventosItem(ref pVal, ref BubbleEvent);
                else if (pVal.FormTypeEx == "HCO_OPOL")
                    this.poliza.ManejarEventosItem(ref pVal, ref BubbleEvent);
                else if (pVal.FormTypeEx == "HCO_APOL")
                    this.AsisPoliza.ManejarEventosItem(ref pVal, ref BubbleEvent);

            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText(ex.Message);
            }            
        }

        public void SBO_Application_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo oBusinessInfo, out  bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (oBusinessInfo.FormTypeEx == "HCO_OOFR")
                    this.PrestamoRecibido.ManejarEventosFormData(ref oBusinessInfo, ref BubbleEvent);
                else if (oBusinessInfo.FormTypeEx == "HCO_OOFE")
                    this.PrestamoEfectuado.ManejarEventosFormData(ref oBusinessInfo, ref BubbleEvent);
                else if (oBusinessInfo.FormTypeEx == "HCO_OPOL")
                    this.poliza.ManejarEventosFormData(ref oBusinessInfo, ref BubbleEvent);
            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText(ex.Message);
            }
        }

        public void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            if (pVal.MenuUID == "HCO_MOFR")
            {
                if (pVal.BeforeAction)
                {
                    this.PrestamoRecibido = new PrestamoRecibido(SBO_Application, oCompany, Pais);
                    this.PrestamoRecibido.CrearFormulario();
                }
            }
            else if (pVal.MenuUID == "HCO_MOFE")
            {
                if (pVal.BeforeAction)
                {
                    this.PrestamoEfectuado = new PrestamoEfectuado(SBO_Application, oCompany, Pais);
                    this.PrestamoEfectuado.CrearFormulario();
                }
            }
            else if (pVal.MenuUID == "HCO_MAOFR")
            {
                if (pVal.BeforeAction)
                {
                    this.AsisPrestamoRecibido = new AsisPrestamoRecibido(SBO_Application, oCompany, Pais);
                    this.AsisPrestamoRecibido.CrearFormulario();
                }
            }
            else if (pVal.MenuUID == "HCO_MAOFE")
            {
                if (pVal.BeforeAction)
                {
                    this.asisPrestamoEfectuado = new AsisPrestamoEfectuado(SBO_Application, oCompany, Pais);
                    this.asisPrestamoEfectuado.CrearFormulario();
                }
            }
            else if (pVal.MenuUID == "HCO_MPOL")
            {
                if (pVal.BeforeAction)
                {
                    this.poliza = new Poliza(SBO_Application, oCompany, Pais);
                    this.poliza.CrearFormulario();
                }
            }
            else if (pVal.MenuUID == "HCO_MAPOL")
            {
                if (pVal.BeforeAction)
                {
                    this.AsisPoliza = new AsisPoliza(SBO_Application, oCompany, Pais);
                    this.AsisPoliza.CrearFormulario();
                }
            }
            try
            {
                var form = SBO_Application.Forms.ActiveForm;
                if (form.TypeEx == "HCO_OOFR")
                {
                    this.PrestamoRecibido.ManejarEventosMenus(ref pVal, ref BubbleEvent);
                }
                else if (form.TypeEx == "HCO_OOFE")
                    this.PrestamoEfectuado.ManejarEventosMenus(ref pVal, ref BubbleEvent);
                else if (form.TypeEx == "HCO_OPOL")
                    this.poliza.ManejarEventosMenus(ref pVal, ref BubbleEvent);
            }
            catch (Exception)
            {                
            }
        }

        public void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes eventTypes)
        {

            try
            {

                SAPbouiCOM.Menus oMenus;

                if (eventTypes == BoAppEventTypes.aet_CompanyChanged || eventTypes == BoAppEventTypes.aet_FontChanged ||
                    eventTypes == BoAppEventTypes.aet_LanguageChanged || eventTypes == BoAppEventTypes.aet_ServerTerminition ||
                    eventTypes == BoAppEventTypes.aet_ShutDown)
                {

                    if (SBO_Application.Forms.Count > 0)
                    {
                        foreach (SAPbouiCOM.Form oForm in SBO_Application.Forms)
                        {
                            if (new string[] { "HCO_OOFR", "HCO_OOFE", "HCO_AOFR", "HCO_AOFE", "HCO_OPOL", "HCO_APOL" }.Contains(oForm.TypeEx))
                            {
                                oForm.Close();
                            }
                        }
                    }
                  
                    oMenus = SBO_Application.Menus;

                    if (oMenus.Exists("HCO_MOFR"))
                        oMenus.RemoveEx("HCO_MOFR");

                    if (oMenus.Exists("HCO_MOFE"))
                        oMenus.RemoveEx("HCO_MOFE");

                    if (oMenus.Exists("HCO_MAOFR"))
                        oMenus.RemoveEx("HCO_MAOFR");

                    if (oMenus.Exists("HCO_MAOFE"))
                        oMenus.RemoveEx("HCO_MAOFE");

                    if (oMenus.Exists("HCO_MPOL"))
                        oMenus.RemoveEx("HCO_MPOL");

                    if (oMenus.Exists("HCO_MAPOL"))
                        oMenus.RemoveEx("HCO_MAPOL");

                    if (oCompany.Connected)
                    {
                        oCompany.Disconnect();
                    }

                    System.Windows.Forms.Application.Exit();

                }

            }
            catch (Exception e)
            {
                SBO_Application.StatusBar.SetText(e.Message, BoMessageTime.bmt_Medium, BoStatusBarMessageType.smt_Error);
            }

        }

    }
}
