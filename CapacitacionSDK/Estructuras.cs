using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace CapacitacionSDK
{
    public class Estructuras
    {
        private SAPbouiCOM.Application sboAplication;
        private SAPbobsCOM.Company oCompany;

        public Estructuras(SAPbouiCOM.Application application, SAPbobsCOM.Company company)
        {
            this.sboAplication = application;
            this.oCompany = company;
        }

        public void crearEstructuras()
        {
            sboAplication.StatusBar.SetText("Inicio de creación de estructuras.", SAPbouiCOM.BoMessageTime.bmt_Short,
                                SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            CrearUDT();
            CrearUDF();
            CrearUDO();

            sboAplication.StatusBar.SetText("Fin de creación de estructuras.", SAPbouiCOM.BoMessageTime.bmt_Short,
                                SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
        }

        private void CrearUDT()
        {
            List<UserTable> UserTableList = new List<UserTable>();
            XmlDocument xDoc = new XmlDocument();
            xDoc.Load(System.Windows.Forms.Application.StartupPath + @"\Estructuras\UDT.xml");
            XmlSerializer serializer = new XmlSerializer(typeof(List<UserTable>), new XmlRootAttribute("UserTables"));
            string xmlString = xDoc.InnerXml.ToString();
            StringReader stringReader = new StringReader(xmlString);
            UserTableList = (List<UserTable>)serializer.Deserialize(stringReader);

            for (int i = 0; i < UserTableList.Count; i++)
            {
                if (!consultarEstructura(UserTableList[i].TableName, UserTableList[i].Descr, "UDT"))
                {
                    SAPbobsCOM.UserTablesMD oUserTablesMD = (SAPbobsCOM.UserTablesMD)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);
                    oUserTablesMD.TableName = UserTableList[i].TableName.ToString();
                    oUserTablesMD.TableDescription = UserTableList[i].Descr.ToString();
                    oUserTablesMD.TableType = (SAPbobsCOM.BoUTBTableType)UserTableList[i].ObjectType;

                    if (oUserTablesMD.Add() != 0)
                    {
                        sboAplication.StatusBar.SetText("Error al crear la estructura: " + oCompany.GetLastErrorDescription() + ".", SAPbouiCOM.BoMessageTime.bmt_Short,
                                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    }
                    else
                    {
                        sboAplication.StatusBar.SetText("Estructura creada con exito: " + (i + 1).ToString() + " de " + UserTableList.Count + ".",
                            SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                    }

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD);
                    oUserTablesMD = null;
                    GC.Collect();
                }
            }
        }

        private void CrearUDF()
        {
            List<UserField> UserFieldList = new List<UserField>();
            XmlDocument xDoc = new XmlDocument();
            xDoc.Load(System.Windows.Forms.Application.StartupPath + @"\Estructuras\UDF.xml");
            XmlSerializer serializer = new XmlSerializer(typeof(List<UserField>), new XmlRootAttribute("UserFields"));
            string xmlString = xDoc.InnerXml.ToString();
            StringReader stringReader = new StringReader(xmlString);
            UserFieldList = (List<UserField>)serializer.Deserialize(stringReader);

            for (int i = 0; i < UserFieldList.Count; i++)
            {
                if (!consultarEstructura(UserFieldList[i].AliasID.ToString(), UserFieldList[i].TableID.ToString(), "UDF"))
                {
                    SAPbobsCOM.UserFieldsMD oUserFieldsMD = (SAPbobsCOM.UserFieldsMD)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
                    oUserFieldsMD.Name = UserFieldList[i].AliasID.ToString();
                    oUserFieldsMD.Type = (SAPbobsCOM.BoFieldTypes)UserFieldList[i].Tipo;
                    oUserFieldsMD.Description = UserFieldList[i].Descr.ToString();
                    oUserFieldsMD.SubType = (SAPbobsCOM.BoFldSubTypes)UserFieldList[i].SubTipo;
                    oUserFieldsMD.TableName = UserFieldList[i].TableID.ToString();
                    oUserFieldsMD.EditSize = UserFieldList[i].EditSize;
                    oUserFieldsMD.DefaultValue = UserFieldList[i].Dflt.ToString();
                    for (int j = 0; j < UserFieldList[i].ValidValues.Count; j++)
                    {
                        oUserFieldsMD.ValidValues.Value = UserFieldList[i].ValidValues[j].FldValue;
                        oUserFieldsMD.ValidValues.Description = UserFieldList[i].ValidValues[j].Descr;
                        oUserFieldsMD.ValidValues.Add();
                    }

                    if (oUserFieldsMD.Add() != 0)
                    {
                        sboAplication.StatusBar.SetText("Error al crear la estructura: " + oCompany.GetLastErrorDescription() + ".", SAPbouiCOM.BoMessageTime.bmt_Short,
                                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    }
                    else
                    {
                        sboAplication.StatusBar.SetText("Estructura creada con exito: " + (i + 1).ToString() + " de " + UserFieldList.Count + ".",
                            SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);

                        sboAplication.RemoveWindowsMessage(SAPbouiCOM.BoWindowsMessageType.bo_WM_TIMER, true);
                    }

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD);
                    oUserFieldsMD = null;
                    GC.Collect();
                }
            }
        }

        private void CrearUDO()
        {
            List<UDO> UserObjectList = new List<UDO>();
            XmlDocument xDoc = new XmlDocument();
            xDoc.Load(System.Windows.Forms.Application.StartupPath + @"\Estructuras\UDO.xml");
            XmlSerializer serializer = new XmlSerializer(typeof(List<UDO>), new XmlRootAttribute("UDOS"));
            string xmlString = xDoc.InnerXml.ToString();
            StringReader stringReader = new StringReader(xmlString);
            UserObjectList = (List<UDO>)serializer.Deserialize(stringReader);

            for (int i = 0; i < UserObjectList.Count; i++)
            {

                if (!consultarEstructura(UserObjectList[i].Code, UserObjectList[i].Name, "UDO"))
                {
                    SAPbobsCOM.UserObjectsMD oUserObjectsMD = (SAPbobsCOM.UserObjectsMD)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);
                    oUserObjectsMD.Code = UserObjectList[i].Code.ToString();
                    oUserObjectsMD.Name = UserObjectList[i].Name.ToString();
                    oUserObjectsMD.TableName = UserObjectList[i].TableName.ToString();
                    oUserObjectsMD.LogTableName = UserObjectList[i].LogTable.ToString();
                    oUserObjectsMD.ObjectType = (SAPbobsCOM.BoUDOObjType)UserObjectList[i].TYPE;
                    oUserObjectsMD.ExtensionName = UserObjectList[i].ExtName.ToString();
                    oUserObjectsMD.ManageSeries = UserObjectList[i].MngSeries == "Y" ? SAPbobsCOM.BoYesNoEnum.tYES : SAPbobsCOM.BoYesNoEnum.tNO;
                    oUserObjectsMD.CanDelete = UserObjectList[i].CanDelete == "Y" ? SAPbobsCOM.BoYesNoEnum.tYES : SAPbobsCOM.BoYesNoEnum.tNO;
                    oUserObjectsMD.CanCancel = UserObjectList[i].CanCancel == "Y" ? SAPbobsCOM.BoYesNoEnum.tYES : SAPbobsCOM.BoYesNoEnum.tNO;
                    oUserObjectsMD.CanFind = UserObjectList[i].CanFind == "Y" ? SAPbobsCOM.BoYesNoEnum.tYES : SAPbobsCOM.BoYesNoEnum.tNO;
                    oUserObjectsMD.CanYearTransfer = UserObjectList[i].CanYrTrnsf == "Y" ? SAPbobsCOM.BoYesNoEnum.tYES : SAPbobsCOM.BoYesNoEnum.tNO;
                    oUserObjectsMD.CanCreateDefaultForm = UserObjectList[i].CanDefForm == "Y" ? SAPbobsCOM.BoYesNoEnum.tYES : SAPbobsCOM.BoYesNoEnum.tNO;
                    oUserObjectsMD.CanLog = UserObjectList[i].CanLog == "Y" ? SAPbobsCOM.BoYesNoEnum.tYES : SAPbobsCOM.BoYesNoEnum.tNO;
                    oUserObjectsMD.OverwriteDllfile = UserObjectList[i].OvrWrtDll == "Y" ? SAPbobsCOM.BoYesNoEnum.tYES : SAPbobsCOM.BoYesNoEnum.tNO;
                    oUserObjectsMD.EnableEnhancedForm = UserObjectList[i].EnaEnhForm == "Y" ? SAPbobsCOM.BoYesNoEnum.tYES : SAPbobsCOM.BoYesNoEnum.tNO;

                    for (int j = 0; j < UserObjectList[i].ChildTables.Count; j++)
                    {
                        oUserObjectsMD.ChildTables.TableName = UserObjectList[i].ChildTables[j].TableName;
                        oUserObjectsMD.ChildTables.Add();
                    }

                    for (int j = 0; j < UserObjectList[i].FormColumns.Count; j++)
                    {
                        oUserObjectsMD.FormColumns.FormColumnAlias = UserObjectList[i].FormColumns[j].Alias;
                        oUserObjectsMD.FormColumns.FormColumnDescription = UserObjectList[i].FormColumns[j].Description;
                        oUserObjectsMD.FormColumns.Add();
                    }

                    if (oUserObjectsMD.Add() != 0)
                    {
                        sboAplication.StatusBar.SetText("Error al crear la estructura: " + oCompany.GetLastErrorDescription() + ".", SAPbouiCOM.BoMessageTime.bmt_Short,
                                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    }
                    else
                    {
                        sboAplication.StatusBar.SetText("Estructura creada con exito: " + (i + 1).ToString() + " de " + UserObjectList.Count + ".",
                            SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                    }

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectsMD);
                    oUserObjectsMD = null;
                    GC.Collect();
                }
            }
        }

        private bool consultarEstructura(string codigoObjeto, string codigoTabla, string tipo)
        {
            bool creada = false;
            SAPbobsCOM.Recordset oRecordset = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SQL.SQL sql = new SQL.SQL("ObligacionesFinan.SQL.Consultar" + tipo + ".sql");


            oRecordset.DoQuery(string.Format(sql.getQuery(), codigoObjeto, codigoTabla));

            if (oRecordset.RecordCount > 0)
            {
                creada = true;
            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset);
            oRecordset = null;
            GC.Collect();

            return creada;
        }

        public class UserTable
        {
            public string TableName { get; set; }

            public string Descr { get; set; }

            public int TblNum { get; set; }

            public int ObjectType { get; set; }

            public string UsedInObj { get; set; }

            public string LogTable { get; set; }
        }

        public class UserField
        {
            public string TableID { get; set; }
            public int FieldID { get; set; }
            public string AliasID { get; set; }
            public string Descr { get; set; }
            public int Tipo { get; set; }
            public int SubTipo { get; set; }
            public int SizeID { get; set; }
            public int EditSize { get; set; }
            public string Dflt { get; set; }
            public string NotNull { get; set; }
            public string IndexID { get; set; }
            public string RTable { get; set; }
            public string RField { get; set; }
            public string Action { get; set; }
            public string Sys { get; set; }
            public List<ValidValue> ValidValues { get; set; }
        }

        public class ValidValue
        {
            public string FldValue { get; set; }
            public string Descr { get; set; }
        }

        public class UDO
        {
            public string Code { get; set; }

            public string Name { get; set; }

            public string TableName { get; set; }

            public string LogTable { get; set; }

            public int TYPE { get; set; }

            public string MngSeries { get; set; }

            public string CanDelete { get; set; }

            public string CanCancel { get; set; }

            public string ExtName { get; set; }

            public string CanFind { get; set; }

            public string CanYrTrnsf { get; set; }

            public string CanDefForm { get; set; }

            public string CanLog { get; set; }

            public string OvrWrtDll { get; set; }
            public string EnaEnhForm { get; set; }

            public List<ChildTable> ChildTables { get; set; }
            public List<FormColumn> FormColumns { get; set; }
        }

        public class ChildTable
        {
            public int SonNum { get; set; }

            public string TableName { get; set; }

            public string LogName { get; set; }
        }

        public class FormColumn
        {
            public string Alias { get; set; }

            public string Description { get; set; }
        }
    }
}
