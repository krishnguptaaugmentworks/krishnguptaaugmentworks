using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Collections;
using System.Data.SqlClient;
using System.Xml;
using System.Globalization;
using System.Data;
using System.Net;
using System.Diagnostics;
using System.Threading;
using System.Security.Cryptography;
using System.Configuration;

namespace AgingReport
{
    public partial class AddOnInfo
    {
        SAPbouiCOM.SboGuiApi __GUI = null;
        SAPbobsCOM.Company ____bobCompany = null;
        SAPbobsCOM.Company MainCompany = new SAPbobsCOM.Company();
        SAPbouiCOM.Application __app = null;
        SAPbouiCOM.Form __Form, __Form1 = null;
        public DataTable dTable = new DataTable();
        public DataTable dTable1 = new DataTable();
        public DataTable Htable = new DataTable();
        public DataRow Hrow; 
        SAPbobsCOM.Recordset oRs = null;
        SAPbobsCOM.Recordset oRs1, oRs2, oRs3, oRs4 = null;
        SAPbouiCOM.Grid oGrid = null;
        SAPbouiCOM.Matrix oMatrix = null;
        SAPbouiCOM.EditText oEdit = null;
        SAPbouiCOM.EditText oEdit1, oEdit2, oEdit3 = null;
        SAPbouiCOM.ComboBox oComboBox = null;
        public SAPbouiCOM.DataTable DT, DT1;
        public SAPbouiCOM.ProgressBar oProgBar;
        DataSet dSet = new DataSet();
        DataSet dSet1 = new DataSet();
        bool Call = false;
        bool Manual = false;
        public bool Status = false;
        public bool IsValidate = true;
        public int FormCode = 0;
        public string MainEnq = "";       

        public AddOnInfo()
        {

            //if (addrow == null)
            //    addrow = new ArrayList();
        }

        public void StartAddOn()
        {
            int result = Connect();
            if (result == -1)
            {
                Environment.Exit(0);
            }
        }

        public int Connect()
        {
            try
            {
                __GUI = new SAPbouiCOM.SboGuiApi();
                __GUI.Connect(Environment.GetCommandLineArgs().GetValue(1).ToString());
                __app = __GUI.GetApplication(-1);

                __app.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(__app_MenuEvent);
                __app.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(__app_ItemEvent);
                __app.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(__app_AppEvent);
                __app.FormDataEvent += new SAPbouiCOM._IApplicationEvents_FormDataEventEventHandler(__app_FormDataEvent);

                //return 0;
                int i = 0;


                //Check For License
                //string LicenseKey = ConfigurationManager.AppSettings["LicenseDate"].ToString();

                //DateTime LicenseDate = DateTime.Parse(LicenseKey);



                //if (DateTime.Now.Date <= LicenseDate)
                //{
                ____bobCompany = (SAPbobsCOM.Company)__app.Company.GetDICompany();

                TableCreation();

                //int RetVal = ____bobCompany.Connect();
                //Int32 ErrorCode = 0;
                //string ErrorMessage = "";

                //if (RetVal == 0)
                //{
                XmlDocument oXMLDoc = new XmlDocument();
                string MenuPath = Application.StartupPath + "\\FileGen_Menu.xml";

                if (__app.Menus.Exists("MDC_MNU"))
                {
                    __app.Menus.RemoveEx("MDC_MNU");
                }

                oXMLDoc.Load(MenuPath);
                __app.LoadBatchActions(oXMLDoc.InnerXml);

                i = 0;              



                return i;


            }
            catch (COMException e)
            {
                if (e.ErrorCode == -7201)
                    return e.ErrorCode;
                else
                {
                    __app.MessageBox(e.Message, 0, "Ok", null, null);
                    return 0;
                }
            }
        }

        private void TableCreation()
        {
            try
            {
                ____bobCompany.StartTransaction();

                AddTables("AW_B1CZHDR", "B1 Customize Header", SAPbobsCOM.BoUTBTableType.bott_MasterData);
                AddTables("AW_B1CZDTL", "B1 Customize Rows", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines);
                AddFields("@AW_B1CZHDR", "DocEntry", "DocEntry", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 11, null, null, false, null);

                AddFields("@AW_B1CZDTL", "DocEntry", "DocEntry", SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, 11, null, null, false, null);
                AddFields("@AW_B1CZDTL", "AsColNotes", "Collection Notes", SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, 254, null, null, false, null);
                AddFields("@AW_B1CZDTL", "UpdateDate", "Update Date", SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, 12, null, null, false, null);

                if (-1 == RegisterObject("UDOB1CZHDR", "B1 Customize", "MD", "AW_B1CZHDR", "AW_B1CZDTL", false, false, false, true, true, false, false, "DocEntry", false, "", true))
                {
                    __app.StatusBar.SetText("Object Registration failed.....", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error); return;
                }

                if (____bobCompany.InTransaction == true)
                    ____bobCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
            }
            catch (Exception ex)
            {
                __app.SetStatusBarMessage(ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        public void AddTables(string strTab, string strDesc, SAPbobsCOM.BoUTBTableType nType)
        {
            GC.Collect();
            SAPbobsCOM.UserTablesMD oUserTablesMD = null;
            try
            {
                oUserTablesMD = null;
                oUserTablesMD = ____bobCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);
                // Adding Table
                if (!oUserTablesMD.GetByKey(strTab))
                {
                    oUserTablesMD.TableName = strTab;
                    oUserTablesMD.TableDescription = strDesc;
                    oUserTablesMD.TableType = nType;
                    if (oUserTablesMD.Add() != 0)
                        throw new Exception(____bobCompany.GetLastErrorDescription());
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD);
                oUserTablesMD = null;
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
        }

        public void AddFields(string tablename, string fieldName, string Description, SAPbobsCOM.BoFieldTypes datatype, SAPbobsCOM.BoFldSubTypes subdatatype, int size, ArrayList validvalue, string defaultValue, bool SetLinkTable, string LinkTable)
        {
            SAPbobsCOM.UserFieldsMD oUserFieldsMD = null;
            string errDesc = "";
            GC.Collect();
            bool ISUSERTABLE = true;
            bool ISFieldExist = true;

            GC.Collect();
            try
            {
                string sqlStr1 = "";
                sqlStr1 = "select * from \"OUTB\" Where \"TableName\" = '" + tablename + "'";
                SAPbobsCOM.Recordset oRset1 = null;
                oRset1 = ____bobCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRset1.DoQuery(sqlStr1);
                if (oRset1.RecordCount > 0)
                    ISUSERTABLE = true;
                else
                    ISUSERTABLE = false;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRset1);
                oRset1 = null;
                GC.Collect();

                string sqlStr = "";
                SAPbobsCOM.Recordset oRset = null;
                oRset = ____bobCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                if (ISUSERTABLE == true)
                    sqlStr = "select  *  from CUFD Where \"TableID\" = '@" + tablename + "' and \"AliasID\"= '" + fieldName + "'";
                else
                    sqlStr = "select  *  from CUFD Where \"TableID\" = '" + tablename + "' and \"AliasID\"= '" + fieldName + "'";
                oRset.DoQuery(sqlStr);
                if (oRset.RecordCount > 0)
                    ISFieldExist = true;
                else
                    ISFieldExist = false;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRset);
                oRset = null;
                GC.Collect();

                if (ISFieldExist == false)
                {
                    oUserFieldsMD = null;
                    oUserFieldsMD = ____bobCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
                    oUserFieldsMD.TableName = tablename;
                    oUserFieldsMD.Name = fieldName;
                    oUserFieldsMD.Description = Description;
                    oUserFieldsMD.Type = datatype;
                    oUserFieldsMD.SubType = subdatatype;
                    if (subdatatype.ToString() != "83")
                    {
                        oUserFieldsMD.Size = size;
                        oUserFieldsMD.EditSize = size;
                    }

                    if (SetLinkTable == true)
                    {
                        if (LinkTable != "")
                            oUserFieldsMD.LinkedTable = LinkTable;
                    }
                    else if (validvalue != null)
                    {
                        for (int iVal = 0; iVal <= validvalue.Count - 1; iVal++)
                        {
                            //oUserFieldsMD.ValidValues.Value = validvalue[iVal].ToString();
                            //oUserFieldsMD.ValidValues.Description = validvalue.Item(iVal).Description;
                            //oUserFieldsMD.ValidValues.Add();
                        }
                        if (defaultValue != "")
                        {
                            oUserFieldsMD.DefaultValue = defaultValue;
                            oUserFieldsMD.Mandatory = SAPbobsCOM.BoYesNoEnum.tYES;
                        }
                    }

                    int UDFCode = oUserFieldsMD.Add();
                    if (UDFCode == 0)
                    {
                        errDesc = "The Field " + Description + " is Added in table " + tablename + ".";
                        __app.SetStatusBarMessage(errDesc, SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                    }
                    if (UDFCode != 0)
                    {
                        errDesc = ____bobCompany.GetLastErrorDescription();
                        __app.SetStatusBarMessage(errDesc, SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                    }
                    Marshal.ReleaseComObject(oUserFieldsMD);
                    oUserFieldsMD = null;
                    GC.Collect();
                }
                else
                {
                }
            }
            catch (Exception ex)
            {
                errDesc = ex.Message;
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        public int RegisterObject(string Code, string Name, string ObjectType, string TableName, string ChildTableName, bool CanCancel, bool CanClose, bool CanCreateDefaultForm, bool CanDelete, bool CanFind, bool CanYearTransfer, bool ManageSeries, string FindColumns, bool defaultform, string defaultformFields, bool CanLog)
        {
            try
            {
                int errCode;
                string errMsg = "";

                if (!____bobCompany.Connected)
                    ____bobCompany.Connect();
                SAPbobsCOM.UserObjectsMD UDO = ____bobCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);

                if (UDO.GetByKey(Code) == false)
                {

                    // UDO.CanCancel = GetYesNo(CanCancel)
                    // UDO.CanClose = GetYesNo(CanClose)
                    // UDO.CanDelete = GetYesNo(CanDelete)
                    UDO.CanCancel = GetYesNo(false);
                    UDO.CanClose = GetYesNo(false);
                    UDO.CanDelete = GetYesNo(CanDelete);
                    UDO.CanFind = GetYesNo(CanFind);
                    UDO.CanCreateDefaultForm = GetYesNo(defaultform);
                    UDO.ManageSeries = GetYesNo(ManageSeries);
                    UDO.CanLog = GetYesNo(CanLog);

                    if (UDO.CanCreateDefaultForm == SAPbobsCOM.BoYesNoEnum.tYES)
                    {
                        string[] s = defaultformFields.Split(',');
                        foreach (string o in s)
                        {
                            string[] s1 = o.Split('|');
                            if (s1.Length > 0)
                                UDO.FormColumns.FormColumnAlias = s1[0].ToString();
                            if (s1.Length > 1)
                                UDO.FormColumns.FormColumnDescription = s1[1].ToString();
                            UDO.FormColumns.Add();
                        }
                    }

                    if (UDO.CanFind == SAPbobsCOM.BoYesNoEnum.tYES)
                    {
                        if (FindColumns != "")
                        {
                            string[] s = FindColumns.Split(',');
                            // If s.Length > 1 Then
                            if (s.Length > 0)
                            {
                                foreach (string o in s)
                                {
                                    string s1 = "";
                                    if (o.ToString() == "Code" | o.ToString() == "Name" | o.ToString() == "DocNum" | o.ToString() == "DocEntry")
                                        s1 = o.ToString();
                                    else
                                        s1 = "U_" + o.ToString();
                                    UDO.FindColumns.ColumnAlias = s1;
                                    UDO.FindColumns.Add();
                                }
                            }
                        }
                    }

                    if (UDO.CanLog == SAPbobsCOM.BoYesNoEnum.tYES)
                        UDO.LogTableName = "A" + TableName;
                    UDO.Code = Code;
                    UDO.Name = Name;
                    UDO.ObjectType = GetUDOType(ObjectType);
                    UDO.TableName = TableName;

                    if (ChildTableName != "")
                    {
                        string[] childTables = ChildTableName.Split(',');
                        if (childTables.Length > 0)
                        {
                            foreach (string o in childTables)
                            {
                                UDO.ChildTables.TableName = o.ToString();
                                UDO.ChildTables.Add();
                            }
                        }
                    }
                    int i = -1;
                    i = UDO.Add();

                    if (i != 0)
                    {
                        ____bobCompany.GetLastError(out errCode, out errMsg);
                        __app.MessageBox(errMsg);
                        __app.StatusBar.SetText(errMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        UDO.Update();
                        if (Marshal.IsComObject(UDO))
                        {
                            Marshal.ReleaseComObject(UDO);
                            UDO = null/* TODO Change to default(_) if this is not a reference type */;
                            GC.Collect();
                            return errCode;
                        }
                    }
                    UDO.Update();
                    if (Marshal.IsComObject(UDO))
                        Marshal.ReleaseComObject(UDO);
                    UDO = null/* TODO Change to default(_) if this is not a reference type */;
                    // GC.WaitForPendingFinalizers();
                    GC.Collect();
                    // oProgressBar.Value = oProgressBar.Value + 1
                    return 0;
                }
                else
                    __app.StatusBar.SetText("Table Already Exists......", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None);
                return 0;
            }
            catch (Exception ex)
            {
                GC.Collect();
                return -1;
            }
        }

        private SAPbobsCOM.BoUDOObjType GetUDOType(string Type)
        {
            if (Type == "MD")
                return SAPbobsCOM.BoUDOObjType.boud_MasterData;
            else
                return SAPbobsCOM.BoUDOObjType.boud_Document;
        }

        private SAPbobsCOM.BoYesNoEnum GetYesNo(bool s)
        {
            if (s == true)
                return SAPbobsCOM.BoYesNoEnum.tYES;
            else
                return SAPbobsCOM.BoYesNoEnum.tNO;
        }

        internal string GenerateCode(string Field, string TableName)
        {
            string Code = "";
            string SqlQuery = "select (case when (IFNULL(max(convert(numeric," + Field + ") ),0))=0 then 1 else (max(convert(numeric," + Field + ")) + 1) end) as CODE from [dbo].[" + TableName + "]";
            //SAPbobsCOM.Company bCompany = __app.Company.GetDICompany() as SAPbobsCOM.Company;
            SAPbobsCOM.Recordset oRs1 = (SAPbobsCOM.Recordset)____bobCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRs1.DoQuery(SqlQuery);
            Code = oRs1.Fields.Item(0).Value.ToString();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRs1);
            oRs1 = null;
            return Code;
        }

        internal void ActivateForm(SAPbouiCOM.Form f, string defaultBrowser)
        {
            f.Freeze(true);
            f.EnableMenu("1288", true);
            f.EnableMenu("1289", true);
            f.EnableMenu("1290", true);
            f.EnableMenu("1291", true);
            f.EnableMenu("1292", false);
            f.EnableMenu("1293", false);
            f.EnableMenu("6913", false);
            f.EnableMenu("1283", false);
            f.Freeze(false);
            f.Update();
            //f.DefButton = "1";
            f.DataBrowser.BrowseBy = defaultBrowser;
            f.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
        }

        internal string GetData(string SqlQuery, string FieldName)
        {
            SAPbobsCOM.Recordset oRs = (SAPbobsCOM.Recordset)____bobCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRs.DoQuery(SqlQuery);
            string val = "";
            if (oRs.RecordCount > 0)
                val = oRs.Fields.Item(FieldName).Value.ToString();
            else
                val = "";
            Marshal.ReleaseComObject(oRs);
            GC.Collect();
            return val;
        }

        internal bool GetData(string SqlQuery)
        {
            SAPbobsCOM.Recordset oRs = (SAPbobsCOM.Recordset)____bobCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRs.DoQuery(SqlQuery);
            bool val = false;
            if (oRs.RecordCount > 0)
                val = true;
            else
                val = false;
            Marshal.ReleaseComObject(oRs);
            GC.Collect();
            return val;
        }

        #region "User Authorization Method"
        public void CreateAuthorization(string FileName, string Suffix)
        {
            try
            {
                XmlDocument Table = new XmlDocument();
                XmlTextReader xmlReader = new XmlTextReader(FileName);

                string fatherCardID = "";
                while (xmlReader.Read())
                {
                    if (xmlReader.NodeType == XmlNodeType.Element)
                    {
                        if (xmlReader.Name == "SubMenu")
                        {
                            SAPbobsCOM.UserPermissionTree permissionTree = (SAPbobsCOM.UserPermissionTree)____bobCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserPermissionTree);
                            try
                            {
                                if (!permissionTree.GetByKey(xmlReader["ID"].ToString()))
                                {
                                    permissionTree.PermissionID = xmlReader["ID"].ToString();
                                    permissionTree.Name = xmlReader["Text"].ToString() + " - " + Suffix;
                                    permissionTree.Options = SAPbobsCOM.BoUPTOptions.bou_FullNone;
                                    //permissionTree.Levels = 1;
                                    fatherCardID = xmlReader["ID"].ToString();
                                    int Result = permissionTree.Add();
                                    if (Result != 0)
                                        //SAPWrapper.Common.Global.WriteLog(__Company.GetLastErrorDescription(), System.Diagnostics.EventLogEntryType.Error);
                                        ____bobCompany.GetLastErrorDescription();
                                }
                                //break;
                            }
                            catch (COMException e1)
                            {
                                __app.MessageBox(e1.Message, 1, "Ok", "", "");
                                //if (OnError != null)
                                //    OnError(e1, this.GetType(), e1.Message);
                                //SAPWrapper.Common.Global.WriteLog(e1.Message, System.Diagnostics.EventLogEntryType.Error);
                                //break;
                            }
                            finally
                            {
                                if (Marshal.IsComObject(permissionTree))
                                    Marshal.ReleaseComObject(permissionTree);
                            }
                        }
                        else if (xmlReader.Name == "PopUp")
                        {
                            SAPbobsCOM.UserPermissionTree permissionTree = (SAPbobsCOM.UserPermissionTree)____bobCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserPermissionTree);
                            try
                            {
                                if (!permissionTree.GetByKey(xmlReader["ID"].ToString()))
                                {
                                    permissionTree.PermissionID = xmlReader["ID"].ToString();
                                    //if (xmlReader["Text"].ToString() == "Leave Master")
                                    //    System.Diagnostics.Debugger.Break();
                                    permissionTree.Name = xmlReader["Text"].ToString();
                                    permissionTree.Options = SAPbobsCOM.BoUPTOptions.bou_FullNone;
                                    //permissionTree.Levels = 2;
                                    permissionTree.ParentID = fatherCardID;
                                    if (xmlReader["formuid"] != null)
                                        permissionTree.UserPermissionForms.FormType = xmlReader["formuid"].ToString();
                                    int Result = permissionTree.Add();
                                    if (Result != 0)
                                        ____bobCompany.GetLastErrorDescription();
                                    //SAPWrapper.Common.Global.WriteLog(__Company.GetLastErrorDescription(), System.Diagnostics.EventLogEntryType.Error);
                                }
                                //break;
                            }
                            catch (COMException e1)
                            {
                                __app.MessageBox(e1.Message, 1, "Ok", "", "");
                                //if (OnError != null)
                                //    OnError(e1, this.GetType(), e1.Message);
                                ////SAPWrapper.Common.Global.WriteLog(e1.Message, System.Diagnostics.EventLogEntryType.Error); 
                                //break;
                            }
                            finally
                            {
                                if (Marshal.IsComObject(permissionTree))
                                    Marshal.ReleaseComObject(permissionTree);
                            }
                        }
                    }
                }//end while
            }
            catch (COMException e1)
            {
                __app.MessageBox(e1.Message, 1, "Ok", "", "");
                //if (OnError != null)
                //    OnError(e1, this.GetType(), e1.Message);
                //SAPWrapper.Common.Global.WriteLog(e1.Message, System.Diagnostics.EventLogEntryType.Error);
            }
        }
        #endregion

        #region "Event Handling"

        void __app_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (BusinessObjectInfo.FormTypeEx == "CTIAMDC" && BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD && BusinessObjectInfo.BeforeAction == false)
                {
                    __Form = __app.Forms.GetForm("CTIAMDC", 0);
                    __Form.EnableMenu("1282", true);
                    //__Form.EnableMenu("1289", true);
                    __Form.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE;
                }

                if (BusinessObjectInfo.FormTypeEx == "CTIAMDC" && BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD && BusinessObjectInfo.ActionSuccess == true && BusinessObjectInfo.BeforeAction == false)
                {
                    __Form = __app.Forms.GetForm("CTIAMDC", 0);
                    __app.ActivateMenuItem("1289");

                }
            }
            catch (COMException e1)
            {
                __app.MessageBox(e1.Message, 0, "Ok", "", "");
            }
            catch (SqlException e1)
            {
                __app.MessageBox(e1.Message, 0, "Ok", "", "");
            }
            catch (Exception e1)
            {
                __app.MessageBox(e1.Message, 0, "Ok", "", "");
            }
        }

        void __app_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    {
                        Environment.Exit(0);
                        break;
                    }
            }
        }

        void __app_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (pVal.FormTypeEx == "AGR")
            {
                AR_ItemEvent(FormUID, ref pVal, out BubbleEvent);
            }

            if (pVal.FormTypeEx == "UDOB1CZ")
            {
                B1CZ_ItemEvent(FormUID, ref pVal, out BubbleEvent);
            }

            //if (pVal.FormTypeEx == "FreightCharges")
            //{
            //    Freight_ItemEvent(FormUID, ref pVal, out BubbleEvent);
            //}
        }

        void __app_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            #region "Main Menu"
            if (pVal.BeforeAction == false)
            {
                if (pVal.MenuUID == "MDC_MNU_GRP")
                {
                    XmlDocument oXMLDoc = new XmlDocument();
                    string MenuPath = Application.StartupPath + "\\SrfFiles\\AGR.srf";
                    oXMLDoc.Load(MenuPath);
                    __app.LoadBatchActions(oXMLDoc.InnerXml);
                    string UName = __app.Company.UserName.ToString();


                    oRs1 = (SAPbobsCOM.Recordset)____bobCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRs2 = (SAPbobsCOM.Recordset)____bobCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRs3 = (SAPbobsCOM.Recordset)____bobCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    try
                    {
                        Manual = false;
                        __Form = __app.Forms.GetForm("AGR", 0);
                        oGrid = __Form.Items.Item("grid").Specific;

                        __Form = __app.Forms.ActiveForm;
                        __Form.EnableMenu("4870", true);                     

                        __Form.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                        __Form.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized;


                        string SqlDoc = "";
                        SqlDoc = "CALL AGING_REPORT (NULL)";
                        oRs2.DoQuery(SqlDoc);


                        DT = null;

                        if (oRs2.RecordCount > 0)
                        {

                            if (DT == null)
                            {
                                DT = __Form.DataSources.DataTables.Add("DT1" + DateTime.Now.Second);
                            }

                            DT.ExecuteQuery(SqlDoc);

                            oGrid.DataTable = DT;
                            oGrid.AutoResizeColumns();

                            oGrid.Columns.Item("Customer Code").Editable = false;
                            oGrid.Columns.Item("Customer Name").Editable = false;

                            oGrid.Columns.Item("Customer Name").Editable = false;
                            oGrid.Columns.Item("Type").Editable = false;
                            SAPbouiCOM.EditTextColumn Typcol = (SAPbouiCOM.EditTextColumn)oGrid.Columns.Item("DocEntry");
                            Typcol.LinkedObjectType = "13";
                            oGrid.Columns.Item("TransType").Visible = false;
                            oGrid.Columns.Item("DocEntry").Editable = false;                           
                            
                            oGrid.Columns.Item("Document No.").Editable = false;
                            oGrid.Columns.Item("Customer Ref. No.").Editable = false;
                            oGrid.Columns.Item("Posting Date").Editable = false;
                            oGrid.Columns.Item("Due Date").Editable = false;
                            oGrid.Columns.Item("Future").Editable = false;
                            oGrid.Columns.Item("0-30 Days").Editable = false;
                            oGrid.Columns.Item("31-60 Days").Editable = false;
                            oGrid.Columns.Item("61-90 Days").Editable = false;
                            oGrid.Columns.Item("91-120 Days").Editable = false;
                            oGrid.Columns.Item("121+ Days").Editable = false;
                            oGrid.Columns.Item("Balance Due").Editable = false;

                            SAPbouiCOM.EditTextColumn col = (SAPbouiCOM.EditTextColumn)oGrid.Columns.Item("AWDocEntry");
                            col.LinkedObjectType = "UDOB1CZHDR";
                            col.Width = 15;

                            oGrid.Columns.Item("DocEntry").Editable = false;
                            oGrid.Columns.Item("Previous Collection Notes").Editable = false;                           

                            oGrid.CollapseLevel = 1;  
                            DataColumn column;                         
                          

                            // Create new DataColumn, set DataType, ColumnName and add to DataTable.
                            column = new DataColumn();
                            if (Htable.Columns.Contains("DocEntry"))
                            {
                                Htable.Columns.Remove("DocEntry");
                            }
                            column.DataType = System.Type.GetType("System.Int32");
                            column.ColumnName = "DocEntry";
                            Htable.Columns.Add(column);

                            // Create second column.
                            column = new DataColumn();

                            if (Htable.Columns.Contains("Collection Notes"))
                            {
                                Htable.Columns.Remove("Collection Notes");
                            }
                            column.DataType = Type.GetType("System.String");
                            column.ColumnName = "Collection Notes";
                            Htable.Columns.Add(column);
                        }


                    }
                    catch (Exception e)
                    {
                        MessageBox.Show(e.Message.ToString());
                        BubbleEvent = false;
                    }

                }


            }
            #endregion

            //#region "Navigate Menu"

            //if (pVal.BeforeAction == true)
            //{
            //    if (pVal.MenuUID == "1281")
            //    {
            //        //__XForm = __app.Forms.ActiveForm;
            //        //__XForm.Items.Item("Item_1").Enabled = true;   
            //    }
            //}
            //#endregion
        }

        #endregion
    }
}
