using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Runtime.InteropServices;
using SAPbobsCOM;
using System.Xml;
using System.Windows.Forms;

namespace AgingReport
{
    public partial class AddOnInfo
    {
        public void AR_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.BeforeAction == false && pVal.ItemUID == "1")
                {
                    InsertDate_AfterAddBtnPressed(FormUID, ref pVal, out BubbleEvent);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.BeforeAction == false && pVal.ItemUID == "FCW")
                {
                    FCW_AfterAddBtnPressed(FormUID, ref pVal, out BubbleEvent);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.BeforeAction == true && pVal.ItemUID == "1")
                {
                    InsertDate_BeforeAddBtnPressed(FormUID, ref pVal, out BubbleEvent);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_LOST_FOCUS && pVal.BeforeAction == false)
                {
                    LostFocussed(FormUID, ref pVal, out BubbleEvent);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.BeforeAction == false && pVal.ItemUID == "Item_1")
                {
                    Search_AfterAddBtnPressed(FormUID, ref pVal, out BubbleEvent);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED && pVal.BeforeAction == true)
                {
                    LinkBtn_BeforeBtnPressed(FormUID, ref pVal, out BubbleEvent);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED && pVal.BeforeAction == false)
                {
                    LinkBtn_AfterBtnPressed(FormUID, ref pVal, out BubbleEvent);
                }

                if (pVal.BeforeAction == true && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                {
                    ChosFrmList_Before(FormUID, ref pVal, out BubbleEvent);
                }
                if (pVal.BeforeAction == false && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                {
                    AfterCFL(FormUID, ref pVal, out BubbleEvent);
                }
            }
            catch (SqlException e)
            {
                __app.MessageBox(e.Message, 1, "Ok", "", "");
                BubbleEvent = false;
            }
            catch (COMException e)
            {
                __app.MessageBox(e.Message, 1, "Ok", "", "");
                BubbleEvent = false;
            }
            catch (Exception e)
            {
                __app.MessageBox(e.Message, 1, "Ok", "", "");
                BubbleEvent = false;
            }
        }

        private void FCW_AfterAddBtnPressed(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            __Form = __app.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);
            oGrid = __Form.Items.Item("grid").Specific;
            oGrid.AutoResizeColumns();
        }

        private void InsertDate_BeforeAddBtnPressed(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            __Form = __app.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);
            __Form.Items.Item("Item_0").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        }

        private void LostFocussed(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            __Form = __app.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);
            oGrid = __Form.Items.Item("grid").Specific;

            try
            {
                if (pVal.ColUID == "Collection Notes")
                {
                    int Index = oGrid.GetDataTableRowIndex(pVal.Row);
                    int DocEntry = oGrid.DataTable.GetValue("DocEntry", Index);
                    string ColNotes = oGrid.DataTable.GetValue("Collection Notes", Index);

                    Htable.TableName = "Historty";
                    Hrow = Htable.NewRow();

                    Hrow["DocEntry"] = DocEntry;
                    Hrow["Collection Notes"] = ColNotes;

                    Htable.Rows.Add(Hrow);
                }
            }
            catch (SqlException e)
            {
                __app.MessageBox(e.Message, 1, "Ok", "", "");
                BubbleEvent = false;
            }
            catch (COMException e)
            {
                __app.MessageBox(e.Message, 1, "Ok", "", "");
                BubbleEvent = false;
            }
            catch (Exception e)
            {
                __app.MessageBox(e.Message, 1, "Ok", "", "");
                BubbleEvent = false;
            }
        }

        private void LinkBtn_BeforeBtnPressed(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                if (pVal.ColUID == "DocEntry")
                {
                    int Index = oGrid.GetDataTableRowIndex(pVal.Row);
                    string Value = oGrid.DataTable.GetValue("TransType", Index);

                    SAPbouiCOM.EditTextColumn oColumns = (SAPbouiCOM.EditTextColumn)oGrid.Columns.Item("DocEntry");
                    oColumns.LinkedObjectType = Value; 

                }
            }
            catch (SqlException e)
            {
                __app.MessageBox(e.Message, 1, "Ok", "", "");
                BubbleEvent = false;
            }
            catch (COMException e)
            {
                __app.MessageBox(e.Message, 1, "Ok", "", "");
                BubbleEvent = false;
            }
            catch (Exception e)
            {
                __app.MessageBox(e.Message, 1, "Ok", "", "");
                BubbleEvent = false;
            }
        }

        private void Search_AfterAddBtnPressed(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            __Form = __app.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);
            oRs = (SAPbobsCOM.Recordset)____bobCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            oGrid = __Form.Items.Item("grid").Specific;

            //DT.Clear();
            //DT = null;

            string CardCode = __Form.Items.Item("Item_0").Specific.Value;

            string SqlDoc = "";
            if (!string.IsNullOrEmpty(CardCode))
            {
                SqlDoc = "CALL AGING_REPORT ('" + CardCode + "')";

            }
            else
            {
                SqlDoc = "CALL AGING_REPORT (NULL)";
            }

            oRs.DoQuery(SqlDoc);

            if (oRs.RecordCount > 0)
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

                SAPbouiCOM.EditTextColumn col = (SAPbouiCOM.EditTextColumn)oGrid.Columns.Item("AWDocEntry");
                col.LinkedObjectType = "UDOB1CZHDR";
                col.Width = 15;

                oGrid.Columns.Item("DocEntry").Editable = false;
                oGrid.Columns.Item("Balance Due").Editable = false;
                oGrid.Columns.Item("Previous Collection Notes").Editable = false;

                oGrid.CollapseLevel = 1;
            }

        }

        private void ChosFrmList_Before(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            __Form = __app.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

            try
            {
                if (pVal.ItemUID == "Item_0")
                {
                    SAPbouiCOM.IChooseFromListEvent chl = (SAPbouiCOM.IChooseFromListEvent)pVal;
                    SAPbouiCOM.ChooseFromList oCFL = null;
                    string cflID = chl.ChooseFromListUID;
                    oCFL = __Form.ChooseFromLists.Item(cflID);
                    oCFL.SetConditions(null);
                    SAPbouiCOM.Conditions oCons = oCFL.GetConditions();
                    SAPbouiCOM.Condition c = null;
                    c = oCons.Add();
                    c.BracketOpenNum = 1;
                    c.Alias = "CardType";
                    c.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    c.CondVal = "C";
                    c.BracketCloseNum = 1;
                    oCFL.SetConditions(oCons);
                }

            }
            catch (SqlException e)
            {
                __app.SetStatusBarMessage(e + "...", SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                BubbleEvent = false;
            }
            catch (COMException e)
            {
                __app.SetStatusBarMessage(e + "...", SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                BubbleEvent = false;
            }
            catch (Exception e)
            {
                __app.SetStatusBarMessage(e + "...", SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                BubbleEvent = false;
            }
        }

        private void AfterCFL(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            __Form = __app.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

            try
            {
                if (pVal.ItemUID == "Item_0")
                {
                    SAPbouiCOM.EditText BPCode = (SAPbouiCOM.EditText)__Form.Items.Item("Item_0").Specific;

                    SAPbouiCOM.IChooseFromListEvent chl = (SAPbouiCOM.IChooseFromListEvent)pVal;
                    SAPbouiCOM.ChooseFromList oCFL = null;
                    string cflID = chl.ChooseFromListUID;
                    oCFL = __Form.ChooseFromLists.Item(cflID);
                    SAPbouiCOM.DataTable dTable = chl.SelectedObjects;

                    if (dTable != null)
                    {
                        try
                        {
                            BPCode.String = dTable.GetValue("CardCode", 0);
                            __Form.Items.Item("Item_1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        }
                        catch (Exception ex)
                        {
                            __Form.Items.Item("Item_1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        }


                    }
                }

            }
            catch (SqlException e)
            {
                __app.MessageBox(e.Message, 1, "Ok", "", "");
                BubbleEvent = false;
            }
            catch (COMException e)
            {
                __app.MessageBox(e.Message, 1, "Ok", "", "");
                BubbleEvent = false;
            }
            catch (Exception e)
            {
                __app.MessageBox(e.Message, 1, "Ok", "", "");
                BubbleEvent = false;
            }
        }

        private void LinkBtn_AfterBtnPressed(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {

                __Form = __app.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);
                oGrid = __Form.Items.Item("grid").Specific;

                if (pVal.ColUID == "AWDocEntry")
                {
                    int Index = oGrid.GetDataTableRowIndex(pVal.Row);

                    int Value = oGrid.DataTable.GetValue("DocEntry", Index);                    
                   
                    XmlDocument oXMLDoc = new XmlDocument();
                    string MenuPath = Application.StartupPath + "\\SrfFiles\\B1Cuz.srf";
                    oXMLDoc.Load(MenuPath);

                    try
                    {
                        __app.LoadBatchActions(oXMLDoc.InnerXml);
                        __Form1 = __app.Forms.GetForm("UDOB1CZ", 0);
                        oMatrix = __Form1.Items.Item("Item_1").Specific;
                        oMatrix.AutoResizeColumns();

                        __Form1.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                        __Form1.Items.Item("Item_0").Specific.Value = Value;
                        __Form1.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    }
                    catch (Exception)
                    {
                        __Form1 = __app.Forms.GetForm("UDOB1CZ", 0);
                        oMatrix = __Form1.Items.Item("Item_1").Specific;
                        oMatrix.AutoResizeColumns();

                        __Form1.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                        __Form1.Items.Item("Item_0").Specific.Value = Value;
                        __Form1.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        __Form1.Select();
                    }
                    

                    
                }

            }
            catch (SqlException e)
            {
                __app.MessageBox(e.Message, 1, "Ok", "", "");
                BubbleEvent = false;
            }
            catch (COMException e)
            {
                __app.MessageBox(e.Message, 1, "Ok", "", "");
                BubbleEvent = false;
            }
            catch (Exception e)
            {
                __app.MessageBox(e.Message, 1, "Ok", "", "");
                BubbleEvent = false;
            }
        }

        private void InsertDate_AfterAddBtnPressed(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                __Form = __app.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);
                
                oRs = (SAPbobsCOM.Recordset)____bobCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRs1 = (SAPbobsCOM.Recordset)____bobCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oGrid = __Form.Items.Item("grid").Specific;

                int Progress = 0;               
                SAPbouiCOM.ProgressBar oProgressBar = null;
                oProgressBar = __app.StatusBar.CreateProgressBar("Updating History ", Htable.Rows.Count, true);

                #region UpdateHistoryTable
                for (int i = 0; i < Htable.Rows.Count; i++)
                {
                    try
                    {
                        int DocEntry = int.Parse(Htable.Rows[i][0].ToString()); //oGrid.DataTable.GetValue("DocEntry", i);
                        string CollectionNotes = Htable.Rows[i][1].ToString(); //oGrid.DataTable.GetValue("As Collection Notes", i);

                        oProgressBar.Text = "Updating History " + i + " rows...";
                        Progress += 1;
                        oProgressBar.Value = Progress;

                        if (!string.IsNullOrEmpty(CollectionNotes))
                        {

                            SAPbobsCOM.GeneralService oGeneralService;
                            SAPbobsCOM.GeneralData oGeneralData;
                            SAPbobsCOM.CompanyService sCmp = null;
                            SAPbobsCOM.GeneralData oChild;
                            SAPbobsCOM.GeneralDataCollection oChildren;
                            SAPbobsCOM.GeneralDataParams oGeneralParams;

                            sCmp = ____bobCompany.GetCompanyService();

                            oGeneralService = sCmp.GetGeneralService("UDOB1CZHDR");
                            oGeneralData = (SAPbobsCOM.GeneralData)oGeneralService.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralData);

                            string SqlQry = "";
                            SqlQry = "SELECT \"Code\" FROM \"@AW_B1CZHDR\" Where \"U_DocEntry\"=" + DocEntry;
                            oRs.DoQuery(SqlQry);

                            if (oRs.RecordCount > 0)
                            {
                                ____bobCompany.StartTransaction();
                                string SqlCode = "";
                                SqlCode = "SELECT T0.\"Code\",T1.\"U_AsColNotes\",T1.\"U_UpdateDate\"  FROM \"@AW_B1CZHDR\" T0 INNER JOIN \"@AW_B1CZDTL\" T1 ON T1.\"U_DocEntry\"=T0.\"U_DocEntry\" and T0.\"Code\"=T1.\"Code\" " +
                                           " Where T0.\"U_DocEntry\"=" + DocEntry;
                                oRs1.DoQuery(SqlCode);

                                oGeneralData.SetProperty("Code", Convert.ToString(oRs1.Fields.Item("Code").Value));
                                oGeneralData.SetProperty("Name", Convert.ToString(oRs.Fields.Item("Code").Value));
                                oGeneralData.SetProperty("U_DocEntry", DocEntry);


                                oChildren = oGeneralData.Child("AW_B1CZDTL");

                                for (int j = 0; j < oRs1.RecordCount; j++)
                                {
                                    oChild = oChildren.Add();
                                    oChild.SetProperty("U_DocEntry", DocEntry);
                                    oChild.SetProperty("U_AsColNotes", oRs1.Fields.Item("U_AsColNotes").Value);
                                    oChild.SetProperty("U_UpdateDate", oRs1.Fields.Item("U_UpdateDate").Value);
                                    oRs1.MoveNext();
                                }

                                oChild = oChildren.Add();
                                oChild.SetProperty("U_DocEntry", DocEntry);
                                oChild.SetProperty("U_AsColNotes", CollectionNotes);
                                oChild.SetProperty("U_UpdateDate", DateTime.Now);

                                oGeneralService.Update(oGeneralData);
                                if (____bobCompany.InTransaction)
                                {
                                    ____bobCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                                }
                            }
                            else
                            {
                                ____bobCompany.StartTransaction();
                                string SqlCode = "";
                                SqlCode = "SELECT IFNULL(MAX((CAST(\"Code\" AS int))),0)+1 \"MaxCode\",IFNULL(MAX((CAST(\"DocEntry\" AS int))),0)+1 \"MaxDocEntry\" FROM \"@AW_B1CZHDR\"";
                                oRs.DoQuery(SqlCode);

                                oGeneralData.SetProperty("Code", Convert.ToString(oRs.Fields.Item("MaxCode").Value));
                                oGeneralData.SetProperty("Name", Convert.ToString(oRs.Fields.Item("MaxCode").Value));
                                oGeneralData.SetProperty("U_DocEntry", DocEntry);

                                oChildren = oGeneralData.Child("AW_B1CZDTL");
                                oChild = oChildren.Add();
                                oChild.SetProperty("U_DocEntry", DocEntry);
                                oChild.SetProperty("U_AsColNotes", CollectionNotes);
                                oChild.SetProperty("U_UpdateDate", DateTime.Now.Date);

                                oGeneralService.Add(oGeneralData);
                                if (____bobCompany.InTransaction)
                                {
                                    ____bobCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                                }
                            }
                        }
                    }
                    catch (Exception)
                    {

                    }
                }

                oProgressBar.Stop();
                Htable.Clear();
                DT.Clear();
                DT = null;
                #endregion

                string CardCode = __Form.Items.Item("Item_0").Specific.Value;

                string SqlDoc = "";
                if (!string.IsNullOrEmpty(CardCode))
                {
                    SqlDoc = "CALL AGING_REPORT ('" + CardCode + "')";

                }
                else
                {
                    SqlDoc = "CALL AGING_REPORT (NULL)";
                }
                oRs.DoQuery(SqlDoc);

                if (oRs.RecordCount > 0)
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

                    SAPbouiCOM.EditTextColumn col = (SAPbouiCOM.EditTextColumn)oGrid.Columns.Item("AWDocEntry");
                    col.LinkedObjectType = "UDOB1CZHDR";
                    col.Width = 15;

                    oGrid.Columns.Item("DocEntry").Editable = false;
                    oGrid.Columns.Item("Balance Due").Editable = false;
                    oGrid.Columns.Item("Previous Collection Notes").Editable = false;

                    oGrid.CollapseLevel = 1;
                }

            }
            catch (SqlException e)
            {
                __app.MessageBox(e.Message, 1, "Ok", "", "");
                BubbleEvent = false;
                
            }
            catch (COMException e)
            {
                __app.MessageBox(e.Message, 1, "Ok", "", "");
                BubbleEvent = false;

            }
            catch (Exception e)
            {
                __app.MessageBox(e.Message, 1, "Ok", "", "");
                BubbleEvent = false;

            }
        }
    }
}
