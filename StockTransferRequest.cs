using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BDO_Localisation_AddOn
{
    static partial class StockTransferRequest
    {
        public static void createFormItems(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            Dictionary<string, object> formItems = new Dictionary<string, object>();

            SAPbouiCOM.Item oItem_s = oForm.Items.Item("1470000099");
            SAPbouiCOM.Item oItem_e = oForm.Items.Item("1470000101");

            int left_s = oItem_s.Left;
            int left_e = oItem_e.Left;
            int height = oItem_e.Height;
            int top = oItem_e.Top + height + 10;
            int width_s = oItem_s.Width;
            int width_e = oItem_e.Width;

            bool multiSelection = false;
            string objectType = "63";
            string uniqueID_lf_prj_CFL = "Prj_CFL";
            FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_lf_prj_CFL);

            formItems = new Dictionary<string, object>();
            string itemName = "PrjCodeS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("FromProject"));
            formItems.Add("LinkTo", "PrjCodeE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "PrjCodeE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OWTQ");
            formItems.Add("Alias", "U_BDOSFrPrj");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", true);
            formItems.Add("ChooseFromListUID", uniqueID_lf_prj_CFL);
            formItems.Add("ChooseFromListAlias", "PrjCode");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            //golden errow
            formItems = new Dictionary<string, object>();
            itemName = "PrjCodeLB"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            formItems.Add("Left", left_e - 20);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("LinkTo", "PrjCodeE");
            formItems.Add("LinkedObjectType", objectType);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            GC.Collect();
        }

        public static void createUserFields(out string errorText)
        {
            errorText = null;
            Dictionary<string, object> fieldskeysMap;

            //From Project
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSFrPrj");
            fieldskeysMap.Add("TableName", "OWTQ");
            fieldskeysMap.Add("Description", "From Project");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);
        }

        public static void setVisibleFormItems(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            try
            {
                string docEntry = oForm.DataSources.DBDataSources.Item("OWTQ").GetValue("DocEntry", 0).Trim();
                bool docEntryIsEmpty = string.IsNullOrEmpty(docEntry);

                oForm.Items.Item("PrjCodeE").Enabled = (docEntryIsEmpty == true);
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
            finally
            {
                GC.Collect();
            }
        }

        public static void uiApp_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(BusinessObjectInfo.FormTypeEx, Program.currentFormCount);

            if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD)
            {
                if (BusinessObjectInfo.BeforeAction == false && BusinessObjectInfo.ActionSuccess == false)
                {
                    BubbleEvent = false;
                }

                if (BusinessObjectInfo.ActionSuccess != BusinessObjectInfo.BeforeAction)
                {
                    //დოკუმენტის გატარების დროს გატარდეს ბუღლტრული გატარება
                    SAPbouiCOM.DBDataSource DocDBSource = oForm.DataSources.DBDataSources.Item(0);
                    if (DocDBSource.GetValue("CANCELED", 0) == "N")
                    {
                        string DocEntry = DocDBSource.GetValue("DocEntry", 0);
                        string fromPrjCode = DocDBSource.GetValue("U_BDOSFrPrj", 0).Trim();
                        //SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("23").Specific;

                        CommonFunctions.StartTransaction();

                        UpdateJournalEntry(DocEntry, "67", fromPrjCode, out errorText);

                        if (!string.IsNullOrEmpty(errorText))
                        {
                            Program.uiApp.MessageBox(errorText);
                            BubbleEvent = false;
                        }

                        //თუ დოკუმენტი გატარდა, მერე ვაკეთებს ბუღალტრულ გატარებას
                        if (BusinessObjectInfo.ActionSuccess == true && BusinessObjectInfo.BeforeAction == false)
                        {
                            CommonFunctions.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                        }
                        else
                        {
                            CommonFunctions.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                        }
                    }
                }
            }

            if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD && !BusinessObjectInfo.BeforeAction)
            {
                setVisibleFormItems(oForm, out errorText);
            }
        }

        public static void uiApp_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD && pVal.BeforeAction == true)
                {
                    createFormItems(oForm, out errorText);
                    Program.FORM_LOAD_FOR_VISIBLE = true;
                    Program.FORM_LOAD_FOR_ACTIVATE = true;
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE && pVal.BeforeAction == false)
                {
                    if (Program.FORM_LOAD_FOR_ACTIVATE == true)
                    {
                        setVisibleFormItems(oForm, out errorText);
                        Program.FORM_LOAD_FOR_ACTIVATE = false;
                    }
                }

                if (pVal.ItemUID == "PrjCodeE" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                {
                    SAPbouiCOM.IChooseFromListEvent oCFLEvento = null;
                    oCFLEvento = ((SAPbouiCOM.IChooseFromListEvent)(pVal));

                    chooseFromList(oForm, oCFLEvento, pVal.ItemUID, pVal.Row, pVal.BeforeAction, out errorText);
                }
            }
        }

        public static void UpdateJournalEntry(string DocEntry, string TransType, string fromPrjCode, out string errorText)
        {
            errorText = "";

            if (DocEntry != "")
            {
                SAPbobsCOM.Recordset oRecordSet_OIVL = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string query_OIVL = "";

                SAPbobsCOM.Recordset oRecordSet_Update = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string query_Update = "";

                SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string query = @"SELECT 
                            *  
                            FROM ""OJDT"" 
                            WHERE ""StornoToTr"" IS NULL   
                            AND ""TransType"" = '" + TransType + @"'  
                            AND ""CreatedBy"" = '" + DocEntry + "' ";

                oRecordSet.DoQuery(query);

                if (!oRecordSet.EoF)
                {
                    SAPbobsCOM.JournalEntries oJounalEntry = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
                    oJounalEntry.GetByKey(oRecordSet.Fields.Item("TransId").Value);


                    for (int i = 0; i < oJounalEntry.Lines.Count; i++)
                    {
                        oJounalEntry.Lines.SetCurrentLine(i);
                        if (oJounalEntry.Lines.Credit > 0)
                        {
                            //int docLine = oJounalEntry.Lines.DocumentLine;
                            //string prjCode = oMatrix.Columns.Item("U_BDOSFrPrj").Cells.Item(docLine).Specific.Value;

                            oJounalEntry.Lines.ProjectCode = fromPrjCode;

                            //OIVL ცხრილის აფდეითი                            
                            query_OIVL = @"SELECT ""MessageID"" 
                                                    FROM ""OIVL"" 
                                                    WHERE ""OutQty"" > 0 
                                                            AND ""TransType"" = '" + TransType + @"' 
                                                            AND ""CreatedBy"" = '" + DocEntry + @"' ";
                            //AND ""DocLineNum"" = '" + oJounalEntry.Lines.DocumentLine + "' ";


                            oRecordSet_OIVL.DoQuery(query_OIVL);

                            if (!oRecordSet_OIVL.EoF)
                            {
                                query_Update = @"UPDATE ""OILM"" SET ""PrjCode"" = '" + fromPrjCode + @"' where ""MessageID"" = '" + oRecordSet_OIVL.Fields.Item("MessageID").Value + "'";
                                oRecordSet_Update.DoQuery(query_Update);
                            }
                        }
                    }

                    int updateCode = 0;
                    updateCode = oJounalEntry.Update();

                    if (updateCode != 0)
                    {
                        Program.oCompany.GetLastError(out updateCode, out errorText);
                    }
                }
            }
        }

        public static void chooseFromList(SAPbouiCOM.Form oForm, SAPbouiCOM.IChooseFromListEvent oCFLEvento, string itemUID, int row, bool beforeAction, out string errorText)
        {
            errorText = null;

            try
            {
                string sCFL_ID = oCFLEvento.ChooseFromListUID;
                SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(sCFL_ID);

                if (beforeAction == true)
                {

                }
                else
                {
                    SAPbouiCOM.DataTable oDataTable = null;
                    oDataTable = oCFLEvento.SelectedObjects;

                    if (oDataTable != null)
                    {
                        if (sCFL_ID == "Prj_CFL")
                        {
                            string PrjCode = Convert.ToString(oDataTable.GetValue("PrjCode", 0));

                            try
                            {
                                //SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("23").Specific));
                                //oMatrix.Columns.Item("U_BDOSFrPrj").Cells.Item(row).Specific.Value = PrjCode;

                                oForm.Items.Item("PrjCodeE").Specific.Value = PrjCode;
                            }
                            catch { }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                int errCode;
                string errMsg;

                Program.oCompany.GetLastError(out errCode, out errMsg);
                errorText = BDOSResources.getTranslate("ErrorDescription") + " : " + errMsg + "! " + BDOSResources.getTranslate("Code") + " : " + errCode + "! " + BDOSResources.getTranslate("OtherInfo") + " : " + ex.Message;
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
