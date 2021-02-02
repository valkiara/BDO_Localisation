using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace BDO_Localisation_AddOn
{
    static partial class StockTransfer
    {
        public static bool ReserveInvoiceAsService = false;

        public static void createFormItems(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            Dictionary<string, object> formItems = new Dictionary<string, object>();
            string itemName = "";
            SAPbouiCOM.Item oItem = oForm.Items.Item("31");
            int height = oItem.Height;
            int top = oForm.Items.Item("18").Top - 7;
            int left_s = oForm.Items.Item("33").Left;
            int left_e = oItem.Left;
            int width = oItem.Width;

            string caption = BDOSResources.getTranslate("CreateWaybill");
            itemName = "BDO_WblTxt"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", caption);
            formItems.Add("TextStyle", 4);
            formItems.Add("FontSize", 10);
            formItems.Add("Enabled", true);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            bool multiSelection = false;
            string objectType = "UDO_F_BDO_WBLD_D"; //Waybill document
            string uniqueID_WaybillCFL = "Waybill_CFL";
            FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_WaybillCFL);

            formItems = new Dictionary<string, object>();
            itemName = "BDO_WblDoc"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "UserDataSources");
            formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
            formItems.Add("Length", 11);
            formItems.Add("TableName", "");
            formItems.Add("Alias", itemName);
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width / 2);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("AffectsFormMode", false);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", false);
            formItems.Add("ChooseFromListUID", uniqueID_WaybillCFL);
            formItems.Add("ChooseFromListAlias", "DocEntry");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "BDO_WblLB"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            formItems.Add("Left", left_e - 20);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("LinkTo", "BDO_WblDoc");
            formItems.Add("LinkedObjectType", objectType);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "BDO_WblID"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", "");
            //formItems.Add("TextStyle", 4);
            formItems.Add("FontSize", 10);
            formItems.Add("Enabled", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "BDO_WblNum"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", "");
            //formItems.Add("TextStyle", 4);
            formItems.Add("FontSize", 10);
            formItems.Add("Enabled", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "BDO_WblSts"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width * 2);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", "");
            //formItems.Add("TextStyle", 4);
            formItems.Add("FontSize", 10);
            formItems.Add("Enabled", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            itemName = "";

            SAPbouiCOM.Item oItem_s = oForm.Items.Item("1470000099");
            SAPbouiCOM.Item oItem_e = oForm.Items.Item("1470000101");

            left_s = oItem_s.Left;
            left_e = oItem_e.Left;
            height = oItem_e.Height;
            //top = oItem_e.Top;
            top = top + height + 10;
            int width_s = oItem_s.Width;
            int width_e = oItem_e.Width;

            multiSelection = false;
            objectType = "63";
            string uniqueID_lf_prj_CFL = "Prj_CFL";
            FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_lf_prj_CFL);

            formItems = new Dictionary<string, object>();
            itemName = "PrjCodeS"; //10 characters
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
            formItems.Add("TableName", "OWTR");
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
            fieldskeysMap.Add("TableName", "OWTR");
            fieldskeysMap.Add("Description", "From Project");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "CANCELED");
            fieldskeysMap.Add("TableName", "OWTR");
            fieldskeysMap.Add("Description", "Cancelled");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 1);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

        }

        public static void formDataLoad(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            try
            {
                int docEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("OWTR").GetValue("DocEntry", 0));
                string caption = BDOSResources.getTranslate("CreateWaybill");
                int wblDocEntry = 0;
                string wblID = "";
                string wblNum = "";
                string wblSts = "";

                if (docEntry != 0)
                {
                    Dictionary<string, string> wblDocInfo = BDO_Waybills.getWaybillDocumentInfo(docEntry, "67", out errorText);
                    wblDocEntry = Convert.ToInt32(wblDocInfo["DocEntry"]);
                    wblID = wblDocInfo["wblID"];
                    wblNum = wblDocInfo["number"];
                    wblSts = wblDocInfo["status"];

                    if (wblDocEntry != 0)
                    {
                        caption = BDOSResources.getTranslate("WaybillDocEntry");
                    }
                }
                else
                {
                    caption = BDOSResources.getTranslate("CreateWaybill");
                    wblDocEntry = 0;
                }

                oForm.DataSources.UserDataSources.Item("BDO_WblDoc").ValueEx = wblDocEntry == 0 ? "" : wblDocEntry.ToString();
                SAPbouiCOM.StaticText oStaticText = (SAPbouiCOM.StaticText)oForm.Items.Item("BDO_WblTxt").Specific;
                oStaticText.Caption = caption; oStaticText = (SAPbouiCOM.StaticText)oForm.Items.Item("BDO_WblID").Specific;
                oStaticText.Caption = wblID != "" ? "ID : " + wblID : "";
                oStaticText = (SAPbouiCOM.StaticText)oForm.Items.Item("BDO_WblNum").Specific;
                oStaticText.Caption = wblNum != "" ? "№ " + wblNum : "";
                oStaticText = (SAPbouiCOM.StaticText)oForm.Items.Item("BDO_WblSts").Specific;
                oStaticText.Caption = wblSts != "" ? BDOSResources.getTranslate("Status") + " : " + wblSts : "";


                //oForm.Items.Item("BDO_WblDoc").Enabled = (oForm.DataSources.DBDataSources.Item(0).GetValue("CANCELED", 0) == "N");
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

        public static void cancellation()
        {
            string errorText;

            try
            {
                Dictionary<string, string> wblDocInfo = BDO_Waybills.getWaybillDocumentInfo(Program.canceledDocEntry, "67", out errorText);
                int wblDocEntry = Convert.ToInt32(wblDocInfo["DocEntry"]);

                if (wblDocEntry != 0)
                {
                    int answer = Program.uiApp.MessageBox(BDOSResources.getTranslate("DocumentLinkedToWaybillCancel"), 1, BDOSResources.getTranslate("Yes"), BDOSResources.getTranslate("No"), "");
                    string operation = answer == 1 ? "Update" : "Cancel";
                    BDO_Waybills.cancellation(wblDocEntry, operation, out errorText);
                }

                SAPbobsCOM.StockTransfer oStockTransfer = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer);
                if (oStockTransfer.GetByKey(Program.canceledDocEntry))
                {
                    oStockTransfer.UserFields.Fields.Item("U_CANCELED").Value = "Y";
                }

                int resultCode = oStockTransfer.Update();
                if (resultCode == 0)
                {
                    StringBuilder query = new StringBuilder();
                    query.Append("SELECT TOP 1 \"DocEntry\", \"DocNum\", \"Comments\" \n");
                    query.Append("FROM OWTR \n");
                    query.Append("WHERE CONTAINS(\"Comments\", '*" + oStockTransfer.DocNum + "*') ORDER BY \"DocEntry\" DESC");

                    Marshal.ReleaseComObject(oStockTransfer);

                    SAPbobsCOM.Recordset oRecordset = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    oRecordset.DoQuery(query.ToString());
                    if (!oRecordset.EoF)
                    {
                        int docEntry = oRecordset.Fields.Item("DocEntry").Value;

                        oStockTransfer = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer);
                        if (oStockTransfer.GetByKey(docEntry))
                            oStockTransfer.UserFields.Fields.Item("U_CANCELED").Value = "Y";
                        resultCode = oStockTransfer.Update();

                        if (resultCode != 0)
                        {
                            int errCode;
                            string errMsg;
                            Program.oCompany.GetLastError(out errCode, out errMsg);
                            Program.uiApp.StatusBar.SetSystemMessage(errMsg, SAPbouiCOM.BoMessageTime.bmt_Short);
                        }
                        Marshal.ReleaseComObject(oStockTransfer);
                    }
                }
                else
                {
                    int errCode;
                    string errMsg;
                    Program.oCompany.GetLastError(out errCode, out errMsg);
                    Program.uiApp.StatusBar.SetSystemMessage(errMsg, SAPbouiCOM.BoMessageTime.bmt_Short);
                }
            }
            catch (Exception ex)
            {
                Program.uiApp.StatusBar.SetSystemMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short);
            }
            finally
            {
                GC.Collect();
            }
        }

        public static void setVisibleFormItems(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            SAPbouiCOM.Item oItem = null;

            try
            {
                string docEntry = oForm.DataSources.DBDataSources.Item("OWTR").GetValue("DocEntry", 0).Trim();
                bool docEntryIsEmpty = string.IsNullOrEmpty(docEntry);

                oForm.Items.Item("PrjCodeE").Enabled = (docEntryIsEmpty == true);


                if (oForm.DataSources.DBDataSources.Item("OWTR").GetValue("DocStatus", 0) == "O")
                {
                    oItem = oForm.Items.Item("BDO_WblTxt");
                    oItem.Enabled = true;
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
            finally
            {
                GC.Collect();
            }

            FormsB1.WB_TAX_AuthorizationsItems(oForm);
        }

        public static void uiApp_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(BusinessObjectInfo.FormTypeEx, Program.currentFormCount);

            if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD)
            {
                if (!BusinessObjectInfo.BeforeAction && !BusinessObjectInfo.ActionSuccess)
                {
                    BubbleEvent = false;
                }

                if (BusinessObjectInfo.BeforeAction)
                {                   
                    //ძირითადი საშუალებების შემოწმება
                    if (BatchNumberSelection.SelectedBatches != null)
                    {
                        CommonFunctions.blockAssetInvoice(oForm, "OWTR", out var rejectionAsset);
                        if (rejectionAsset)
                        {
                            Program.uiApp.MessageBox(BDOSResources.getTranslate("DocumentCannotBeAdded") + " : " + BDOSResources.getTranslate("ThereIsDepreciationAmountsInCurrentMonthForItem"));
                            BubbleEvent = false;
                        }
                    }
                }

                if (BusinessObjectInfo.ActionSuccess != BusinessObjectInfo.BeforeAction)
                {
                    //დოკუმენტის გატარების დროს გატარდეს ბუღლტრული გატარება
                    SAPbouiCOM.DBDataSource DocDBSource = oForm.DataSources.DBDataSources.Item(0);
                    if (DocDBSource.GetValue("CANCELED", 0) == "N")
                    {
                        string DocEntry = DocDBSource.GetValue("DocEntry", 0);
                        string fromPrjCode = DocDBSource.GetValue("U_BDOSFrPrj", 0).Trim();

                        CommonFunctions.StartTransaction();

                        UpdateJournalEntry(DocEntry, "67", fromPrjCode, out errorText);

                        if (!string.IsNullOrEmpty(errorText))
                        {
                            Program.uiApp.MessageBox(errorText);
                            BubbleEvent = false;
                        }

                        //თუ დოკუმენტი გატარდა, მერე ვაკეთებს ბუღალტრულ გატარებას
                        if (BusinessObjectInfo.ActionSuccess && !BusinessObjectInfo.BeforeAction)
                        {
                            CommonFunctions.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                            BatchNumberSelection.SelectedBatches = null;
                        }
                        else
                            CommonFunctions.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                    }
                }
            }

            if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD & !BusinessObjectInfo.BeforeAction)
            {
                setVisibleFormItems(oForm, out errorText);
            }

            if (oForm.TypeEx == "940")
            {
                if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD && !BusinessObjectInfo.BeforeAction)
                {
                    formDataLoad(oForm, out errorText);
                    //setVisibleFormItems( oForm, out errorText);
                }
                if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE && !BusinessObjectInfo.BeforeAction && BusinessObjectInfo.ActionSuccess)
                {
                    if (Program.cancellationTrans && Program.canceledDocEntry != 0)
                    {
                        cancellation();
                        Program.canceledDocEntry = 0;
                    }
                }

                if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD)
                {
                    if (BusinessObjectInfo.BeforeAction)
                    {
                        SAPbouiCOM.DBDataSource DocDBSource = oForm.DataSources.DBDataSources.Item(0);
                        if (DocDBSource.GetValue("CANCELED", 0) == "N")
                        {
                            //უარყოფითი ნაშთების კონტროლი დოკ.თარიღით
                            bool rejection = false;
                            CommonFunctions.blockNegativeStockByDocDate(oForm, "OWTR", "WTR1", "FromWhsCod", out rejection);
                            if (rejection)
                            {
                                Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("DocumentCannotBeAdded"));
                                BubbleEvent = false;
                            }
                        }
                    }
                }
            }
        }

        public static void uiApp_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent, out string errorText)
        {
            errorText = null;
            BubbleEvent = true;

            ////----------------------------->Cancel <-----------------------------
            //SAPbouiCOM.Form oForm = Program.uiApp.Forms.ActiveForm;

            //if (pVal.BeforeAction && pVal.MenuUID == "1284")
            //{
            //    //ძირითადი საშუალებების შემოწმება
            //    bool rejectionAsset = false;
            //    CommonFunctions.blockAssetInvoice(oForm, "OWTR", out rejectionAsset);
            //    if (rejectionAsset)
            //    {
            //        BubbleEvent = false;
            //    }
            //}
        }

        public static void uiApp_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD & pVal.BeforeAction == true)
                {
                    createFormItems(oForm, out errorText);
                    Program.FORM_LOAD_FOR_VISIBLE = true;
                    Program.FORM_LOAD_FOR_ACTIVATE = true;

                    formDataLoad(oForm, out errorText);
                    setVisibleFormItems(oForm, out errorText);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE & pVal.BeforeAction == false)
                {
                    if (Program.FORM_LOAD_FOR_ACTIVATE == true)
                    {
                        setVisibleFormItems(oForm, out errorText);
                        Program.FORM_LOAD_FOR_ACTIVATE = false;
                    }
                }

                if (pVal.ItemUID == "PrjCodeE" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                {
                    SAPbouiCOM.IChooseFromListEvent oCFLEvento = null;
                    oCFLEvento = ((SAPbouiCOM.IChooseFromListEvent)(pVal));

                    chooseFromList(oForm, oCFLEvento, pVal.ItemUID, pVal.Row, pVal.BeforeAction, out errorText);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                //& oCFLEvento.ChooseFromListUID == "Waybill_CFL" & pVal.BeforeAction == true)
                {
                    if (pVal.BeforeAction == true)
                    {
                        SAPbouiCOM.ChooseFromListEvent oCFLEvento = null;
                        oCFLEvento = ((SAPbouiCOM.ChooseFromListEvent)(pVal));
                        string sCFL_ID = oCFLEvento.ChooseFromListUID;
                        SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(sCFL_ID);
                        SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();

                        if (oCFLEvento.ChooseFromListUID == "Waybill_CFL")
                        {
                            string query = @"Select ""DocEntry"" from ""@BDO_WBLD"" where ""U_baseDoc"" =0 and  ""U_baseDocT"" = 67";

                            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                            oRecordSet.DoQuery(query);

                            SAPbouiCOM.Condition oCon = null;
                            while (!oRecordSet.EoF)
                            {
                                oCon = oCons.Add();
                                oCon.Alias = "DocEntry";
                                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                                oCon.CondVal = oRecordSet.Fields.Item("DocEntry").Value.ToString();

                                oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR;

                                oRecordSet.MoveNext();
                            }
                            oCon = oCons.Add();
                            oCon.Alias = "DocEntry";
                            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                            //oCon.CondVal = "";


                            oCFL.SetConditions(oCons);
                        }
                    }
                    else
                    {
                        SAPbouiCOM.ChooseFromListEvent oCFLEvento = null;
                        oCFLEvento = ((SAPbouiCOM.ChooseFromListEvent)(pVal));

                        if (oCFLEvento.ChooseFromListUID == "Waybill_CFL")
                        {
                            SAPbouiCOM.DataTable oDataTableSelectedObjects = oCFLEvento.SelectedObjects;

                            if (oDataTableSelectedObjects == null)
                            {
                                return;
                            }

                            int answer = Program.uiApp.MessageBox(BDOSResources.getTranslate("LinkWaybillToDocument"), 1, BDOSResources.getTranslate("Yes"), BDOSResources.getTranslate("No"), "");
                            if (answer == 2)
                            {
                                return;
                            }

                            oForm.Freeze(true);

                            int newDocEntry = oDataTableSelectedObjects.GetValue("DocEntry", 0);
                            int docEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("OWTR").GetValue("DocEntry", 0));


                            //არჩეულ ზედნადებში უნდა ჩავწეროთ ამ დოკუმენტის ნომრები
                            SAPbobsCOM.CompanyService oCompanyService = null;
                            SAPbobsCOM.GeneralService oGeneralService = null;
                            SAPbobsCOM.GeneralData oGeneralData = null;
                            SAPbobsCOM.GeneralDataParams oGeneralParams = null;
                            oCompanyService = Program.oCompany.GetCompanyService();
                            oGeneralService = oCompanyService.GetGeneralService("UDO_F_BDO_WBLD_D");
                            oGeneralData = ((SAPbobsCOM.GeneralData)(oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)));
                            //Get UDO record
                            oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                            oGeneralParams.SetProperty("DocEntry", newDocEntry);
                            oGeneralData = oGeneralService.GetByParams(oGeneralParams);

                            oGeneralData.SetProperty("U_baseDoc", docEntry);
                            oGeneralData.SetProperty("U_baseDTxt", docEntry.ToString());
                            oGeneralService.Update(oGeneralData);

                            oForm.DataSources.UserDataSources.Item("BDO_WblDoc").ValueEx = newDocEntry.ToString();
                            formDataLoad(oForm, out errorText);

                            BubbleEvent = true;
                            oForm.Freeze(false);
                            oForm.Update();

                        }
                    }
                }

                if (pVal.ItemUID == "BDO_WblTxt" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED & pVal.BeforeAction == false)
                {
                    oForm.Freeze(true);

                    int docEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("OWTR").GetValue("DocEntry", 0));
                    string cancelled = oForm.DataSources.DBDataSources.Item("OWTR").GetValue("CANCELED", 0).Trim();
                    string BDO_WblDoc = oForm.DataSources.UserDataSources.Item("BDO_WblDoc").ValueEx;
                    int newDocEntry = 0;

                    if (docEntry != 0 & (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_VIEW_MODE))
                    {
                        if (BDO_WblDoc == "" && cancelled == "N")
                        {

                            string objectType = "67";
                            BDO_Waybills.createDocument(objectType, docEntry, null, null, null, null, out newDocEntry, out errorText);
                            oForm.DataSources.UserDataSources.Item("BDO_WblDoc").ValueEx = newDocEntry.ToString();
                            if (errorText == null & newDocEntry != 0)
                            {
                                Program.uiApp.MessageBox(BDOSResources.getTranslate("WaybillCreatedSuccesfully") + " DocEntry : " + newDocEntry);
                                formDataLoad(oForm, out errorText);
                            }
                            else
                            {
                                Program.uiApp.MessageBox(errorText);
                            }
                        }
                        else if (cancelled != "N")
                        {
                            errorText = BDOSResources.getTranslate("DocumentMustNotBeCancelledOrCancellation");
                        }
                        else if (BDO_WblDoc != "")
                        {
                            errorText = BDOSResources.getTranslate("DocumentLinkedToWaybill");
                        }
                        BubbleEvent = true;
                    }
                    else
                    {
                        errorText = BDOSResources.getTranslate("ToCreateWaybillWriteDocument");
                    }

                    oForm.Freeze(false);
                    oForm.Update();

                    if (newDocEntry != 0)
                    {
                        Program.uiApp.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_UserDefinedObject, "UDO_F_BDO_WBLD_D", newDocEntry.ToString());
                    }
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
