using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;
using System.Data;
using System.Runtime.InteropServices;

namespace BDO_Localisation_AddOn
{
    class BDOSDepreciationAccrualDocument
    {
        const int clientHeight = 540;
        const int clientWidth = 600;

        public static void createDocumentUDO(out string errorText)
        {
            errorText = null;
            string tableName = "BDOSDEPACR";
            string description = "Depreciation Accrual Document";

            SAPbobsCOM.UserObjectsMD oUserObjectsMD = null;
            oUserObjectsMD = (SAPbobsCOM.UserObjectsMD)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);
            Marshal.ReleaseComObject(oUserObjectsMD);

            int result = UDO.addUserTable(tableName, description, SAPbobsCOM.BoUTBTableType.bott_Document, out errorText);

            if (result != 0)
            {
                return;
            }

            Dictionary<string, object> fieldskeysMap;

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "DocDate");
            fieldskeysMap.Add("TableName", "BDOSDEPACR");
            fieldskeysMap.Add("Description", "Posting Date");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Date);
            fieldskeysMap.Add("EditSize", 15);
            fieldskeysMap.Add("Mandatory", SAPbobsCOM.BoYesNoEnum.tYES);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "AccrMnth");
            fieldskeysMap.Add("TableName", "BDOSDEPACR");
            fieldskeysMap.Add("Description", "Accrual Month");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Date);
            fieldskeysMap.Add("EditSize", 15);
            fieldskeysMap.Add("Mandatory", SAPbobsCOM.BoYesNoEnum.tYES);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "Retirement");
            fieldskeysMap.Add("TableName", "BDOSDEPACR");
            fieldskeysMap.Add("Description", "Retirement");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 1);
            fieldskeysMap.Add("DefaultValue", "N");

            UDO.addUserTableFields(fieldskeysMap, out errorText);


            tableName = "BDOSDEPAC1";
            description = "Depreciation Accrual Child1";

            result = UDO.addUserTable(tableName, description, SAPbobsCOM.BoUTBTableType.bott_DocumentLines, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "ItemCode");
            fieldskeysMap.Add("TableName", "BDOSDEPAC1");
            fieldskeysMap.Add("Description", "Item Code");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "DistNumber");
            fieldskeysMap.Add("TableName", "BDOSDEPAC1");
            fieldskeysMap.Add("Description", "Dist Number");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSUsLife");
            fieldskeysMap.Add("TableName", "BDOSDEPAC1");
            fieldskeysMap.Add("Description", "Useful Life");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Numeric);
            fieldskeysMap.Add("EditSize", 10);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "Project");
            fieldskeysMap.Add("TableName", "BDOSDEPAC1");
            fieldskeysMap.Add("Description", "Project");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 150);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "Quantity");
            fieldskeysMap.Add("TableName", "BDOSDEPAC1");
            fieldskeysMap.Add("Description", "Quantity");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Quantity);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "DeprAmt");
            fieldskeysMap.Add("TableName", "BDOSDEPAC1");
            fieldskeysMap.Add("Description", "Depreciation Amount");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "AccmDprAmt");
            fieldskeysMap.Add("TableName", "BDOSDEPAC1");
            fieldskeysMap.Add("Description", "Accumulated Depreciation Amount");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "RetDprAmt");
            fieldskeysMap.Add("TableName", "BDOSDEPAC1");
            fieldskeysMap.Add("Description", "Retired Depreciation Amount");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);
        }

        public static void registerUDO(out string errorText)
        {
            errorText = null;
            string code = "UDO_F_BDOSDEPACR_D"; //20 characters (must include at least one alphabetical character).
            Dictionary<string, object> formProperties;

            formProperties = new Dictionary<string, object>();
            formProperties.Add("Name", "Depreciation Accrual Document"); //100 characters
            formProperties.Add("TableName", "BDOSDEPACR");
            formProperties.Add("ObjectType", SAPbobsCOM.BoUDOObjType.boud_Document);
            formProperties.Add("CanCancel", SAPbobsCOM.BoYesNoEnum.tYES);
            formProperties.Add("CanClose", SAPbobsCOM.BoYesNoEnum.tYES);
            formProperties.Add("CanCreateDefaultForm", SAPbobsCOM.BoYesNoEnum.tYES);
            formProperties.Add("CanDelete", SAPbobsCOM.BoYesNoEnum.tNO);
            formProperties.Add("CanFind", SAPbobsCOM.BoYesNoEnum.tYES);
            formProperties.Add("CanYearTransfer", SAPbobsCOM.BoYesNoEnum.tYES);
            formProperties.Add("ManageSeries", SAPbobsCOM.BoYesNoEnum.tNO);
            formProperties.Add("CanLog", SAPbobsCOM.BoYesNoEnum.tYES);

            List<Dictionary<string, object>> listFindColumns = new List<Dictionary<string, object>>();
            List<Dictionary<string, object>> listFormColumns = new List<Dictionary<string, object>>();
            List<Dictionary<string, object>> listChildTables = new List<Dictionary<string, object>>();

            Dictionary<string, object> fieldskeysMap;

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "U_DocDate");
            fieldskeysMap.Add("ColumnDescription", "Posting Date");
            listFindColumns.Add(fieldskeysMap);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "U_AccrMnth");
            fieldskeysMap.Add("ColumnDescription", "Accrual Month");
            listFindColumns.Add(fieldskeysMap);

            formProperties.Add("FindColumns", listFindColumns);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("FormColumnAlias", "DocEntry");
            fieldskeysMap.Add("FormColumnDescription", "DocEntry");
            listFormColumns.Add(fieldskeysMap);

            formProperties.Add("FormColumns", listFormColumns);

            //ცხრილური ნაწილები
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("TableName", "BDOSDEPAC1");
            fieldskeysMap.Add("ObjectName", "BDOSDEPAC1");
            listChildTables.Add(fieldskeysMap);

            formProperties.Add("ChildTables", listChildTables);
            //ცხრილური ნაწილები

            UDO.registerUDO(code, formProperties, out errorText);
        }

        public static void addMenus()
        {
            try
            {
                SAPbouiCOM.MenuItem fatherMenuItem = Program.uiApp.Menus.Item("9201");

                // Add a pop-up menu item
                SAPbouiCOM.MenuCreationParams oCreationPackage = (SAPbouiCOM.MenuCreationParams)Program.uiApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                oCreationPackage.Checked = false;
                oCreationPackage.Enabled = true;
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "UDO_F_BDOSDEPACR_D";
                oCreationPackage.String = BDOSResources.getTranslate("DepreciationAccrual");
                oCreationPackage.Position = fatherMenuItem.SubMenus.Count - 1;

                SAPbouiCOM.MenuItem menuItem = fatherMenuItem.SubMenus.AddEx(oCreationPackage);
            }
            catch
            {

            }
        }

        public static void uiApp_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            SAPbouiCOM.Form oDocForm = Program.uiApp.Forms.ActiveForm;

            if (pVal.MenuUID == "BDOSAddRow")
                addMatrixRow(oDocForm);
            else if (pVal.MenuUID == "BDOSDelRow")
                delMatrixRow(oDocForm);
        }

        public static void uiApp_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(BusinessObjectInfo.FormTypeEx, Program.currentFormCount);

            if (oForm.TypeEx == "UDO_FT_UDO_F_BDOSDEPACR_D")
            {
                //if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD)
                //{
                //    if (BusinessObjectInfo.BeforeAction)
                //    {
                //        //checkDoc(oForm, out errorText);
                //        if (errorText != null)
                //        {
                //            Program.uiApp.StatusBar.SetSystemMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                //            Program.uiApp.MessageBox(BDOSResources.getTranslate("OperationUnsuccesfullSeeLog"));
                //            BubbleEvent = false;
                //        }
                //        else
                //        {
                //            //updateAsset(oForm, false, out errorText);
                //            if (errorText != null)
                //            {
                //                Program.uiApp.StatusBar.SetSystemMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                //                Program.uiApp.MessageBox(BDOSResources.getTranslate("OperationUnsuccesfullSeeLog"));
                //                BubbleEvent = false;
                //            }
                //        }
                //    }
                //}
                if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE && BusinessObjectInfo.BeforeAction)
                {
                    if (Program.cancellationTrans && Program.canceledDocEntry != 0)
                    {
                        int answer = 0;
                        answer = Program.uiApp.MessageBox(BDOSResources.getTranslate("DoYouReallyWantCancelDoc") + "?", 1, BDOSResources.getTranslate("Yes"), BDOSResources.getTranslate("No"), "");
                        if (answer == 1)
                            cancellation(oForm, Program.canceledDocEntry, out errorText);
                        else
                            BubbleEvent = false;

                        Program.canceledDocEntry = 0;
                    }
                }
                if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD && !BusinessObjectInfo.BeforeAction)
                {
                    formDataLoad(oForm);
                }
            }
        }

        public static void cancellation(SAPbouiCOM.Form oForm, int docEntry, out string errorText)
        {
            errorText = null;
            JournalEntry.cancellation(oForm, docEntry, "UDO_F_BDOSDEPACR_D", out errorText);
            if (!string.IsNullOrEmpty(errorText))
                throw new Exception(errorText);
        }

        public static void uiApp_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD && pVal.BeforeAction)
                {
                    createFormItems(oForm);
                    Program.FORM_LOAD_FOR_VISIBLE = true;
                    Program.FORM_LOAD_FOR_ACTIVATE = true;
                }

                else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE && !pVal.BeforeAction)
                {
                    if (Program.FORM_LOAD_FOR_VISIBLE)
                    {
                        setSizeForm(oForm);
                        oForm.Freeze(true);
                        oForm.Title = BDOSResources.getTranslate("DepreciationAccrual");
                        oForm.Freeze(false);
                        Program.FORM_LOAD_FOR_VISIBLE = false;
                    }
                }

                else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_RESIZE && !pVal.BeforeAction)
                {
                    resizeForm(oForm);
                }

                else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE && !pVal.BeforeAction)
                {
                    if (Program.FORM_LOAD_FOR_ACTIVATE)
                    {
                        oForm.Freeze(true);
                        try
                        {
                            SAPbouiCOM.StaticText staticText = oForm.Items.Item("0_U_S").Specific;
                            staticText.Caption = BDOSResources.getTranslate("DocEntry");

                            Program.FORM_LOAD_FOR_ACTIVATE = false;
                        }
                        catch (Exception ex)
                        {
                            throw new Exception(ex.Message);
                        }
                        finally
                        {
                            oForm.Freeze(false);
                        }
                    }
                }
            }
        }

        public static void uiApp_RightClickEvent(SAPbouiCOM.Form oForm, SAPbouiCOM.ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (eventInfo.ItemUID == "DepAcrMTR")
            {
                SAPbouiCOM.MenuItem oMenuItem;
                SAPbouiCOM.Menus oMenus;
                SAPbouiCOM.MenuCreationParams oCreationPackage;

                try
                {
                    oCreationPackage = (SAPbouiCOM.MenuCreationParams)Program.uiApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                    oCreationPackage.UniqueID = "BDOSAddRow";
                    oCreationPackage.String = BDOSResources.getTranslate("AddNewRow");
                    oCreationPackage.Enabled = true;
                    oCreationPackage.Position = -1;

                    oMenuItem = Program.uiApp.Menus.Item("1280");
                    oMenus = oMenuItem.SubMenus;
                    oMenus.AddEx(oCreationPackage);
                }
                catch (Exception ex)
                {
                    string errorText = ex.Message;
                }

                try
                {
                    oCreationPackage = (SAPbouiCOM.MenuCreationParams)Program.uiApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                    oCreationPackage.UniqueID = "BDOSDelRow";
                    oCreationPackage.String = BDOSResources.getTranslate("DeleteRow");
                    oCreationPackage.Enabled = true;
                    oCreationPackage.Position = -1;

                    oMenuItem = Program.uiApp.Menus.Item("1280");
                    oMenus = oMenuItem.SubMenus;
                    oMenus.AddEx(oCreationPackage);
                }
                catch (Exception ex)
                {
                    string errorText = ex.Message;
                }
            }
            else
            {
                try
                {
                    Program.uiApp.Menus.RemoveEx("BDOSAddRow");
                }
                catch { }
                try
                {
                    Program.uiApp.Menus.RemoveEx("BDOSDelRow");
                }
                catch { }
            }
        }

        public static void createFormItems(SAPbouiCOM.Form oForm)
        {
            string errorText;

            oForm.AutoManaged = true;

            Dictionary<string, object> formItems;
            string itemName;

            int height = 15;
            int width_s = 120;
            int width_e = 140;
            int left_s = 6;
            int left_e = left_s + width_s + 20;
            int top = 5;
            int left_s2 = 305;
            int left_e2 = left_s2 + width_s + 20;
            int top2 = 5;

            top += (height + 1);

            formItems = new Dictionary<string, object>();
            itemName = "AccrMnthS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("AccrualMonth"));
            formItems.Add("LinkTo", "AccrMnthE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }

            formItems = new Dictionary<string, object>();
            itemName = "AccrMnthE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSDEPACR");
            formItems.Add("Alias", "U_AccrMnth");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }
            oForm.Items.Item(itemName).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False); //All modes
            oForm.Items.Item(itemName).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True); //Find mode

            top = top + height + 1;

            FormsB1.addChooseFromList(oForm, false, "30", "JournalEntryCFL");
            formItems = new Dictionary<string, object>();
            itemName = "TransIdS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("TransactionNo"));
            formItems.Add("LinkTo", "TransIdE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }

            formItems = new Dictionary<string, object>();
            itemName = "TransIdE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "UserDataSources");
            formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
            formItems.Add("Length", 11);
            formItems.Add("TableName", "");
            formItems.Add("Alias", itemName);
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }
            oForm.Items.Item(itemName).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False); //All modes

            formItems = new Dictionary<string, object>();
            itemName = "TransIdLB"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            formItems.Add("Left", left_e - 20);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("LinkTo", "TransIdE");
            formItems.Add("LinkedObjectType", "30"); //Journal Entry

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "RtrmntCH"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSDEPACR");
            formItems.Add("Alias", "U_Retirement");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Retirement"));
            formItems.Add("ValOff", "N");
            formItems.Add("ValOn", "Y");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }
            oForm.Items.Item(itemName).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False); //All modes

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "No.S"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s2);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top2);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Number"));
            formItems.Add("LinkTo", "DocNumE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }

            formItems = new Dictionary<string, object>();
            itemName = "DocNumE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSDEPACR");
            formItems.Add("Alias", "DocNum");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e2);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top2);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }
            oForm.Items.Item(itemName).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False); //All modes
            oForm.Items.Item(itemName).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True); //Find mode

            top2 += height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "StatusS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s2);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top2);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Status"));
            formItems.Add("LinkTo", "StatusC");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }

            formItems = new Dictionary<string, object>();
            itemName = "StatusC"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSDEPACR");
            formItems.Add("Alias", "Status");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            formItems.Add("Left", left_e2);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top2);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
            formItems.Add("DisplayDesc", true);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }
            oForm.Items.Item(itemName).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False); //All modes
            oForm.Items.Item(itemName).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True); //Find mode

            top2 += height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "CanceledS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s2);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top2);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Canceled"));
            formItems.Add("LinkTo", "CanceledC");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }

            formItems = new Dictionary<string, object>();
            itemName = "CanceledC"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSDEPACR");
            formItems.Add("Alias", "Canceled");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            formItems.Add("Left", left_e2);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top2);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
            formItems.Add("DisplayDesc", true);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }
            oForm.Items.Item(itemName).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False); //All modes
            oForm.Items.Item(itemName).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True); //Find mode

            top2 += height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "CreateDatS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s2);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top2);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("CreateDate"));
            formItems.Add("LinkTo", "CreateDatE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }

            formItems = new Dictionary<string, object>();
            itemName = "CreateDatE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSDEPACR");
            formItems.Add("Alias", "CreateDate");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e2);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top2);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }
            oForm.Items.Item(itemName).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False); //All modes
            oForm.Items.Item(itemName).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True); //Find mode

            top2 += height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "UpdateDatS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s2);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top2);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("UpdateDate"));
            formItems.Add("LinkTo", "UpdateDatE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }

            formItems = new Dictionary<string, object>();
            itemName = "UpdateDatE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSDEPACR");
            formItems.Add("Alias", "UpdateDate");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e2);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top2);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }
            oForm.Items.Item(itemName).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False); //All modes
            oForm.Items.Item(itemName).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True); //Find mode

            top2 += height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "DocDateS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s2);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top2);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("PostingDate"));
            formItems.Add("LinkTo", "DocDateE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }

            formItems = new Dictionary<string, object>();
            itemName = "DocDateE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSDEPACR");
            formItems.Add("Alias", "U_DocDate");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e2);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top2);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }
            oForm.Items.Item(itemName).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False); //OK mode
            oForm.Items.Item(itemName).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 8, SAPbouiCOM.BoModeVisualBehavior.mvb_False); //View mode

            top += (3 * height + 1);

            formItems = new Dictionary<string, object>();
            itemName = "FillMTR"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            formItems.Add("Left", left_s);
            formItems.Add("Width", 70);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Fill"));
            formItems.Add("Visible", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }
            oForm.Items.Item(itemName).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False); //OK mode
            oForm.Items.Item(itemName).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 8, SAPbouiCOM.BoModeVisualBehavior.mvb_False); //View mode

            top += height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "DepAcrMTR"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
            formItems.Add("Left", left_s);
            formItems.Add("Top", top);
            formItems.Add("UID", itemName);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }

            oForm.DataSources.DBDataSources.Add("@BDOSDEPAC1");

            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("DepAcrMTR").Specific;
            oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
            SAPbouiCOM.Columns oColumns = oMatrix.Columns;

            SAPbouiCOM.Column oColumn = oColumns.Add("LineID", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "#";
            oColumn.Editable = false;
            oColumn.DataBind.SetBound(true, "@BDOSDEPAC1", "LineId");

            oColumn = oColumns.Add("ItemCode", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("AssetCode");
            oColumn.Editable = false;
            oColumn.ExtendedObject.LinkedObjectType = "4";
            oColumn.DataBind.SetBound(true, "@BDOSDEPAC1", "U_ItemCode");

            oColumn = oColumns.Add("DistNumber", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("DistNumber");
            oColumn.Editable = false;
            oColumn.DataBind.SetBound(true, "@BDOSDEPAC1", "U_DistNumber");

            oColumn = oColumns.Add("BDOSUsLife", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("UsefulLife");
            oColumn.Editable = false;
            oColumn.DataBind.SetBound(true, "@BDOSDEPAC1", "U_BDOSUsLife");

            oColumn = oColumns.Add("Quantity", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Quantity");
            oColumn.Editable = false;
            oColumn.DataBind.SetBound(true, "@BDOSDEPAC1", "U_Quantity");

            oColumn = oColumns.Add("DeprAmt", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("DepreciationAmount");
            oColumn.Editable = false;
            oColumn.DataBind.SetBound(true, "@BDOSDEPAC1", "U_DeprAmt");

            oColumn = oColumns.Add("AccmDprAmt", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("AccumulatedDepreciationAmount");
            oColumn.Editable = false;
            oColumn.DataBind.SetBound(true, "@BDOSDEPAC1", "U_AccmDprAmt");

            oColumn = oColumns.Add("Project", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Project");
            oColumn.ExtendedObject.LinkedObjectType = "144";
            oColumn.Editable = false;
            oColumn.DataBind.SetBound(true, "@BDOSDEPAC1", "U_Project");

            //oColumn = oColumns.Add("InvEntry", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            //oColumn.TitleObject.Caption = BDOSResources.getTranslate("InvEntry");
            //oColumn.Editable = false;
            //oColumn.Visible = false;
            //oColumn.DataBind.SetBound(true, "@BDOSDEPAC1", "U_InvEntry");

            //oColumn = oColumns.Add("InvType", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            //oColumn.TitleObject.Caption = BDOSResources.getTranslate("InvType");
            //oColumn.Editable = false;
            //oColumn.Visible = false;
            //oColumn.DataBind.SetBound(true, "@BDOSDEPAC1", "U_InvType");

            top = top + oForm.Items.Item("DepAcrMTR").Height + 10;

            formItems = new Dictionary<string, object>();
            itemName = "CreatorS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Creator"));
            formItems.Add("LinkTo", "CreatorE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }

            formItems = new Dictionary<string, object>();
            itemName = "CreatorE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSDEPACR");
            formItems.Add("Alias", "Creator");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }
            oForm.Items.Item(itemName).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False); //All modes
            oForm.Items.Item(itemName).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True); //Find mode

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "RemarksS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Remarks"));
            formItems.Add("LinkTo", "RemarksE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }

            formItems = new Dictionary<string, object>();
            itemName = "RemarksE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSDEPACR");
            formItems.Add("Alias", "Remark");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EXTEDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e * 3);
            formItems.Add("Top", top);
            formItems.Add("Height", 3 * height);
            formItems.Add("UID", itemName);
            formItems.Add("ScrollBars", SAPbouiCOM.BoScrollBars.sb_Vertical);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }
        }

        public static void resizeForm(SAPbouiCOM.Form oForm)
        {
            try
            {
                oForm.Freeze(true);

                int width_s = 120;
                int width_e = 140;
                int left_s = 6;
                int left_e = left_s + width_s + 20;

                oForm.Items.Item("0_U_E").Left = left_e;
                oForm.Items.Item("0_U_E").Width = width_e;

                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("DepAcrMTR").Specific;
                int mtrWidth = oForm.ClientWidth;
                oForm.Items.Item("DepAcrMTR").Width = mtrWidth;
                oForm.Items.Item("DepAcrMTR").Height = oForm.ClientHeight / 2;
                FormsB1.resetWidthMatrixColumns(oForm, "DepAcrMTR", "LineID", mtrWidth);

                int height = 15;
                int top = oForm.Items.Item("DepAcrMTR").Top;

                if (oForm.ClientHeight <= clientHeight)
                    top += oForm.Items.Item("DepAcrMTR").Height + 10;
                else
                    top = oForm.ClientHeight - (8 * height + 10);

                oForm.Items.Item("CreatorS").Top = top;
                oForm.Items.Item("CreatorE").Top = top;

                top += height + 1;

                oForm.Items.Item("RemarksS").Top = top;
                oForm.Items.Item("RemarksE").Top = top;

                top += 6 * height + 1;

                oForm.Items.Item("1").Top = top;
                oForm.Items.Item("2").Top = top;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        public static void setSizeForm(SAPbouiCOM.Form oForm)
        {
            oForm.Freeze(true);
            try
            {
                oForm.ClientHeight = clientHeight;
                oForm.ClientWidth = clientWidth;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        public static void addMatrixRow(SAPbouiCOM.Form oForm)
        {
            oForm.Freeze(true);
            try
            {
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("DepAcrMTR").Specific;
                oMatrix.FlushToDataSource();

                SAPbouiCOM.DBDataSource oDBDataSourceMTR = oForm.DataSources.DBDataSources.Item("@BDOSDEPAC1");
                if (!string.IsNullOrEmpty(oDBDataSourceMTR.GetValue("U_ItemCode", oDBDataSourceMTR.Size - 1)))
                    oDBDataSourceMTR.InsertRecord(oDBDataSourceMTR.Size);
                oDBDataSourceMTR.SetValue("LineId", oDBDataSourceMTR.Size - 1, oDBDataSourceMTR.Size.ToString());

                oMatrix.LoadFromDataSource();

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        public static void delMatrixRow(SAPbouiCOM.Form oForm)
        {
            oForm.Freeze(true);
            try
            {
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("DepAcrMTR").Specific;
                oMatrix.FlushToDataSource();
                int firstRow = 0;
                int row = 0;
                int deletedRowCount = 0;

                while (row != -1)
                {
                    row = oMatrix.GetNextSelectedRow(firstRow, SAPbouiCOM.BoOrderType.ot_RowOrder);
                    if (row > -1)
                    {
                        deletedRowCount++;
                        oForm.DataSources.DBDataSources.Item("@BDOSDEPAC1").RemoveRecord(row - deletedRowCount);
                        firstRow = row;
                    }
                }

                SAPbouiCOM.DBDataSource oDBDataSourceMTR = oForm.DataSources.DBDataSources.Item("@BDOSDEPAC1");
                int rowCount = oDBDataSourceMTR.Size;

                for (int i = 1; i <= rowCount; i++)
                {
                    string crLnCode = oDBDataSourceMTR.GetValue("U_ItemCode", i - 1);
                    if (!string.IsNullOrEmpty(crLnCode))
                        oDBDataSourceMTR.SetValue("LineId", i - 1, i.ToString());
                }
                oMatrix.LoadFromDataSource();

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE && deletedRowCount > 0)
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        public static DataTable createAdditionalEntries(SAPbouiCOM.Form oForm, SAPbobsCOM.GeneralData oGeneralData, int docEntry, DateTime DocDate, DataTable AccountTable = null)
        {
            DataTable jeLines = JournalEntry.JournalEntryTable();
            SAPbobsCOM.GeneralDataCollection oChild = null;
            SAPbouiCOM.DBDataSource oDBDataSource = null;
            string isRetirement;
            if (AccountTable == null)
                AccountTable = CommonFunctions.GetOACTTable();

            int jeCount = 0;

            if (oForm == null)
            {
                isRetirement = oGeneralData.GetProperty("U_Retirement");
                oChild = oGeneralData.Child("BDOSDEPAC1");
                jeCount = oChild.Count;
            }
            else
            {
                isRetirement = oForm.DataSources.DBDataSources.Item("@BDOSDEPACR").GetValue("U_Retirement", 0);
                oDBDataSource = oForm.DataSources.DBDataSources.Item("@BDOSDEPAC1");
                jeCount = oDBDataSource.Size;
            }

            for (int i = 0; i < jeCount; i++)
            {
                decimal deprAmt = Convert.ToDecimal(CommonFunctions.getChildOrDbDataSourceValue(oDBDataSource, oChild, null, "U_DeprAmt", i), CultureInfo.InvariantCulture);
                decimal accmDprAmt = Convert.ToDecimal(CommonFunctions.getChildOrDbDataSourceValue(oDBDataSource, oChild, null, "U_AccmDprAmt", i), CultureInfo.InvariantCulture);
                string itemCode = ((string)CommonFunctions.getChildOrDbDataSourceValue(oDBDataSource, oChild, null, "U_ItemCode", i)).Trim();
                string prjCode = ((string)CommonFunctions.getChildOrDbDataSourceValue(oDBDataSource, oChild, null, "U_Project", i)).Trim();
                //string invEntry = ((string)CommonFunctions.getChildOrDbDataSourceValue(oDBDataSource, oChild, null, "U_InvEntry", i)).Trim();
                //string invType = ((string)CommonFunctions.getChildOrDbDataSourceValue(oDBDataSource, oChild, null, "U_InvType", i)).Trim();

                SAPbobsCOM.Items oItem;
                oItem = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems);
                oItem.GetByKey(itemCode);

                SAPbobsCOM.ItemGroups oItemGroup;
                oItemGroup = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItemGroups);
                oItemGroup.GetByKey(oItem.ItemsGroupCode);

                string AccDepAccount = oItemGroup.UserFields.Fields.Item("U_BDOSAccDep").Value.ToString();
                string ExpDepAccount = oItemGroup.UserFields.Fields.Item("U_BDOSExpDep").Value.ToString();
                string SaleCostAc = oItemGroup.CostAccount;
                JournalEntry.AddJournalEntryRow(AccountTable, jeLines, "Full", ExpDepAccount, AccDepAccount, deprAmt, 0, "", "", "", "", "", "", prjCode, "", "");

                if (isRetirement == "Y")
                {
                    JournalEntry.AddJournalEntryRow(AccountTable, jeLines, "Full", AccDepAccount, SaleCostAc, accmDprAmt, 0, "", "", "", "", "", "", prjCode, "", "");
                }
            }

            if (jeLines.Rows.Count > 0)
            {
                jeLines = jeLines.AsEnumerable()
                      .GroupBy(row => new
                      {
                          AccountCode = row.Field<string>("AccountCode"),
                          ShortName = row.Field<string>("ShortName"),
                          ContraAccount = row.Field<string>("ContraAccount"),
                          FCCurrency = row.Field<string>("FCCurrency"),
                          CostingCode = row.Field<string>("CostingCode"),
                          CostingCode2 = row.Field<string>("CostingCode2"),
                          CostingCode3 = row.Field<string>("CostingCode3"),
                          CostingCode4 = row.Field<string>("CostingCode4"),
                          CostingCode5 = row.Field<string>("CostingCode5"),
                          ProjectCode = row.Field<string>("ProjectCode"),
                          VatGroupCode = row.Field<string>("VatGroup"),
                          U_BDOSEmpID = row.Field<string>("U_BDOSEmpID")
                      })
                      .Select(g =>
                      {
                          var row = jeLines.NewRow();
                          row["AccountCode"] = g.Key.AccountCode;
                          row["ShortName"] = g.Key.ShortName;
                          row["ContraAccount"] = g.Key.ContraAccount;
                          row["FCCurrency"] = g.Key.FCCurrency;
                          row["CostingCode"] = g.Key.CostingCode;
                          row["CostingCode2"] = g.Key.CostingCode2;
                          row["CostingCode3"] = g.Key.CostingCode3;
                          row["CostingCode4"] = g.Key.CostingCode4;
                          row["CostingCode5"] = g.Key.CostingCode5;
                          row["ProjectCode"] = g.Key.ProjectCode;
                          row["VatGroup"] = g.Key.VatGroupCode;
                          row["U_BDOSEmpID"] = g.Key.U_BDOSEmpID;

                          row["Credit"] = g.Sum(r => r.Field<double>("Credit"));
                          row["Debit"] = g.Sum(r => r.Field<double>("Debit"));
                          row["FCCredit"] = g.Sum(r => r.Field<double>("FCCredit"));
                          row["FCDebit"] = g.Sum(r => r.Field<double>("FCDebit"));
                          return row;
                      }).CopyToDataTable();
            }
            return jeLines;
        }

        public static void JrnEntry(string DocEntry, string DocNum, DateTime DocDate, DataTable JrnLinesDT, string emloyeeCode, out string errorText)
        {
            errorText = null;

            try
            {
                JournalEntry.JrnEntry(DocEntry, "UDO_F_BDOSDEPACR_D", "Depreciatin accruing: " + DocNum, DocDate, JrnLinesDT, out errorText);
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
        }

        public static void formDataLoad(SAPbouiCOM.Form oForm)
        {
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                // გატარებები
                SAPbouiCOM.DBDataSource oDBDataSource = oForm.DataSources.DBDataSources.Item("@BDOSDEPACR");
                string Ref1 = oDBDataSource.GetValue("DocEntry", 0);
                string Ref2 = "UDO_F_BDOSDEPACR_D";
                string strdate = oDBDataSource.GetValue("U_DocDate", 0);
                DateTime DocDate = DateTime.TryParseExact(strdate, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DocDate) == false ? DateTime.Now : DocDate;

                string query = "SELECT " +
                                "\"TransId\" " +
                                "FROM \"OJDT\"  " +
                                "WHERE \"Ref1\" = '" + Ref1 + "' " +
                                "AND \"Ref2\" = '" + Ref2 + "' " +
                                "AND \"RefDate\" = '" + DocDate.ToString("yyyyMMdd") + "' ";
                oRecordSet.DoQuery(query);

                if (!oRecordSet.EoF)
                    oForm.DataSources.UserDataSources.Item("TransIdE").ValueEx = oRecordSet.Fields.Item("TransId").Value.ToString();
                else
                    oForm.DataSources.UserDataSources.Item("TransIdE").ValueEx = "";
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                Marshal.ReleaseComObject(oRecordSet);
            }
        }

        //public static decimal getDepreciationPriceDistNumber(string ItemCode, string DistNumber)
        //{

        //    string query = @"select
	       //                  ""@BDOSDEPAC1"".""U_ItemCode"",
	       //                  ""@BDOSDEPAC1"".""U_DistNumber"",
	       //                  SUM(""@BDOSDEPAC1"".""U_DeprAmt"") as ""U_DeprAmt"",
	       //                  SUM(""@BDOSDEPAC1"".""U_Quantity"") as ""U_Quantity"" 
        //                from ""@BDOSDEPAC1"" 
        //                inner join ""@BDOSDEPACR"" on ""@BDOSDEPACR"".""DocEntry"" = ""@BDOSDEPAC1"".""DocEntry"" 
        //                and ""@BDOSDEPACR"".""Canceled"" = 'N' and ""@BDOSDEPAC1"".""U_ItemCode"" = '" + ItemCode + @"' and ""@BDOSDEPAC1"".""U_DistNumber"" = '" + DistNumber + @"'
        //                group by ""@BDOSDEPAC1"".""U_ItemCode"",
	       //                  ""@BDOSDEPAC1"".""U_DistNumber""";

        //    SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //    oRecordSet.DoQuery(query);

        //    if (!oRecordSet.EoF)
        //    {

        //        decimal DeprQty = Convert.ToDecimal(oRecordSet.Fields.Item("U_Quantity").Value);
        //        decimal AlrDeprAmt = 0;
        //        if (DeprQty > 0)
        //        {
        //            AlrDeprAmt = Convert.ToDecimal(oRecordSet.Fields.Item("U_DeprAmt").Value) / DeprQty;
        //        }

        //        return AlrDeprAmt;
        //    }

        //    return 0;
        //}
    }
}


