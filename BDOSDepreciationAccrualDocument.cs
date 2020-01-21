using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;
using System.Data;

namespace BDO_Localisation_AddOn
{
    class BDOSDepreciationAccrualDocument
    {

        public static bool openFormEvent = false;

        public static void createDocumentUDO(out string errorText)
        {
            errorText = null;
            string tableName = "BDOSDEPACR";
            string description = "Depreciation Accrual Document";

            SAPbobsCOM.UserObjectsMD oUserObjectsMD = null;
            oUserObjectsMD = (SAPbobsCOM.UserObjectsMD)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectsMD);

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

            

            //ცხრილური ნაწილი
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
            fieldskeysMap.Add("Name", "AlrDeprAmt");
            fieldskeysMap.Add("TableName", "BDOSDEPAC1");
            fieldskeysMap.Add("Description", "Already Depreciation Amount");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "InvEntry");
            fieldskeysMap.Add("TableName", "BDOSDEPAC1");
            fieldskeysMap.Add("Description", "Invoice doc entry");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 150);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "InvType");
            fieldskeysMap.Add("TableName", "BDOSDEPAC1");
            fieldskeysMap.Add("Description", "Invoice doc type");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 150);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            GC.Collect();

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

            GC.Collect();
        }

        public static void addMenus()
        {
            SAPbouiCOM.MenuItem menuItem;
            SAPbouiCOM.MenuItem fatherMenuItem;
            SAPbouiCOM.MenuCreationParams oCreationPackage;

            try
            {
                fatherMenuItem = Program.uiApp.Menus.Item("9201");

                // Add a pop-up menu item
                oCreationPackage = (SAPbouiCOM.MenuCreationParams)Program.uiApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                oCreationPackage.Checked = false;
                oCreationPackage.Enabled = true;
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "UDO_F_BDOSDEPACR_D";
                oCreationPackage.String = BDOSResources.getTranslate("DepreciationAccrual");
                oCreationPackage.Position = fatherMenuItem.SubMenus.Count - 1;

                menuItem = fatherMenuItem.SubMenus.AddEx(oCreationPackage);
            }
            catch
            {

            }
        }

        public static void uiApp_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent, out string errorText)
        {
            errorText = null;
            BubbleEvent = true;

            //----------------------------->Cancel <-----------------------------
            try
            {
                SAPbouiCOM.Form oDocForm = Program.uiApp.Forms.ActiveForm;


                if (pVal.BeforeAction == false && pVal.MenuUID == "BDOSAddRow")
                {
                    addMatrixRow(oDocForm, out errorText);
                }
                else if (pVal.BeforeAction == false && pVal.MenuUID == "BDOSDelRow")
                {
                    delMatrixRow(oDocForm, out errorText);
                }
            }
            catch (Exception ex)
            {
                Program.uiApp.MessageBox(ex.ToString(), 1, "", "");
            }
        }
        
        public static void uiApp_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD & BusinessObjectInfo.BeforeAction == true)
            {
                return;
            }

            SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(BusinessObjectInfo.FormTypeEx, Program.currentFormCount);

            if (oForm.TypeEx == "UDO_FT_UDO_F_BDOSDEPACR_D")
            {
                if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD)
                {
                    if (BusinessObjectInfo.BeforeAction == true)
                    {
                        //checkDoc(oForm, out errorText);
                        if (errorText != null)
                        {
                            Program.uiApp.StatusBar.SetSystemMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            Program.uiApp.MessageBox(BDOSResources.getTranslate("OperationUnsuccesfullSeeLog"));
                            BubbleEvent = false;
                        }
                        else
                        {
                            //updateAsset(oForm, false, out errorText);
                            if (errorText != null)
                            {
                                Program.uiApp.StatusBar.SetSystemMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                Program.uiApp.MessageBox(BDOSResources.getTranslate("OperationUnsuccesfullSeeLog"));
                                BubbleEvent = false;
                            }
                        }
                    }
                }
                if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE && BusinessObjectInfo.BeforeAction)
                {
                    int answer = 0;
                    answer = Program.uiApp.MessageBox(BDOSResources.getTranslate("DoYouReallyWantCancelDoc") + "?", 1, BDOSResources.getTranslate("Yes"), BDOSResources.getTranslate("No"), "");
                    if (answer == 1 && Program.cancellationTrans & Program.canceledDocEntry != 0)
                    {
                        cancellation(oForm, Program.canceledDocEntry, out errorText);
                        Program.canceledDocEntry = 0;
                    }
                    else
                    {
                        BubbleEvent = false;                      
                    }
                }

                if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD & BusinessObjectInfo.BeforeAction == false)
                {
                    formDataLoad(oForm, out errorText);
                }

            }
        }

        public static void cancellation(SAPbouiCOM.Form oForm, int docEntry, out string errorText)
        {
            errorText = null;

            try
            {
                JournalEntry.cancellation(oForm, docEntry, "UDO_F_BDOSDEPACR_D", out errorText);
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


        public static void uiApp_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);
                setVisibleFormItems(oForm, out errorText);

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD & pVal.BeforeAction == true)
                {
                    createFormItems(oForm, out BubbleEvent, out errorText);
                    Program.FORM_LOAD_FOR_VISIBLE = true;
                    Program.FORM_LOAD_FOR_ACTIVATE = true;
                }


                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_GOT_FOCUS && oForm.Visible == true && oForm.VisibleEx == true && openFormEvent == false)
                {
                   
                    string docEntry = oForm.DataSources.DBDataSources.Item("@BDOSDEPACR").GetValue("DocEntry", 0).Trim();
                    if (string.IsNullOrEmpty(docEntry))
                    {
                        addMatrixRow(oForm, out errorText);
                    }

                    openFormEvent = true;
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE & pVal.BeforeAction == false)
                {
                    oForm.Freeze(true);
                    if (Program.FORM_LOAD_FOR_VISIBLE == true)
                    {
                        setSizeForm(oForm, out errorText);
                        oForm.Title = BDOSResources.getTranslate("DepreciationAccrual");
                        Program.FORM_LOAD_FOR_VISIBLE = false;
                    }
                    oForm.Freeze(false);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_RESIZE & pVal.BeforeAction == false)
                {
                    oForm.Freeze(true);
                    resizeForm(oForm, out errorText);
                    oForm.Freeze(false);
                }
                

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE & pVal.BeforeAction == false)
                {
                    if (Program.FORM_LOAD_FOR_ACTIVATE == true)
                    {
                        oForm.Freeze(true);
                        formDataLoad(oForm, out errorText);
                        oForm.Freeze(false);
                        Program.FORM_LOAD_FOR_ACTIVATE = false;
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

        public static void createFormItems(SAPbouiCOM.Form oForm, out bool BubbleEvent, out string errorText)
        {
            BubbleEvent = true;
            errorText = null;

            Dictionary<string, object> formItems;
            string itemName = "";

            int left_s = 6;
            int left_e = 130;
            int height = 15;
            int top = 6;
            int width_s = 120;
            int width_e = 148;

            string objectTypeLocation = "144";
            string objectTypeItem = "4";
            string objectTypeDist = "10000044";

            
            top = top + height + 1;
            oForm.AutoManaged = true;
            

            formItems = new Dictionary<string, object>();
            itemName = "DocDateS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("PostingDate"));
            formItems.Add("LinkTo", "DocDateE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "DocDateE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSDEPACR");
            formItems.Add("Alias", "U_DocDate");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("SetAutoManaged", true);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }


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
                return;
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
            formItems.Add("DisplayDesc", true);
            formItems.Add("SetAutoManaged", true);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            

            
            left_s = 6;
            left_e = 127;
            top = top + 2 * height + 1;

            //საკონტროლო პანელი
            formItems = new Dictionary<string, object>();
            itemName = "FillMTR"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            formItems.Add("Left", left_s);
            formItems.Add("Width", 100);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Fill"));
            formItems.Add("SetAutoManaged", true);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            //ცხრილური ნაწილები
            left_s = 6;
            left_e = 127;
            top = top + 2 * height + 1;

            //მატრიცა
            formItems = new Dictionary<string, object>();
            itemName = "DepAcrMTR"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
            formItems.Add("Left", left_s);
            formItems.Add("Width", oForm.Width);
            formItems.Add("Top", top);
            formItems.Add("Height", 70);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("AffectsFormMode", false);
            formItems.Add("SetAutoManaged", true);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }
            top = top + 5+70;

            //შემქმნელი
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
                return;
            }

            top = top + 5;


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
            formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", false);
            formItems.Add("SetAutoManaged", true);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }
            
            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("DepAcrMTR").Specific));
            oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
            SAPbouiCOM.Columns oColumns = oMatrix.Columns;
            SAPbouiCOM.Column oColumn;

            oColumn = oColumns.Add("LineID", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "  ";
            oColumn.Width = 20;
            oColumn.Editable = false;
            oColumn.Visible = true;

            oColumn = oColumns.Add("ItemCode", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("ItemCode");
            oColumn.Width = 60;
            oColumn.Editable = true;
            oColumn.Visible = true;
            oColumn.ExtendedObject.LinkedObjectType = objectTypeItem;

            oColumn = oColumns.Add("DistNumber", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("DistNumber");
            oColumn.Width = 60;
            oColumn.Editable = true;
            oColumn.Visible = true;

            oColumn = oColumns.Add("BDOSUsLife", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("UseLife");
            oColumn.Width = 60;
            oColumn.Editable = false;
            oColumn.Visible = true;

            oColumn = oColumns.Add("Quantity", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Quantity");
            oColumn.Width = 60;
            oColumn.Editable = true;
            oColumn.Visible = true;

            oColumn = oColumns.Add("DeprAmt", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("DepreciationAmount");
            oColumn.Width = 60;
            oColumn.Editable = true;
            oColumn.Visible = true;

            oColumn = oColumns.Add("AlrDeprAmt", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("AlrDeprAmt");
            oColumn.Width = 60;
            oColumn.Editable = true;
            oColumn.Visible = true;

            oColumn = oColumns.Add("Project", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Project");
            oColumn.Width = 60;
            oColumn.Editable = true;
            oColumn.Visible = true;
            oColumn.ExtendedObject.LinkedObjectType = objectTypeLocation;
            
            oColumn = oColumns.Add("InvEntry", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("InvEntry");
            oColumn.Width = 60;
            oColumn.Editable = true;
            oColumn.Visible = true;

            oColumn = oColumns.Add("InvType", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("InvType");
            oColumn.Width = 60;
            oColumn.Editable = true;
            oColumn.Visible = true;



            SAPbouiCOM.DBDataSource oDBDataSource;
            oDBDataSource = oForm.DataSources.DBDataSources.Add("@BDOSDEPAC1");

            oColumn = oColumns.Item("LineID");
            oColumn.DataBind.SetBound(true, "@BDOSDEPAC1", "LineID");

            oColumn = oColumns.Item("ItemCode");
            oColumn.DataBind.SetBound(true, "@BDOSDEPAC1", "U_ItemCode");
                       
            oColumn = oColumns.Item("DistNumber");
            oColumn.DataBind.SetBound(true, "@BDOSDEPAC1", "U_DistNumber");

            oColumn = oColumns.Item("BDOSUsLife");
            oColumn.DataBind.SetBound(true, "@BDOSDEPAC1", "U_BDOSUsLife");

            oColumn = oColumns.Item("Quantity");
            oColumn.DataBind.SetBound(true, "@BDOSDEPAC1", "U_Quantity");

            oColumn = oColumns.Item("DeprAmt");
            oColumn.DataBind.SetBound(true, "@BDOSDEPAC1", "U_DeprAmt");

            oColumn = oColumns.Item("AlrDeprAmt");
            oColumn.DataBind.SetBound(true, "@BDOSDEPAC1", "U_AlrDeprAmt");
            
            oColumn = oColumns.Item("Project");
            oColumn.DataBind.SetBound(true, "@BDOSDEPAC1", "U_Project");

            oColumn = oColumns.Item("InvEntry");
            oColumn.DataBind.SetBound(true, "@BDOSDEPAC1", "U_InvEntry");

            oColumn = oColumns.Item("InvType");
            oColumn.DataBind.SetBound(true, "@BDOSDEPAC1", "U_InvType");

            //მარჯვენა რიგი
            top = 6;
            width_s = 120;
            left_s = 295;
            left_e = left_s + 121;

            formItems = new Dictionary<string, object>();
            itemName = "CanceledS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Canceled"));
            formItems.Add("LinkTo", "DocDate");
            formItems.Add("FromPane", 1);
            formItems.Add("ToPane", 2);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            Dictionary<string, string> StatusesList = new Dictionary<string, string>();
            StatusesList.Add("Y", BDOSResources.getTranslate("Canceled"));
            StatusesList.Add("N", BDOSResources.getTranslate("Active"));

            formItems = new Dictionary<string, object>();
            itemName = "CanceledE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSDEPACR");
            formItems.Add("Alias", "Canceled");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Enabled", false);
            formItems.Add("ValidValues", StatusesList);
            formItems.Add("DisplayDesc", true);
            formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
            formItems.Add("SetAutoManaged", true);
            formItems.Add("FromPane", 1);
            formItems.Add("ToPane", 2);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "BDOSJrnEnS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("TransactionNo"));
            formItems.Add("LinkTo", "BDOSJrnEnt");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "BDOSJrnEnt";
            formItems.Add("Size", 20);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Enabled", false);
            formItems.Add("AffectsFormMode", false);
            formItems.Add("SetAutoManaged", true);


            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "BDOSJEntLB";
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            formItems.Add("Left", left_e - 20);
            formItems.Add("Top", top);
            formItems.Add("UID", itemName);
            formItems.Add("LinkTo", "BDOSJrnEnt");
            formItems.Add("LinkedObjectType", "30");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }


            GC.Collect();
        }

        public static void setVisibleFormItems(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;


            try
            {
                string docEntry = oForm.DataSources.DBDataSources.Item("@BDOSDEPACR").GetValue("DocEntry", 0).Trim();
                bool docEntryIsEmpty = string.IsNullOrEmpty(docEntry);
                
                oForm.Items.Item("BDOSJrnEnt").Enabled = (docEntryIsEmpty == true);
                
                //oForm.Update();
                //oForm.Refresh();
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

        public static void resizeForm(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            try
            {
                reArrangeFormItems(oForm);
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
        }

        public static void reArrangeFormItems(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.Item oItem = null;
            int height = 15;
            int top = 6;
            
            oForm.Items.Item("CanceledS").Top = top;
            oForm.Items.Item("CanceledE").Top = top;

            top = top + height + 1;
            oForm.Items.Item("BDOSJrnEnS").Top = top;
            oForm.Items.Item("BDOSJrnEnt").Top = top;
            oForm.Items.Item("BDOSJEntLB").Top = top;
            
            top = height + 7;
            oForm.Items.Item("DocDateS").Top = top;
            oForm.Items.Item("DocDateE").Top = top;

            top = top + height + 1;
            oForm.Items.Item("AccrMnthS").Top = top;
            oForm.Items.Item("AccrMnthE").Top = top;

            
            top =top + 2 * height + 1;

            top = top + height + 1;
            oForm.Items.Item("FillMTR").Top = top;

            int MTRWidth = oForm.Width - 15;
            top = top + height + 1;
            oItem = oForm.Items.Item("DepAcrMTR");
            oItem.Top = top;
            oItem.Width = MTRWidth;
            oItem.Height = oForm.Height - 220;

            // სვეტების ზომები 
            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("DepAcrMTR").Specific));
            SAPbouiCOM.Columns oColumns = oMatrix.Columns;
            SAPbouiCOM.Column oColumn;
            oColumn = oMatrix.Columns.Item("LineID");
            oColumn.Width = 20 - 1;
            MTRWidth = MTRWidth - 20 - 1;

           
            top = top + 20 * height;

            oForm.Items.Item("CreatorS").Top = oForm.Height - 110;
            oForm.Items.Item("CreatorE").Top = oForm.Height - 110;

            oForm.Items.Item("1").Top = oForm.Height - 80;
            oForm.Items.Item("2").Top = oForm.Height - 80;

        }

        public static void setSizeForm(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            try
            {
                oForm.ClientHeight = Program.uiApp.Desktop.Width / 3;
                oForm.Height = Program.uiApp.Desktop.Width / 2;

                oForm.Left = (Program.uiApp.Desktop.Width - oForm.Width) / 2;
                oForm.Top = (Program.uiApp.Desktop.Height - oForm.Height) / 2;
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
        }

        public static void addMatrixRow(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            oForm.Freeze(true);

            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("DepAcrMTR").Specific));
            SAPbouiCOM.DBDataSource mtrDataSource = oForm.DataSources.DBDataSources.Item("@BDOSDEPAC1");

            oMatrix.FlushToDataSource();
            if (mtrDataSource.GetValue("U_ItemCode", mtrDataSource.Size - 1) != "")
            {
                mtrDataSource.InsertRecord(mtrDataSource.Size);
            }
            mtrDataSource.SetValue("LineId", mtrDataSource.Size - 1, mtrDataSource.Size.ToString());

            oMatrix.LoadFromDataSource();

            //SAPbouiCOM.Column oColumn = oMatrix.Columns.Item("LineID");
            //oColumn.Editable = false;

            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
            {
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
            }

            oForm.Freeze(false);
        }

        public static void delMatrixRow(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            oForm.Freeze(true);

            try
            {
                SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("DepAcrMTR").Specific));
                SAPbouiCOM.DBDataSource mtrDataSource = oForm.DataSources.DBDataSources.Item("@BDOSDEPAC1");

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
                        mtrDataSource.RemoveRecord(row - deletedRowCount);
                        firstRow = row;
                    }
                }

                for (int i = 0; i <= mtrDataSource.Size; i++)
                {
                    mtrDataSource.SetValue("LineId", i, (i + 1).ToString());
                }

                oMatrix.LoadFromDataSource();

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE && deletedRowCount > 0)
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
            finally
            {
                GC.Collect();
                oForm.Freeze(false);
            }
        }
        public static DataTable createAdditionalEntries(SAPbouiCOM.Form oForm, SAPbobsCOM.GeneralData oGeneralData, int docEntry, DateTime DocDate, string EmployeeCode, DataTable WTaxDefinitons, DataTable AccountTable)
        {
            DataTable jeLines = JournalEntry.JournalEntryTable();
            SAPbobsCOM.GeneralDataCollection oChild = null;
            SAPbouiCOM.DBDataSource DBDataSourceTable = null;
            if (AccountTable == null)
            {
                AccountTable = CommonFunctions.GetOACTTable();
            }


            int JEcount = 0;

            string errorText = null;

            if (oForm == null)
            {
                oChild = oGeneralData.Child("BDOSDEPAC1");
                JEcount = oChild.Count;
            }
            else
            {
                DBDataSourceTable = oForm.DataSources.DBDataSources.Item("@BDOSDEPAC1");
                JEcount = DBDataSourceTable.Size;
            }




            for (int i = 0; i < JEcount; i++)
            {

                decimal DeprAmt = Convert.ToDecimal(CommonFunctions.getChildOrDbDataSourceValue(DBDataSourceTable, oChild, null, "U_DeprAmt", i), CultureInfo.InvariantCulture);
                decimal AlrDeprAmt = Convert.ToDecimal(CommonFunctions.getChildOrDbDataSourceValue(DBDataSourceTable, oChild, null, "U_AlrDeprAmt", i), CultureInfo.InvariantCulture);
                
                string ItemCode = ((string)CommonFunctions.getChildOrDbDataSourceValue(DBDataSourceTable, oChild, null, "U_ItemCode", i)).Trim();
                string U_PrjCode = ((string)CommonFunctions.getChildOrDbDataSourceValue(DBDataSourceTable, oChild, null, "U_Project", i)).Trim();
                string U_InvEntry = ((string)CommonFunctions.getChildOrDbDataSourceValue(DBDataSourceTable, oChild, null, "U_InvEntry", i)).Trim();
                string U_InvType = ((string)CommonFunctions.getChildOrDbDataSourceValue(DBDataSourceTable, oChild, null, "U_InvType", i)).Trim();

                SAPbobsCOM.Items oItem;
                oItem = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems);
                oItem.GetByKey(ItemCode);

                SAPbobsCOM.ItemGroups oItemGroup;
                oItemGroup = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItemGroups);
                oItemGroup.GetByKey(oItem.ItemsGroupCode);

                string AccDepAccount = oItemGroup.UserFields.Fields.Item("U_BDOSAccDep").Value.ToString();
                string ExpDepAccount = oItemGroup.UserFields.Fields.Item("U_BDOSExpDep").Value.ToString();
                string SaleCostAc = oItemGroup.CostAccount;
                JournalEntry.AddJournalEntryRow(AccountTable, jeLines, "Full", ExpDepAccount, AccDepAccount, DeprAmt, 0, "", "", "", "", "", "", U_PrjCode, "", "");

                if ( string.IsNullOrEmpty(U_InvEntry)==false)
                {
                    JournalEntry.AddJournalEntryRow(AccountTable, jeLines, "Full", AccDepAccount, SaleCostAc, AlrDeprAmt, 0, "", "", "", "", "", "", U_PrjCode, "", "");
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

        public static void formDataLoad(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            try
            {
                openFormEvent = false;

                string DocEntry = oForm.DataSources.DBDataSources.Item("@BDOSDEPACR").GetValue("DocEntry", 0).Trim();

                setVisibleFormItems(oForm, out errorText);

                // გატარებები
                SAPbouiCOM.DBDataSource DocDBSourceDepAcr = oForm.DataSources.DBDataSources.Item(0);
                string Ref1 = DocDBSourceDepAcr.GetValue("DocEntry", 0);
                string Ref2 = "UDO_F_BDOSDEPACR_D";

                string strdate = DocDBSourceDepAcr.GetValue("U_DocDate", 0);

                DateTime DocDate = DateTime.TryParseExact(strdate, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DocDate) == false ? DateTime.Now : DocDate;

                SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string query = "SELECT " +
                                "*  " +
                                "FROM \"OJDT\"  " +
                                "WHERE \"Ref1\" = '" + Ref1 + "' " +
                                "AND \"Ref2\" = '" + Ref2 + "' " +
                                "AND \"RefDate\" = '" + DocDate.ToString("yyyyMMdd") + "' ";
                oRecordSet.DoQuery(query);

                if (!oRecordSet.EoF)
                {
                    oForm.Items.Item("BDOSJrnEnt").Specific.Value = oRecordSet.Fields.Item("TransId").Value;
                }
                else
                {
                    oForm.Items.Item("BDOSJrnEnt").Specific.Value = "";
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

        public static decimal getDepreciationPriceDistNumber(string ItemCode, string DistNumber)
        {

            string query = @"select
	                         ""@BDOSDEPAC1"".""U_ItemCode"",
	                         ""@BDOSDEPAC1"".""U_DistNumber"",
	                         SUM(""@BDOSDEPAC1"".""U_DeprAmt"") as ""U_DeprAmt"",
	                         SUM(""@BDOSDEPAC1"".""U_Quantity"") as ""U_Quantity"" 
                        from ""@BDOSDEPAC1"" 
                        inner join ""@BDOSDEPACR"" on ""@BDOSDEPACR"".""DocEntry"" = ""@BDOSDEPAC1"".""DocEntry"" 
                        and ""@BDOSDEPACR"".""Canceled"" = 'N' and ""@BDOSDEPAC1"".""U_ItemCode"" = '" + ItemCode + @"' and ""@BDOSDEPAC1"".""U_DistNumber"" = '" + DistNumber + @"'
                        group by ""@BDOSDEPAC1"".""U_ItemCode"",
	                         ""@BDOSDEPAC1"".""U_DistNumber""";

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRecordSet.DoQuery(query);

            if (!oRecordSet.EoF)
            {
               
                decimal DeprQty = Convert.ToDecimal(oRecordSet.Fields.Item("U_Quantity").Value);
                decimal AlrDeprAmt = 0;
                if (DeprQty > 0)
                {
                    AlrDeprAmt = Convert.ToDecimal(oRecordSet.Fields.Item("U_DeprAmt").Value) / DeprQty;
                }

                return AlrDeprAmt;
            }

            return 0;
        }

    }
}


