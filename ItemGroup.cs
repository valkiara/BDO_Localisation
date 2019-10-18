using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BDO_Localisation_AddOn
{
    class ItemGroup
    {
        public static void createUserFields(out string errorText)
        {
            errorText = null;
            Dictionary<string, object> fieldskeysMap;

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSFxAs");
            fieldskeysMap.Add("TableName", "OITB");
            fieldskeysMap.Add("Description", "Fixed asset");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 1);
            fieldskeysMap.Add("DefaultValue", "N");

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSAccDep");
            fieldskeysMap.Add("TableName", "OITB");
            fieldskeysMap.Add("Description", "Accumulative depreciation account");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 150);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSExpDep");
            fieldskeysMap.Add("TableName", "OITB");
            fieldskeysMap.Add("Description", "Depreciation expence account");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 150);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSUsLife");
            fieldskeysMap.Add("TableName", "OITB");
            fieldskeysMap.Add("Description", "Useful life");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Numeric);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSFuel");
            fieldskeysMap.Add("TableName", "OITB");
            fieldskeysMap.Add("Description", "Fuel");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 1);
            fieldskeysMap.Add("DefaultValue", "N");

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            GC.Collect();
        }

        public static void uiApp_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                {
                    SAPbouiCOM.IChooseFromListEvent oCFLEvento = null;
                    oCFLEvento = ((SAPbouiCOM.IChooseFromListEvent)(pVal));

                    chooseFromListAddForm(oForm, oCFLEvento, pVal.ItemUID, pVal.BeforeAction, out errorText);
                    chooseFromList(oForm, pVal.BeforeAction, oCFLEvento, out errorText);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD & pVal.BeforeAction == true)
                {
                    createFormItems(oForm, out errorText);
                }
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_RESIZE & pVal.BeforeAction == false)
                {
                    oForm.Freeze(true);
                    resizeForm(oForm, out errorText);
                    oForm.Freeze(false);
                }

                if (pVal.ItemUID == "BDOSFxAs" || pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
                {
                    setVisibility(oForm, out errorText);
                }             
            }
        }

        public static void uiApp_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(BusinessObjectInfo.FormTypeEx, Program.currentFormCount);

            if (oForm.TypeEx == "63")
            {
                if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD & BusinessObjectInfo.BeforeAction == false)
                {
                    setVisibility(oForm, out errorText);
                }
            }
        }

        public static void reArrangeFormItems(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.Item oItem = null;
            int height = 15;

            int top = oForm.Items.Item("123").Top;
            top = top + height + 1;
            
            oItem = oForm.Items.Item("BDOSFxAs");
            oItem.FromPane = 1;
            oItem.ToPane = 1;
            top = top + height + 1;
            oItem.Top = top;

            oItem = oForm.Items.Item("BDOSFuel");
            oItem.FromPane = 1;
            oItem.ToPane = 1;
            oItem.Top = top;

            oItem = oForm.Items.Item("AccDepS");
            oItem.FromPane = 1;
            oItem.ToPane = 1; top = top + height + 1;
            oItem.Top = top;

            oItem = oForm.Items.Item("AccDep");
            oItem.FromPane = 1;
            oItem.ToPane = 1;
            oItem.Top = top;
           
            oItem = oForm.Items.Item("ExpDepS");
            oItem.FromPane = 1;
            oItem.ToPane = 1;
            top = top + height + 1;
            oItem.Top = top;

            oItem = oForm.Items.Item("ExpDep");
            oItem.FromPane = 1;
            oItem.ToPane = 1;
            oItem.Top = top;

            oItem = oForm.Items.Item("UsLifeS");
            oItem.FromPane = 1;
            oItem.ToPane = 1;
            top = top + height + 1;
            oItem.Top = top;

            oItem = oForm.Items.Item("UsLife");
            oItem.FromPane = 1;
            oItem.ToPane = 1;
            oItem.Top = top;        
        }

        public static void setVisibility(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            oForm.Freeze(true);
            bool visibleProperties = oForm.DataSources.DBDataSources.Item("OITB").GetValue("U_BDOSFxAs", 0) == "Y";

            try
            {
                oForm.Items.Item("AccDep").Visible = visibleProperties;
                oForm.Items.Item("AccDep").FromPane = 1;
                oForm.Items.Item("AccDep").ToPane = 1;
                oForm.Items.Item("AccDepS").Visible = visibleProperties;
                oForm.Items.Item("AccDepS").FromPane = 1;
                oForm.Items.Item("AccDepS").ToPane = 1;
                oForm.Items.Item("ExpDep").Visible = visibleProperties;
                oForm.Items.Item("ExpDep").FromPane = 1;
                oForm.Items.Item("ExpDep").ToPane = 1;
                oForm.Items.Item("ExpDepS").Visible = visibleProperties;
                oForm.Items.Item("ExpDepS").FromPane = 1;
                oForm.Items.Item("ExpDepS").ToPane = 1;
                oForm.Items.Item("UsLife").Visible = visibleProperties;
                oForm.Items.Item("UsLife").FromPane = 1;
                oForm.Items.Item("UsLife").ToPane = 1;
                oForm.Items.Item("UsLifeS").Visible = visibleProperties;
                oForm.Items.Item("UsLifeS").FromPane = 1;
                oForm.Items.Item("UsLifeS").ToPane = 1;                
            }
            catch
            { }

            oForm.Freeze(false);
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

        public static void createFormItems(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            Dictionary<string, object> formItems;
            string itemName = "";

            int Left = oForm.Items.Item("123").Left;
            int Top = oForm.Items.Item("123").Left + 20;
            int Height = oForm.Items.Item("123").Height;
            int Width = oForm.Items.Item("123").Width;
            int pane = 1;

            formItems = new Dictionary<string, object>();
            itemName = "BDOSFxAs";
            formItems.Add("isDataSource", true);
            formItems.Add("Length", 1);
            formItems.Add("DataSource", "UserDataSources");
            formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
            formItems.Add("TableName", "OITB");
            formItems.Add("Alias", "U_"+itemName);
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            formItems.Add("Width", Width);
            formItems.Add("Left", Left);
            formItems.Add("Top", Top);
            formItems.Add("Caption", BDOSResources.getTranslate("FixedAsset"));
            formItems.Add("AffectsFormMode", true);
            formItems.Add("UID", itemName);
            formItems.Add("Enabled", true);
            formItems.Add("ValueOn", "Y");
            formItems.Add("ValueOff", "N");
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "BDOSFuel"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OITB");
            formItems.Add("Alias", "U_BDOSFuel");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            formItems.Add("Width", Width);
            formItems.Add("Left", oForm.Items.Item("124").Left);
            formItems.Add("Top", Top);
            formItems.Add("Height", Height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Fuel"));
            formItems.Add("ValOff", "N");
            formItems.Add("ValOn", "Y");
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            Top = Top + Height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "AccDepS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", Left);
            formItems.Add("Width", Width);
            formItems.Add("Top", Top);
            formItems.Add("Height", Height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("AccumulatedDepreciationAccount"));
            formItems.Add("LinkTo", "AccDep");
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);
            formItems.Add("Visible", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            Left = oForm.Items.Item("124").Left;

            string objectType = "1"; //SAPbouiCOM.BoLinkedObject.lf_GLAccounts, Business Partner object 
            bool multiSelection = false;
            string uniqueID_lf_GLAccCFL_D = "acc_CFL_D";
            FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_lf_GLAccCFL_D);

            formItems = new Dictionary<string, object>();
            itemName = "AccDep"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OITB");
            formItems.Add("Alias", "U_BDOSAccDep");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", Left);
            formItems.Add("Width", Width);
            formItems.Add("Top", Top);
            formItems.Add("Height", Height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("ChooseFromListUID", uniqueID_lf_GLAccCFL_D);
            formItems.Add("ChooseFromListAlias", "AcctCode");
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);
            formItems.Add("Visible", false);
            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            Top = Top + Height + 1;
            Left = oForm.Items.Item("123").Left;

            formItems = new Dictionary<string, object>();
            itemName = "ExpDepS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", Left);
            formItems.Add("Width", Width);
            formItems.Add("Top", Top);
            formItems.Add("Height", Height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("DepreciationExpenseAccount"));
            formItems.Add("LinkTo", "ExpDep");
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);
            formItems.Add("Visible", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            Left = oForm.Items.Item("124").Left;

            objectType = "1"; //SAPbouiCOM.BoLinkedObject.lf_GLAccounts, Business Partner object 
            multiSelection = false;
            uniqueID_lf_GLAccCFL_D = "Exp_CFL_D";
            FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_lf_GLAccCFL_D);

            formItems = new Dictionary<string, object>();
            itemName = "ExpDep"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OITB");
            formItems.Add("Alias", "U_BDOSExpDep");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", Left);
            formItems.Add("Width", Width);
            formItems.Add("Top", Top);
            formItems.Add("Height", Height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("ChooseFromListUID", uniqueID_lf_GLAccCFL_D);
            formItems.Add("ChooseFromListAlias", "AcctCode");
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);
            formItems.Add("Visible", false);
            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            Top = Top + Height + 1;
            Left = oForm.Items.Item("123").Left;

            formItems = new Dictionary<string, object>();
            itemName = "UsLifeS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", Left);
            formItems.Add("Width", Width);
            formItems.Add("Top", Top);
            formItems.Add("Height", Height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("UsLife"));
            formItems.Add("LinkTo", "UsLife");
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);
            formItems.Add("Visible", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            Left = oForm.Items.Item("124").Left;
            
            formItems = new Dictionary<string, object>();
            itemName = "UsLife"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OITB");
            formItems.Add("Alias", "U_BDOSUsLife");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", Left);
            formItems.Add("Width", Width);
            formItems.Add("Top", Top);
            formItems.Add("Height", Height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);
            formItems.Add("Visible", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }
        }

        private static void chooseFromList(SAPbouiCOM.Form oForm, bool BeforeAction, SAPbouiCOM.IChooseFromListEvent oCFLEvento, out string errorText)
        {
            errorText = null;

            string sCFL_ID = oCFLEvento.ChooseFromListUID;
            SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(sCFL_ID);

            SAPbouiCOM.DataTable oDataTable = null;
            oDataTable = oCFLEvento.SelectedObjects;
            try
            {
                if (BeforeAction == false)
                {
                    if (oDataTable != null)
                    {
                        if (sCFL_ID == "acc_CFL_D")
                        {
                            string AcctCode = Convert.ToString(oDataTable.GetValue("AcctCode", 0));

                            SAPbouiCOM.EditText oAccDep = oForm.Items.Item("AccDep").Specific;
                            oAccDep.Value = AcctCode;
                        }

                        else if (sCFL_ID == "Exp_CFL_D")
                        {
                            string AcctCode = Convert.ToString(oDataTable.GetValue("AcctCode", 0));

                            SAPbouiCOM.EditText oExpDep = oForm.Items.Item("ExpDep").Specific;
                            oExpDep.Value = AcctCode;
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
        }
        public static void chooseFromListAddForm(SAPbouiCOM.Form oForm, SAPbouiCOM.IChooseFromListEvent oCFLEvento, string itemUID, bool beforeAction, out string errorText)
        {
            errorText = null;

            try
            {
                string sCFL_ID = oCFLEvento.ChooseFromListUID;
                SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(sCFL_ID);

                if (beforeAction == true)
                {
                    if (sCFL_ID == "acc_CFL_D")
                    {
                        oCFL = oForm.ChooseFromLists.Item("acc_CFL_D");

                        SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();
                        SAPbouiCOM.Condition oCon = oCons.Add();
                        oCon.Alias = "Postable";
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = "Y";

                        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_NONE;
                        

                        oCFL.SetConditions(oCons);
                    }
                    else if (sCFL_ID == "Exp_CFL_D")
                    {
                        oCFL = oForm.ChooseFromLists.Item("Exp_CFL_D");

                        SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();
                        SAPbouiCOM.Condition oCon = oCons.Add();
                        oCon.Alias = "Postable";
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = "Y";

                        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_NONE;


                        oCFL.SetConditions(oCons);
                    }

                }
                
            }
            catch (Exception ex)
            {
                string exsd = ex.Message;
            }
        }
    }
}
