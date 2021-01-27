using System;
using System.Collections.Generic;

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

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                {
                    SAPbouiCOM.IChooseFromListEvent oCFLEvento = null;
                    oCFLEvento = ((SAPbouiCOM.IChooseFromListEvent)(pVal));

                    chooseFromListAddForm(oForm, oCFLEvento, pVal.ItemUID, pVal.BeforeAction);
                    chooseFromList(oForm, pVal.BeforeAction, oCFLEvento);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD && pVal.BeforeAction)
                {
                    createFormItems(oForm);
                    Program.FORM_LOAD_FOR_VISIBLE = true;
                    Program.FORM_LOAD_FOR_ACTIVATE = true;
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE && !pVal.BeforeAction)
                {
                    if (Program.FORM_LOAD_FOR_ACTIVATE)
                    {
                        Program.FORM_LOAD_FOR_ACTIVATE = false;
                        setVisibleFormItems(oForm);
                    }
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_RESIZE & pVal.BeforeAction == false)
                {
                    resizeForm(oForm);
                }

                if (pVal.ItemUID == "BDOSFxAs" || pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
                {
                    setVisibleFormItems(oForm);
                }

                if (pVal.ItemUID == "7" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && !pVal.BeforeAction)
                {
                    setVisibleFormItems(oForm);
                }
            }
        }

        public static void uiApp_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;

            SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(BusinessObjectInfo.FormTypeEx, Program.currentFormCount);

            if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD && !BusinessObjectInfo.BeforeAction)
            {
                setVisibleFormItems(oForm);
            }
        }

        public static void setVisibleFormItems(SAPbouiCOM.Form oForm)
        {
            oForm.Freeze(true);
            try
            {
                bool visibleProperties = oForm.DataSources.DBDataSources.Item("OITB").GetValue("U_BDOSFxAs", 0) == "Y";
                int fromPane = 1;
                int toPane = 1;

                List<string> itemNames = new List<string>
                {
                    "AccDep",
                    "AccDepS",
                    "AccDepLB",
                    "ExpDep",
                    "ExpDepS",
                    "ExpDepLB",
                    "UsLife",
                    "UsLifeS"
                };

                itemNames.ForEach(itemName => setItemVisibility(oForm, itemName, visibleProperties, fromPane, toPane));

                string enableFuelMng = (string)CommonFunctions.getOADM("U_BDOSEnbFlM");
                if (enableFuelMng == "Y" && oForm.PaneLevel == 1)
                {
                    oForm.Items.Item("BDOSFuel").Visible = true;
                }
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

        public static void resizeForm(SAPbouiCOM.Form oForm)
        {
            oForm.Freeze(true);
            try
            {
                int height = 15;
                int fromPane = 1;
                int toPane = 1;
                int top = oForm.Items.Item("123").Top + height + 1;

                rearrangeItem(oForm, "BDOSFxAs", fromPane, toPane, top);
                rearrangeItem(oForm, "BDOSFuel", fromPane, toPane, top);

                top += height + 1;
                rearrangeItem(oForm, "AccDepS", fromPane, toPane, top);
                rearrangeItem(oForm, "AccDep", fromPane, toPane, top);
                rearrangeItem(oForm, "AccDepLB", fromPane, toPane, top);

                top += height + 1;
                rearrangeItem(oForm, "ExpDepS", fromPane, toPane, top);
                rearrangeItem(oForm, "ExpDep", fromPane, toPane, top);
                rearrangeItem(oForm, "ExpDepLB", fromPane, toPane, top);

                top += height + 1;
                rearrangeItem(oForm, "UsLifeS", fromPane, toPane, top);
                rearrangeItem(oForm, "UsLife", fromPane, toPane, top);
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

        public static void createFormItems(SAPbouiCOM.Form oForm)
        {
            string errorText;

            Dictionary<string, object> formItems;
            string itemName;

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
            formItems.Add("Alias", "U_" + itemName);
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
                throw new Exception(errorText);
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
            formItems.Add("Visible", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
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
                throw new Exception(errorText);
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
                throw new Exception(errorText);
            }

            formItems = new Dictionary<string, object>();
            itemName = "AccDepLB"; // Golden arrow
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            formItems.Add("Left", Left - 20);
            formItems.Add("Top", Top);
            formItems.Add("Height", Height);
            formItems.Add("UID", itemName);
            formItems.Add("LinkTo", "AccDep");
            formItems.Add("LinkedObjectType", objectType);

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
                throw new Exception(errorText);
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
                throw new Exception(errorText);
            }

            formItems = new Dictionary<string, object>();
            itemName = "ExpDepLB"; // Golden arrow
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            formItems.Add("Left", Left - 20);
            formItems.Add("Top", Top);
            formItems.Add("Height", Height);
            formItems.Add("UID", itemName);
            formItems.Add("LinkTo", "ExpDep");
            formItems.Add("LinkedObjectType", objectType);

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
            formItems.Add("Caption", BDOSResources.getTranslate("UsefulLife"));
            formItems.Add("LinkTo", "UsLife");
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);
            formItems.Add("Visible", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
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
                throw new Exception(errorText);
            }
        }

        private static void chooseFromList(SAPbouiCOM.Form oForm, bool BeforeAction, SAPbouiCOM.IChooseFromListEvent oCFLEvento)
        {
            oForm.Freeze(true);
            try
            {
                string sCFL_ID = oCFLEvento.ChooseFromListUID;
                SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(sCFL_ID);
                SAPbouiCOM.DataTable oDataTable = oCFLEvento.SelectedObjects;

                if (!BeforeAction)
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
                throw new Exception(ex.Message);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        public static void chooseFromListAddForm(SAPbouiCOM.Form oForm, SAPbouiCOM.IChooseFromListEvent oCFLEvento, string itemUID, bool beforeAction)
        {
            oForm.Freeze(true);

            try
            {
                string sCFL_ID = oCFLEvento.ChooseFromListUID;
                SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(sCFL_ID);

                if (beforeAction)
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
                throw new Exception(ex.Message);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        private static void setItemVisibility(SAPbouiCOM.Form oForm, string itemName, bool visibleProperties, int fromPane, int toPane)
        {
            SAPbouiCOM.Item oItem = oForm.Items.Item(itemName);

            oItem.Visible = visibleProperties;
            oItem.FromPane = fromPane;
            oItem.ToPane = toPane;
        }

        private static void rearrangeItem(SAPbouiCOM.Form oForm, string itemName, int fromPane, int toPane, int top)
        {
            SAPbouiCOM.Item oItem = oForm.Items.Item(itemName);

            oItem.FromPane = fromPane;
            oItem.ToPane = toPane;
            oItem.Top = top;
        }
    }
}