using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;
using System.Data;

namespace BDO_Localisation_AddOn
{
    class BDOSStockTransferWizard
    {
        static SAPbouiCOM.Form oFormWizard;
        public static string currentLineIDWhsTable;
        static SAPbouiCOM.Form oFormDetailWizard;
        //static SAPbouiCOM.Form oFormDetailWizard2;

        public static DataTable TableWhsItemsForDetail;

        public static void addMenus(out string errorText)
        {
            errorText = null;

            SAPbouiCOM.MenuItem menuItem;
            SAPbouiCOM.MenuItem fatherMenuItem;
            SAPbouiCOM.MenuCreationParams oCreationPackage;

            try
            {
                fatherMenuItem = Program.uiApp.Menus.Item("43540");

                // Add a pop-up menu item
                oCreationPackage = (SAPbouiCOM.MenuCreationParams)Program.uiApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                oCreationPackage.Checked = false;
                oCreationPackage.Enabled = true;
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "BDOSSTTRWZ";
                oCreationPackage.String = BDOSResources.getTranslate("StockTransferWizard");
                oCreationPackage.Position = fatherMenuItem.SubMenus.Count - 1;

                menuItem = fatherMenuItem.SubMenus.AddEx(oCreationPackage);
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
        }

        public static void createForm(out string errorText)
        {
            errorText = null;

            //ფორმის აუცილებელი თვისებები
            Dictionary<string, object> formProperties = new Dictionary<string, object>();
            formProperties.Add("UniqueID", "BDOSSTTRWZ");
            formProperties.Add("BorderStyle", SAPbouiCOM.BoFormBorderStyle.fbs_Sizable);
            formProperties.Add("Title", BDOSResources.getTranslate("StockTransferWizard"));
            formProperties.Add("Left", 558);
            formProperties.Add("ClientWidth", 750);
            formProperties.Add("Top", 335);
            formProperties.Add("ClientHeight", 400);

            SAPbouiCOM.Form oForm;
            bool newForm;
            bool formExist = FormsB1.createForm(formProperties, out oForm, out newForm, out errorText);

            if (errorText != null)
            {
                return;
            }

            if (formExist == true)
            {
                if (newForm == true)
                {
                    Dictionary<string, object> formItems;
                    string itemName = "";

                    int height = 15;
                    int top = 6;
                    int width_s = 100;
                    int width_e = 200;
                    int pane = 1;
                    int left = 6;

                    oForm.PaneLevel = pane;

                    formItems = new Dictionary<string, object>();
                    itemName = "2"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", oForm.Width - 215);
                    formItems.Add("Width", 65);
                    formItems.Add("Top", oForm.ClientHeight - 50);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("Cancel"));
                    formItems.Add("FromPane", 0);
                    formItems.Add("ToPane", 2);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "Prev"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", oForm.Width - 145);
                    formItems.Add("Width", 65);
                    formItems.Add("Top", oForm.ClientHeight - 50);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("Back"));
                    formItems.Add("FromPane", 0);
                    formItems.Add("ToPane", 2);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "Next"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", oForm.Width - 75);
                    formItems.Add("Width", 65);
                    formItems.Add("Top", oForm.ClientHeight - 50);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("Next"));
                    formItems.Add("FromPane", 0);
                    formItems.Add("ToPane", 2);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "DateS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("Date"));
                    formItems.Add("LinkTo", "DateE");
                    formItems.Add("FromPane", pane);
                    formItems.Add("ToPane", pane);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "DateE"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_DATE);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("Length", 15);
                    formItems.Add("Size", 15);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Left", left + 5 + width_s);
                    formItems.Add("Width", width_e);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("Enabled", true);
                    formItems.Add("FromPane", pane);
                    formItems.Add("ToPane", pane);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    top = top + height + 5;

                    formItems = new Dictionary<string, object>();
                    itemName = "DocTypeS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("DocType"));
                    formItems.Add("LinkTo", "DocTypeE");
                    formItems.Add("FromPane", pane);
                    formItems.Add("ToPane", pane);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    Dictionary<string, string> listValidValuesDict = new Dictionary<string, string>();
                    listValidValuesDict.Add("TransferToProjectWarehouse", BDOSResources.getTranslate("TransferToProjectWarehouse"));
                    listValidValuesDict.Add("TransferToWriteOffWarehouse", BDOSResources.getTranslate("TransferToWriteOffWarehouse"));
                    listValidValuesDict.Add("ReturnToProjectWarehouse", BDOSResources.getTranslate("ReturnToProjectWarehouse"));
                    listValidValuesDict.Add("ReturnToMainWarehouse", BDOSResources.getTranslate("ReturnToMainWarehouse"));
                    listValidValuesDict.Add("TransferWithoutType", BDOSResources.getTranslate("TransferWithoutType"));

                    formItems = new Dictionary<string, object>();
                    itemName = "DocTypeE";
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                    formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("Left", left + 5 + width_s);
                    formItems.Add("Width", width_e);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("ValidValues", listValidValuesDict);
                    formItems.Add("FromPane", pane);
                    formItems.Add("ToPane", pane);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    top = top + height + 5;
                    formItems = new Dictionary<string, object>();
                    itemName = "ItmGrpS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("ItemGroup"));
                    formItems.Add("LinkTo", "ItmGrpE");
                    formItems.Add("FromPane", pane);
                    formItems.Add("ToPane", pane);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    Dictionary<string, string> listValidValuesItemGroups = getItemGroupsList(out errorText);

                    formItems = new Dictionary<string, object>();
                    itemName = "ItmGrpE";
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                    formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("Left", left + 5 + width_s);
                    formItems.Add("Width", width_e);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("ValidValues", listValidValuesItemGroups);
                    formItems.Add("FromPane", pane);
                    formItems.Add("ToPane", pane);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    top = top + height + 5;

                    formItems = new Dictionary<string, object>();
                    itemName = "CatLvlS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("CategoryLevel"));
                    formItems.Add("LinkTo", "CatLvlE");
                    formItems.Add("FromPane", pane);
                    formItems.Add("ToPane", pane);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    Dictionary<string, string> listValidValuesLvl = new Dictionary<string, string>();
                    //listValidValuesLvl.Add("0", "0");
                    listValidValuesLvl.Add("1", "1");
                    listValidValuesLvl.Add("2", "2");
                    listValidValuesLvl.Add("3", "3");
                    listValidValuesLvl.Add("4", "4");
                    listValidValuesLvl.Add("5", "5");
                    listValidValuesLvl.Add("6", "6");
                    listValidValuesLvl.Add("7", "7");
                    listValidValuesLvl.Add("8", "8");
                    listValidValuesLvl.Add("9", "9");
                    listValidValuesLvl.Add("10", "10");

                    formItems = new Dictionary<string, object>();
                    itemName = "CatLvlE";
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                    formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("Left", left + 5 + width_s);
                    formItems.Add("Width", width_e / 3);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("ValidValues", listValidValuesLvl);
                    formItems.Add("FromPane", pane);
                    formItems.Add("ToPane", pane);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    SAPbouiCOM.ComboBox oComboBoxGrp = (SAPbouiCOM.ComboBox)oForm.Items.Item("CatLvlE").Specific;
                    oComboBoxGrp.Select("1", SAPbouiCOM.BoSearchKey.psk_ByValue);

                    top = top + height + 5;

                    formItems = new Dictionary<string, object>();
                    itemName = "CategoryS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("ItemCategory"));
                    formItems.Add("LinkTo", "CategoryE");
                    formItems.Add("FromPane", pane);
                    formItems.Add("ToPane", pane);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    bool multiSelection = false;
                    string objectType = "UDO_F_BDOSITMCTG_D";
                    string uniqueID_Category = "Category_CFL";
                    FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_Category);

                    formItems = new Dictionary<string, object>();
                    itemName = "CategoryE"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("Length", 20);
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Left", left + 5 + width_s);
                    formItems.Add("Width", width_e / 3);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("Enabled", true);
                    formItems.Add("ChooseFromListUID", uniqueID_Category);
                    formItems.Add("ChooseFromListAlias", "Code");
                    formItems.Add("FromPane", pane);
                    formItems.Add("ToPane", pane);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    //golden errow
                    formItems = new Dictionary<string, object>();
                    itemName = "CategoryLB"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    formItems.Add("Left", left + 5 + width_s - 20);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("LinkTo", "CategoryE");
                    formItems.Add("LinkedObjectType", objectType);
                    formItems.Add("FromPane", pane);
                    formItems.Add("ToPane", pane);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "CategoryN"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("Length", 20);
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Left", left + 5 + width_s + width_e / 3 + 5);
                    formItems.Add("Width", width_e * 2 / 3 - 5);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("Enabled", false);
                    formItems.Add("FromPane", pane);
                    formItems.Add("ToPane", pane);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    top = top + height + 5;

                    formItems = new Dictionary<string, object>();
                    itemName = "ItemFrS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("ItemFrom"));
                    formItems.Add("LinkTo", "ItemFrE");
                    formItems.Add("FromPane", pane);
                    formItems.Add("ToPane", pane);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    objectType = "4";
                    string uniqueID_Item = "ItemFr_CFL";
                    FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_Item);

                    //პირობის დადება პროდუქტის არჩევის სიაზე
                    SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(uniqueID_Item);
                    SAPbouiCOM.Conditions oCons = oCFL.GetConditions();
                    SAPbouiCOM.Condition oCon = oCons.Add();
                    oCon.Alias = "ItemType";
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCon.CondVal = "I";

                    oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                    oCon = oCons.Add();
                    oCon.Alias = "InvntItem";
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCon.CondVal = "Y";
                    oCFL.SetConditions(oCons);

                    formItems = new Dictionary<string, object>();
                    itemName = "ItemFrE"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("Length", 20);
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Left", left + 5 + width_s);
                    formItems.Add("Width", width_e / 3);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("Enabled", true);
                    formItems.Add("ChooseFromListUID", uniqueID_Item);
                    formItems.Add("ChooseFromListAlias", "ItemCode");
                    formItems.Add("FromPane", pane);
                    formItems.Add("ToPane", pane);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    //golden errow
                    formItems = new Dictionary<string, object>();
                    itemName = "ItemFrLB"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    formItems.Add("Left", left + 5 + width_s - 20);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("LinkTo", "ItemFrE");
                    formItems.Add("LinkedObjectType", objectType);
                    formItems.Add("FromPane", pane);
                    formItems.Add("ToPane", pane);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    ///////////////////////////////////////////
                    formItems = new Dictionary<string, object>();
                    itemName = "ItemToS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left + 5 + width_s + width_e / 3 + 5);
                    formItems.Add("Width", width_e / 3 - 5);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("To"));
                    formItems.Add("LinkTo", "ItemToE");
                    formItems.Add("FromPane", pane);
                    formItems.Add("ToPane", pane);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    string uniqueID_ItemTo = "ItemTo_CFL";
                    FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_ItemTo);

                    //პირობის დადება პროდუქტის არჩევის სიაზე
                    oCFL = oForm.ChooseFromLists.Item(uniqueID_ItemTo);
                    oCons = oCFL.GetConditions();
                    oCon = oCons.Add();
                    oCon.Alias = "ItemType";
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCon.CondVal = "I";

                    oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                    oCon = oCons.Add();
                    oCon.Alias = "InvntItem";
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCon.CondVal = "Y";
                    oCFL.SetConditions(oCons);

                    formItems = new Dictionary<string, object>();
                    itemName = "ItemToE"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("Length", 20);
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Left", (left + 5 + width_s + width_e / 3 + 5) + width_e / 3 - 5);
                    formItems.Add("Width", width_e / 3);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("Enabled", true);
                    formItems.Add("ChooseFromListUID", uniqueID_ItemTo);
                    formItems.Add("ChooseFromListAlias", "ItemCode");
                    formItems.Add("FromPane", pane);
                    formItems.Add("ToPane", pane);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    //golden errow
                    formItems = new Dictionary<string, object>();
                    itemName = "ItemToLB"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    formItems.Add("Left", (left + 5 + width_s + width_e / 3 + 5) + width_e / 3 - 20);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("LinkTo", "ItemToE");
                    formItems.Add("LinkedObjectType", objectType);
                    formItems.Add("FromPane", pane);
                    formItems.Add("ToPane", pane);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    top = top + height + 5;

                    formItems = new Dictionary<string, object>();
                    itemName = "WhsToS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("WhsTo"));
                    formItems.Add("LinkTo", "WhsToE");
                    formItems.Add("FromPane", pane);
                    formItems.Add("ToPane", pane);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    objectType = "64";
                    string uniqueID_WhsTo = "WhsTo_CFL";
                    FormsB1.addChooseFromList(oForm, true, objectType, uniqueID_WhsTo);

                    formItems = new Dictionary<string, object>();
                    itemName = "WhsToE"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("Length", 20);
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Left", left + 5 + width_s);
                    formItems.Add("Width", width_e);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("Enabled", true);
                    formItems.Add("ChooseFromListUID", uniqueID_WhsTo);
                    formItems.Add("ChooseFromListAlias", "WhsCode");
                    formItems.Add("FromPane", pane);
                    formItems.Add("ToPane", pane);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    //golden errow
                    formItems = new Dictionary<string, object>();
                    itemName = "WhsToLB"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    formItems.Add("Left", left + 5 + width_s - 20);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("LinkTo", "WhsToE");
                    formItems.Add("LinkedObjectType", objectType);
                    formItems.Add("FromPane", pane);
                    formItems.Add("ToPane", pane);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    top = top + height + 5;

                    //formItems = new Dictionary<string, object>();
                    //itemName = "TableTxt"; //10 characters
                    //formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    //formItems.Add("Left", left);
                    //formItems.Add("Width", width_s);
                    //formItems.Add("Top", top);
                    //formItems.Add("Height", height);
                    //formItems.Add("UID", itemName);
                    //formItems.Add("Caption", BDOSResources.getTranslate("Warehouse"));
                    //formItems.Add("TextStyle", 4);
                    //formItems.Add("FontSize", 10);
                    //formItems.Add("Enabled", true);
                    //formItems.Add("FromPane", pane);
                    //formItems.Add("ToPane", pane);

                    //FormsB1.createFormItem(oForm, formItems, out errorText);
                    //if (errorText != null)
                    //{
                    //    return;
                    //}

                    //FindNext
                    left = oForm.Width - width_s - width_e - 25;
                    //left = left + width_s + 5;

                    formItems = new Dictionary<string, object>();
                    itemName = "FindNext"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("FindNext"));
                    formItems.Add("FromPane", pane);
                    formItems.Add("ToPane", pane);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "FindNextE"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("Length", 20);
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Left", left + 5 + width_s);
                    formItems.Add("Width", width_e);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("Enabled", true);
                    formItems.Add("FromPane", pane);
                    formItems.Add("ToPane", pane);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }
                    //FindNext

                    left = 6;
                    top = top + height + 5;

                    //ცხრილი
                    itemName = "FiltTable";
                    formItems = new Dictionary<string, object>();
                    formItems.Add("isDataSource", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_GRID);
                    formItems.Add("Left", left);
                    formItems.Add("Width", oForm.Width - 20);
                    formItems.Add("Top", top);
                    formItems.Add("Height", 350);
                    formItems.Add("UID", itemName);
                    formItems.Add("FromPane", pane);
                    formItems.Add("ToPane", pane);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    SAPbouiCOM.DataTable oDataTable;
                    oDataTable = oForm.DataSources.DataTables.Add("FiltTable");
                    oDataTable.Columns.Add("ChkBx", SAPbouiCOM.BoFieldsType.ft_Text, 1);
                    oDataTable.Columns.Add("Code", SAPbouiCOM.BoFieldsType.ft_Text, 20);
                    oDataTable.Columns.Add("Name", SAPbouiCOM.BoFieldsType.ft_Text, 100);
                    oDataTable.Columns.Add("ProjectCode", SAPbouiCOM.BoFieldsType.ft_Text, 50);
                    oDataTable.Columns.Add("WhseType", SAPbouiCOM.BoFieldsType.ft_Text, 50);

                    //მეორე გვერდი
                    top = 6;
                    pane = 2;

                    //ცხრილური ნაწილები
                    formItems = new Dictionary<string, object>();
                    itemName = "Check";
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 19);
                    formItems.Add("Top", top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("Image", "HANA_CHECKBOX_CH");
                    formItems.Add("FromPane", pane);
                    formItems.Add("ToPane", pane);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "Uncheck";
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left + 20 + 1);
                    formItems.Add("Width", 19);
                    formItems.Add("Top", top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("Image", "HANA_CHECKBOX_UH");
                    formItems.Add("FromPane", pane);
                    formItems.Add("ToPane", pane);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "Fill"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left + 2 * (20 + 1));
                    formItems.Add("Width", 100);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("Fill"));
                    formItems.Add("FromPane", pane);
                    formItems.Add("ToPane", pane);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    top = top + height + 5;

                    //მატრიცა
                    formItems = new Dictionary<string, object>();
                    itemName = "WhsTable"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_GRID);
                    formItems.Add("Left", left);
                    formItems.Add("Width", oForm.Width);
                    formItems.Add("Top", top);
                    formItems.Add("Height", 100);
                    formItems.Add("UID", itemName);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("AffectsFormMode", false);
                    formItems.Add("FromPane", pane);
                    formItems.Add("ToPane", pane);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    oDataTable = oForm.DataSources.DataTables.Add("WhsTable");
                    oDataTable.Columns.Add("ChkBx", SAPbouiCOM.BoFieldsType.ft_Text, 1);
                    oDataTable.Columns.Add("LineID", SAPbouiCOM.BoFieldsType.ft_Text, 20);
                    oDataTable.Columns.Add("WhsFr", SAPbouiCOM.BoFieldsType.ft_Text, 20);
                    oDataTable.Columns.Add("WhseTypeFr", SAPbouiCOM.BoFieldsType.ft_Text, 50);
                    oDataTable.Columns.Add("WhsTo", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20);
                    oDataTable.Columns.Add("PrjFr", SAPbouiCOM.BoFieldsType.ft_Text, 20);
                    oDataTable.Columns.Add("PrjTo", SAPbouiCOM.BoFieldsType.ft_Text, 20);
                    oDataTable.Columns.Add("Position", SAPbouiCOM.BoFieldsType.ft_Sum, 20);
                    oDataTable.Columns.Add("Qty", SAPbouiCOM.BoFieldsType.ft_Sum, 20);
                    oDataTable.Columns.Add("Cost", SAPbouiCOM.BoFieldsType.ft_Sum, 20);
                    oDataTable.Columns.Add("DocID", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20);

                    create_TableWhsItemsForDetail();
                    FormsB1.addChooseFromList(oForm, true, "64", "TableWhsTo_CFL"); //საწყობებით დაჯგუფებული ცხრილისთვის
                }

                SAPbouiCOM.ComboBox oComboBoxOpType = (SAPbouiCOM.ComboBox)oForm.Items.Item("DocTypeE").Specific;
                oComboBoxOpType.Select("TransferToProjectWarehouse", SAPbouiCOM.BoSearchKey.psk_ByValue);

                oForm.Visible = true;
                oForm.Select();
            }

            GC.Collect();

        }

        public static void createDetailForm(string whsFr, string whseTypeFr, string whsTo, string docType, out string errorText)
        {
            errorText = null;

            Dictionary<string, object> formProperties = new Dictionary<string, object>();
            formProperties.Add("UniqueID", "BDOSStockTransferDetail");
            formProperties.Add("BorderStyle", SAPbouiCOM.BoFormBorderStyle.fbs_Sizable);
            formProperties.Add("Title", BDOSResources.getTranslate("StockTransferWizard") + " (" + BDOSResources.getTranslate("InDetail") + ")");
            formProperties.Add("Left", 558);
            formProperties.Add("ClientWidth", 750);
            formProperties.Add("Top", 335);
            formProperties.Add("ClientHeight", 400);
            //formProperties.Add("Modality", SAPbouiCOM.BoFormModality.fm_Modal);

            SAPbouiCOM.Form oForm;
            bool newForm;
            bool formExist = FormsB1.createForm(formProperties, out oForm, out newForm, out errorText);

            if (formExist == true)
            {
                if (newForm == true)
                {
                    Dictionary<string, object> formItems;
                    string itemName = "";

                    int height = 15;
                    int top = 6;
                    int width_s = 100;
                    int width_e = 200;
                    int left = 6;

                    formItems = new Dictionary<string, object>();
                    itemName = "DocTypeS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("DocType"));
                    formItems.Add("LinkTo", "DocTypeE");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    Dictionary<string, string> listValidValuesDict = new Dictionary<string, string>();
                    listValidValuesDict.Add("TransferToProjectWarehouse", BDOSResources.getTranslate("TransferToProjectWarehouse"));
                    listValidValuesDict.Add("TransferToWriteOffWarehouse", BDOSResources.getTranslate("TransferToWriteOffWarehouse"));
                    listValidValuesDict.Add("ReturnToProjectWarehouse", BDOSResources.getTranslate("ReturnToProjectWarehouse"));
                    listValidValuesDict.Add("ReturnToMainWarehouse", BDOSResources.getTranslate("ReturnToMainWarehouse"));
                    listValidValuesDict.Add("TransferWithoutType", BDOSResources.getTranslate("TransferWithoutType"));

                    formItems = new Dictionary<string, object>();
                    itemName = "DocTypeE";
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                    formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("Left", left + 5 + width_s);
                    formItems.Add("Width", width_e);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("ValidValues", listValidValuesDict);
                    formItems.Add("Enabled", false);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    top = top + height + 5;

                    formItems = new Dictionary<string, object>();
                    itemName = "WhsFrS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("WhsFr"));
                    formItems.Add("LinkTo", "WhsFrE");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    string objectType = "64";
                    string uniqueID_WhsFr = "WhsFr_CFL";
                    FormsB1.addChooseFromList(oForm, true, objectType, uniqueID_WhsFr);

                    formItems = new Dictionary<string, object>();
                    itemName = "WhsFrE"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("Length", 20);
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Left", left + 5 + width_s);
                    formItems.Add("Width", width_e);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("Enabled", false);
                    formItems.Add("ChooseFromListUID", uniqueID_WhsFr);
                    formItems.Add("ChooseFromListAlias", "WhsCode");
                    formItems.Add("ValueEx", whsFr);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    //golden errow
                    formItems = new Dictionary<string, object>();
                    itemName = "WhsFrLB"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    formItems.Add("Left", left + 5 + width_s - 20);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("LinkTo", "WhsFrE");
                    formItems.Add("LinkedObjectType", objectType);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left = left + 5 + width_s + width_e + 20;
                    //top = top + height + 5;

                    formItems = new Dictionary<string, object>();
                    itemName = "WhsToS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("WhsTo"));
                    formItems.Add("LinkTo", "WhsToE");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    string uniqueID_WhsTo = "WhsTo_CFL";
                    FormsB1.addChooseFromList(oForm, true, objectType, uniqueID_WhsTo);

                    formItems = new Dictionary<string, object>();
                    itemName = "WhsToE"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("Length", 20);
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Left", left + 5 + width_s);
                    formItems.Add("Width", width_e);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("Enabled", false);
                    formItems.Add("ChooseFromListUID", uniqueID_WhsTo);
                    formItems.Add("ChooseFromListAlias", "WhsCode");
                    formItems.Add("ValueEx", whsTo);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    //golden errow
                    formItems = new Dictionary<string, object>();
                    itemName = "WhsToLB"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    formItems.Add("Left", left + 5 + width_s - 20);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("LinkTo", "WhsToE");
                    formItems.Add("LinkedObjectType", objectType);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    //ცხრილური ნაწილები
                    left = 6;
                    top = top + height + 5;

                    formItems = new Dictionary<string, object>();
                    itemName = "Check";
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 19);
                    formItems.Add("Top", top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("Image", "HANA_CHECKBOX_CH");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "Uncheck";
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left + 20 + 1);
                    formItems.Add("Width", 19);
                    formItems.Add("Top", top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("Image", "HANA_CHECKBOX_UH");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "Split"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left + 2 * (20 + 1));
                    formItems.Add("Width", 100);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("Split"));

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    //FindNext
                    left = left + 5 + width_s + width_e + 20;
                    //top = top + height + 5;

                    formItems = new Dictionary<string, object>();
                    itemName = "FindNext"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("FindNext"));

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "FindNextE"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("Length", 20);
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Left", left + 5 + width_s);
                    formItems.Add("Width", width_e);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("Enabled", true);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }
                    //FindNext

                    left = 6;
                    top = top + height + 5;

                    //მატრიცა
                    formItems = new Dictionary<string, object>();
                    itemName = "WhsTblDt"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_GRID);
                    formItems.Add("Left", left);
                    formItems.Add("Width", oForm.Width);
                    formItems.Add("Top", top);
                    formItems.Add("Height", oForm.Height - top - 100);
                    formItems.Add("UID", itemName);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("AffectsFormMode", false);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Add("WhsTblDt");
                    oDataTable.Columns.Add("ChkBx", SAPbouiCOM.BoFieldsType.ft_Text, 1);
                    oDataTable.Columns.Add("LineID", SAPbouiCOM.BoFieldsType.ft_Text, 20);
                    oDataTable.Columns.Add("WhsFr", SAPbouiCOM.BoFieldsType.ft_Text, 20);
                    oDataTable.Columns.Add("WhseTypeFr", SAPbouiCOM.BoFieldsType.ft_Text, 50);
                    oDataTable.Columns.Add("WhsTo", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20);
                    oDataTable.Columns.Add("ItemCode", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20);
                    oDataTable.Columns.Add("ItemName", SAPbouiCOM.BoFieldsType.ft_Text, 100);
                    oDataTable.Columns.Add("PrjFr", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20);
                    oDataTable.Columns.Add("PrjTo", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20);
                    oDataTable.Columns.Add("UomCode", SAPbouiCOM.BoFieldsType.ft_Text, 20);
                    oDataTable.Columns.Add("InStock", SAPbouiCOM.BoFieldsType.ft_Sum, 20);
                    oDataTable.Columns.Add("Cost", SAPbouiCOM.BoFieldsType.ft_Sum, 20);
                    oDataTable.Columns.Add("CostInStock", SAPbouiCOM.BoFieldsType.ft_Sum, 20);
                    oDataTable.Columns.Add("Qty", SAPbouiCOM.BoFieldsType.ft_Sum, 20);

                    FormsB1.addChooseFromList(oForm, true, "64", "DetailTableWhsTo_CFL"); //საწყობებით დეტალური ცხრილისთვის
                    fillDetailTable(oForm, currentLineIDWhsTable, whsFr, whseTypeFr, whsTo, docType, out errorText);

                    //ღილაკები
                    top = oForm.ClientHeight - 25;
                    width_s = 65;

                    itemName = "1";
                    formItems = new Dictionary<string, object>();
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("OK"));

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    itemName = "2";
                    formItems = new Dictionary<string, object>();
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left + width_s + 1);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    SAPbouiCOM.ComboBox oComboBoxOpType = (SAPbouiCOM.ComboBox)oForm.Items.Item("DocTypeE").Specific;
                    oComboBoxOpType.Select(docType, SAPbouiCOM.BoSearchKey.psk_ByValue);
                }

                oForm.Visible = true;
                oForm.Select();
            }
            GC.Collect();
        }

        public static void createSplitForm(SAPbouiCOM.Form oDocForm, out string errorText)
        {
            errorText = null;

            //ფორმის აუცილებელი თვისებები
            Dictionary<string, object> formProperties = new Dictionary<string, object>();
            formProperties.Add("UniqueID", "BDOSStockTransferSplit");
            formProperties.Add("BorderStyle", SAPbouiCOM.BoFormBorderStyle.fbs_Sizable);
            formProperties.Add("Title", BDOSResources.getTranslate("NewRowQuantity"));
            formProperties.Add("Left", oDocForm.Left);
            formProperties.Add("Width", 200);
            formProperties.Add("Top", oDocForm.Top);
            formProperties.Add("Height", 10);
            formProperties.Add("Modality", SAPbouiCOM.BoFormModality.fm_Modal);

            SAPbouiCOM.Form oForm;
            bool newForm;
            bool formExist = FormsB1.createForm(formProperties, out oForm, out newForm, out errorText);

            if (formExist == true)
            {
                if (newForm == true)
                {
                    //ფორმის ელემენტების თვისებები
                    Dictionary<string, object> formItems = null;

                    int Top = 1;
                    int left = 6;

                    formItems = new Dictionary<string, object>();
                    string itemName = "newQty";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_QUANTITY);
                    formItems.Add("Length", 30);
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    Top = Top + 19 + 5;
                    left = 6;

                    itemName = "Split";
                    formItems = new Dictionary<string, object>();
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 120);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("Split"));

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }
                }

                oForm.Visible = true;
            }


            GC.Collect();
        }

        public static DataTable create_TableWhsItemsForDetail()
        {
            TableWhsItemsForDetail = new DataTable { TableName = "TableWhsItemsForDetail" };
            TableWhsItemsForDetail.Columns.Add("LineIDWhsTable", typeof(Int32));
            TableWhsItemsForDetail.Columns.Add("LineID", typeof(string));
            TableWhsItemsForDetail.Columns.Add("ChkBx", typeof(string));
            TableWhsItemsForDetail.Columns.Add("WhsFr", typeof(string));
            TableWhsItemsForDetail.Columns.Add("WhseTypeFr", typeof(string));
            TableWhsItemsForDetail.Columns.Add("WhsTo", typeof(string));
            TableWhsItemsForDetail.Columns.Add("PrjFr", typeof(string));
            TableWhsItemsForDetail.Columns.Add("PrjTo", typeof(string));
            TableWhsItemsForDetail.Columns.Add("ItemCode", typeof(string));
            TableWhsItemsForDetail.Columns.Add("ItemName", typeof(string));
            TableWhsItemsForDetail.Columns.Add("UomCode", typeof(string));
            TableWhsItemsForDetail.Columns.Add("InStock", typeof(decimal));
            TableWhsItemsForDetail.Columns.Add("Qty", typeof(decimal));
            TableWhsItemsForDetail.Columns.Add("Cost", typeof(decimal));
            TableWhsItemsForDetail.Columns.Add("CostInStock", typeof(decimal));

            return TableWhsItemsForDetail;
        }

        public static void uiApp_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (pVal.FormUID == "BDOSSTTRWZ")
                {
                    //try
                    //{
                    //    oFormDetailWizard2.Close();
                    //}
                    //catch
                    //{

                    //}

                    oFormWizard = oForm;
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_CLOSE & pVal.BeforeAction == true)
                    {
                        int answer = Program.uiApp.MessageBox(BDOSResources.getTranslate("DoYouWantToClose") + " " + BDOSResources.getTranslate("StockTransferWizard") + "?", 1, BDOSResources.getTranslate("Yes"), BDOSResources.getTranslate("No"), "");
                        if (answer == 2)
                        {
                            BubbleEvent = false;
                            return;
                        }
                    }

                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK & pVal.BeforeAction == false)
                    {
                        if (pVal.ItemUID == "Fill")
                        {
                            fillWhsTables(oForm, out errorText);
                            oForm.Update();
                        }
                    }

                    if (pVal.ItemUID == "FindNext" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.BeforeAction == false)
                    {
                        findNextRow(oForm, "FiltTable", "Code", "Name", true, out errorText);
                    }

                    if (pVal.ItemUID == "FindNextE" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_KEY_DOWN && pVal.BeforeAction == false)
                    {
                        findNextRow(oForm, "FiltTable", "Code", "Name", false, out errorText);
                    }

                    if (pVal.ItemUID == "WhsTable" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK & pVal.BeforeAction == false)
                    {
                        if (pVal.Row >= 0)
                        {
                            oFormWizard = oForm;
                            SAPbouiCOM.Grid oGrid = oForm.Items.Item("WhsTable").Specific;
                            SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("WhsTable");

                            int dTableRow = oGrid.GetDataTableRowIndex(pVal.Row);
                            string whsTo = oDataTable.GetValue("WhsTo", dTableRow);
                            string whsFr = oDataTable.GetValue("WhsFr", dTableRow);
                            string whseTypeFr = oDataTable.GetValue("WhseTypeFr", dTableRow);
                            currentLineIDWhsTable = oDataTable.GetValue("LineID", dTableRow);

                            SAPbouiCOM.ComboBox oComboBoxDocType = (SAPbouiCOM.ComboBox)oForm.Items.Item("DocTypeE").Specific;
                            string docType = oComboBoxDocType.Value;

                            createDetailForm(whsFr, whseTypeFr, whsTo, docType, out errorText);
                        }
                    }

                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.BeforeAction == false)
                    {
                        oForm.Freeze(true);
                        setVisibleFormItems(oForm, out errorText);
                        oForm.Freeze(false);
                    }

                    if (pVal.ItemUID == "DocTypeE" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_COMBO_SELECT && pVal.BeforeAction == false)
                    {
                        fillFiltTable(oForm, out errorText);
                    }
                    if (pVal.ItemUID == "Prev" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.BeforeAction == false)
                    {
                        int answer = Program.uiApp.MessageBox(BDOSResources.getTranslate("AllDataWillBeCleared"), 1, BDOSResources.getTranslate("Yes"), BDOSResources.getTranslate("No"), "");
                        if (answer == 1)
                        {
                            if (oForm.Items.Item("Prev").Enabled)
                            {
                                changePane(oForm, -1);
                            }
                        }
                        else
                        {
                            BubbleEvent = false;
                            return;
                        }
                    }

                    if (pVal.ItemUID == "Next" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.BeforeAction == false)
                    {
                        changePane(oForm, 1);
                    }

                    if (pVal.ItemUID == "CategoryE" && pVal.BeforeAction == false)
                    {
                        SAPbouiCOM.EditText oEdit = oForm.Items.Item("CategoryE").Specific;
                        string category = oEdit.Value;
                        if (string.IsNullOrEmpty(category))
                        {
                            oForm.Items.Item("CategoryN").Specific.Value = "";
                        }
                    }

                    if ((pVal.ItemUID == "Check" || pVal.ItemUID == "Uncheck") && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.BeforeAction == false)
                    {
                        checkUncheckTables(oForm, pVal.ItemUID, "WhsTable", out errorText);
                    }

                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                    {
                        SAPbouiCOM.IChooseFromListEvent oCFLEvento = null;
                        oCFLEvento = ((SAPbouiCOM.IChooseFromListEvent)(pVal));

                        chooseFromList(oForm, oCFLEvento, pVal);
                    }

                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_RESIZE && !pVal.BeforeAction)
                    {
                        oForm.Freeze(true);
                        resizeForm(oForm, out errorText);
                        oForm.Freeze(false);
                    }
                }
                else if (pVal.FormUID == "BDOSStockTransferDetail")
                {
                    //oFormDetailWizard2 = oForm;
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.ItemUID == "1" && !pVal.BeforeAction && oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                    {
                        updateDetailTable(oForm, out errorText);
                    }

                    if ((pVal.ItemUID == "Check" || pVal.ItemUID == "Uncheck") && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && !pVal.BeforeAction)
                    {
                        checkUncheckTables(oForm, pVal.ItemUID, "WhsTblDt", out errorText);
                        //updateDetailTable(oForm, out errorText);
                    }

                    if (pVal.ItemUID == "FindNext" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && !pVal.BeforeAction)
                    {
                        findNextRow(oForm, "WhsTblDt", "ItemCode", "ItemName", true, out errorText);
                    }

                    if (pVal.ItemUID == "FindNextE" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_KEY_DOWN && !pVal.BeforeAction)
                    {
                        findNextRow(oForm, "WhsTblDt", "ItemCode", "ItemName", false, out errorText);
                    }

                    if (pVal.ItemUID == "Split" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && !pVal.BeforeAction)
                    {
                        SAPbouiCOM.Grid oGrid = ((SAPbouiCOM.Grid)(oForm.Items.Item("WhsTblDt").Specific));
                        SAPbouiCOM.SelectedRows selectedRows = oGrid.Rows.SelectedRows;

                        if (oGrid.Rows.SelectedRows.Count == 1)
                        {
                            oFormDetailWizard = oForm;
                            createSplitForm(oForm, out errorText);
                        }
                    }

                    if (pVal.ItemUID == "WhsTblDt" && pVal.ColUID == "Qty" && (pVal.EventType == SAPbouiCOM.BoEventTypes.et_VALIDATE || pVal.ItemChanged))
                    {
                        try
                        {
                            oForm.Freeze(true);

                            SAPbouiCOM.Grid oGrid = oForm.Items.Item("WhsTblDt").Specific;
                            SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("WhsTblDt");
                            int dTableRow = oGrid.GetDataTableRowIndex(pVal.Row);

                            decimal qty = Convert.ToDecimal(oDataTable.GetValue("Qty", dTableRow), CultureInfo.InvariantCulture);
                            decimal inStock = Convert.ToDecimal(oDataTable.GetValue("InStock", dTableRow), CultureInfo.InvariantCulture);

                            if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_VALIDATE && pVal.BeforeAction)
                            {
                                if (qty > inStock)
                                {
                                    oDataTable.SetValue("Qty", dTableRow, oDataTable.GetValue("InStock", dTableRow));
                                    Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("InsufficientStockBalance") + "! " + BDOSResources.getTranslate("TableRow") + ": " + (dTableRow + 1), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                    BubbleEvent = true;
                                }
                            }

                            else if (pVal.ItemChanged)
                            {
                                decimal costInStock = Convert.ToDecimal(oDataTable.GetValue("CostInStock", dTableRow), CultureInfo.InvariantCulture);
                                decimal price = inStock != 0 ? costInStock / inStock : 0;
                                decimal newCost = price * qty;

                                oDataTable.SetValue("Cost", dTableRow, Convert.ToDouble(CommonFunctions.roundAmountByGeneralSettings(newCost, "Sum")));
                            }
                        }
                        catch (Exception ex)
                        {
                            Program.uiApp.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                        }
                        finally
                        {
                            oForm.Freeze(false);
                        }
                    }

                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                    {
                        SAPbouiCOM.IChooseFromListEvent oCFLEvento = null;
                        oCFLEvento = ((SAPbouiCOM.IChooseFromListEvent)(pVal));

                        chooseFromList(oForm, oCFLEvento, pVal);
                    }

                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_RESIZE && !pVal.BeforeAction)
                    {
                        oForm.Freeze(true);
                        resizeForm(oForm, out errorText);
                        oForm.Freeze(false);
                    }
                }
                else if ((pVal.FormUID == "BDOSStockTransferSplit"))
                {
                    if (pVal.ItemUID == "Split" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.BeforeAction == false)
                    {
                        splitDetailRow(oForm, out errorText);
                    }
                }
            }
        }

        public static void resizeForm(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            try
            {
                SAPbouiCOM.Item oItem = null;

                if (oForm.UniqueID == "BDOSSTTRWZ")
                {
                    oItem = oForm.Items.Item("FiltTable");
                    oItem.Width = oForm.Width - 15;
                    oItem.Height = oForm.Height - oItem.Top - 100;

                    oItem = oForm.Items.Item("WhsTable");
                    oItem.Width = oForm.Width - 15;
                    oItem.Height = oForm.Height - oItem.Top - 100;

                    SAPbouiCOM.Grid oGrid = oForm.Items.Item("FiltTable").Specific;
                    if (oGrid.Columns.Count > 0)
                    {
                        SAPbouiCOM.GridColumn oColumn = oGrid.Columns.Item("ChkBx");
                        oColumn.Width = 20;
                    }

                    oGrid = oForm.Items.Item("WhsTable").Specific;
                    if (oGrid.Columns.Count > 0)
                    {
                        SAPbouiCOM.GridColumn oColumn = oGrid.Columns.Item("ChkBx");
                        oColumn.Width = 20;
                    }
                }
                else
                {
                    oItem = oForm.Items.Item("WhsToLB");
                    oItem.Left = oForm.Items.Item("WhsToS").Left + oForm.Items.Item("WhsToS").Width - 15;

                    oItem = oForm.Items.Item("WhsToE");
                    oItem.Left = oForm.Items.Item("WhsToS").Left + oForm.Items.Item("WhsToS").Width + 5;

                    oItem = oForm.Items.Item("FindNextE");
                    oItem.Left = oForm.Items.Item("FindNext").Left + oForm.Items.Item("FindNext").Width + 5;

                    oItem = oForm.Items.Item("WhsTblDt");
                    oItem.Width = oForm.Width - 15;
                    oItem.Height = oForm.Height - oItem.Top - 100;

                    SAPbouiCOM.Grid oGrid = oForm.Items.Item("WhsTblDt").Specific;
                    if (oGrid.Columns.Count > 0)
                    {
                        SAPbouiCOM.GridColumn oColumn = oGrid.Columns.Item("ChkBx");
                        oColumn.Width = 20;
                    }
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
        }

        private static void changePane(SAPbouiCOM.Form oForm, int PaneVal)
        {
            string errorText = "";

            oForm.Freeze(true);

            bool checkError = false;

            if (oForm.PaneLevel == 3 && PaneVal > 0)
            {
                oForm.Close();
                return;
            }

            if (oForm.PaneLevel == 1 && PaneVal == 1)
            {
                // ველების შემოწმება DocDate
                SAPbouiCOM.EditText oEditTextDate = (SAPbouiCOM.EditText)oForm.Items.Item("DateE").Specific;
                String date = oEditTextDate.Value;
                if (string.IsNullOrEmpty(date))
                {
                    Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("Date") + " " + BDOSResources.getTranslate("YouCantLeaveEmpty"), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    checkError = true;
                }

                SAPbouiCOM.ComboBox oComboBoxDocType = (SAPbouiCOM.ComboBox)oForm.Items.Item("DocTypeE").Specific;
                string docType = oComboBoxDocType.Value;
                //მიმღები საწყობის შემოწმება
                if (docType == "TransferToProjectWarehouse" || docType == "ReturnToMainWarehouse" || docType == "TransferWithoutType")
                {
                    SAPbouiCOM.EditText oEditWhsTo = (SAPbouiCOM.EditText)oForm.Items.Item("WhsToE").Specific;
                    if (string.IsNullOrEmpty(oEditWhsTo.Value))
                    {
                        Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("WhsTo") + " " + BDOSResources.getTranslate("YouCantLeaveEmpty"), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        checkError = true;
                    }
                }
            }

            if (checkError == false)
            {
                if (oForm.PaneLevel == 1 && PaneVal > 0)
                {
                    oForm.PaneLevel = oForm.PaneLevel + PaneVal;
                    fillWhsTables(oForm, out errorText);
                }
                else if (oForm.PaneLevel == 2 && PaneVal > 0)
                {
                    createStockTransferDocuments(oForm);
                }
                else
                {
                    oForm.PaneLevel = oForm.PaneLevel + PaneVal;
                }

                setVisibleFormItems(oForm, out errorText);
            }

            oForm.Freeze(false);
        }

        public static void chooseFromList(SAPbouiCOM.Form oForm, SAPbouiCOM.IChooseFromListEvent oCFLEvento, SAPbouiCOM.ItemEvent pVal)
        {
            bool beforeAction = pVal.BeforeAction;
            int row = pVal.Row;
            string sCFL_ID = oCFLEvento.ChooseFromListUID;

            try
            {
                SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(sCFL_ID);

                if (!beforeAction)
                {
                    SAPbouiCOM.DataTable oDataTableSelectedObjects = null;
                    oDataTableSelectedObjects = oCFLEvento.SelectedObjects;

                    if (oDataTableSelectedObjects != null)
                    {
                        if (sCFL_ID == "WhsTo_CFL")
                        {
                            string eCode = oDataTableSelectedObjects.GetValue("WhsCode", 0);
                            SAPbouiCOM.EditText oEdit = oForm.Items.Item("WhsToE").Specific;
                            try
                            {
                                oEdit.Value = eCode;
                            }
                            catch { }

                        }
                        else if (sCFL_ID == "Category_CFL")
                        {
                            string eCode = oDataTableSelectedObjects.GetValue("Code", 0);
                            string eName = oDataTableSelectedObjects.GetValue("Name", 0);

                            SAPbouiCOM.EditText oEdit = oForm.Items.Item("CategoryE").Specific;
                            try
                            {
                                oEdit.Value = eCode;
                            }
                            catch { }

                            oEdit = oForm.Items.Item("CategoryN").Specific;
                            try
                            {
                                oEdit.Value = eName;
                            }
                            catch { }
                        }
                        else if (sCFL_ID == "ItemFr_CFL")
                        {
                            string eCode = oDataTableSelectedObjects.GetValue("ItemCode", 0);
                            SAPbouiCOM.EditText oEdit = oForm.Items.Item("ItemFrE").Specific;
                            try
                            {
                                oEdit.Value = eCode;
                            }
                            catch { }
                        }
                        else if (sCFL_ID == "ItemTo_CFL")
                        {
                            string eCode = oDataTableSelectedObjects.GetValue("ItemCode", 0);
                            SAPbouiCOM.EditText oEdit = oForm.Items.Item("ItemToE").Specific;
                            try
                            {
                                oEdit.Value = eCode;
                            }
                            catch { }
                        }
                        else if (sCFL_ID == "TableWhsTo_CFL")
                        {
                            string whsTo = oDataTableSelectedObjects.GetValue("WhsCode", 0);
                            string prjTo = CommonFunctions.getValue("OWHS", "U_BDOSPrjCod", "WhsCode", whsTo).ToString();

                            SAPbouiCOM.Grid oGrid = oForm.Items.Item("WhsTable").Specific;
                            SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("WhsTable");

                            int dTableRow = oGrid.GetDataTableRowIndex(pVal.Row);
                            oDataTable.SetValue("WhsTo", dTableRow, whsTo);
                            oDataTable.SetValue("PrjTo", dTableRow, prjTo);
                        }
                        else if (sCFL_ID == "DetailTableWhsTo_CFL")
                        {
                            string whsTo = oDataTableSelectedObjects.GetValue("WhsCode", 0);
                            string prjTo = CommonFunctions.getValue("OWHS", "U_BDOSPrjCod", "WhsCode", whsTo).ToString();

                            SAPbouiCOM.Grid oGrid = oForm.Items.Item("WhsTblDt").Specific;
                            SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("WhsTblDt");

                            int dTableRow = oGrid.GetDataTableRowIndex(pVal.Row);
                            oDataTable.SetValue("WhsTo", dTableRow, whsTo);
                            oDataTable.SetValue("PrjTo", dTableRow, prjTo);

                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                        }
                    }
                }
                else
                {
                    if (sCFL_ID == "WhsTo_CFL")
                    {
                        string whsType = "";
                        SAPbouiCOM.ComboBox oComboBoxDocTyp = (SAPbouiCOM.ComboBox)oForm.Items.Item("DocTypeE").Specific;
                        string docType = oComboBoxDocTyp.Value;
                        if (docType == "TransferToProjectWarehouse" || docType == "ReturnToMainWarehouse")
                        {
                            whsType = (docType == "TransferToProjectWarehouse" ? "Project" : "Main"); //სხვა ოპ.ტიპის დროს არ აქვს ფილტრი                           
                        }
                        setWarehousesConditions("", whsType, null, oCFL);
                    }
                    else if (sCFL_ID == "TableWhsTo_CFL")
                    {
                        SAPbouiCOM.Grid oGrid = oForm.Items.Item("WhsTable").Specific;
                        SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("WhsTable");

                        int dTableRow = oGrid.GetDataTableRowIndex(pVal.Row);
                        string prjFr = oDataTable.GetValue("PrjFr", dTableRow);
                        //string whsType = oDataTable.GetValue("WhseTypeFr", dTableRow);
                        string whsFr = oDataTable.GetValue("WhsFr", dTableRow);
                        setWarehousesConditions(prjFr, null, whsFr, oCFL);
                    }
                    else if (sCFL_ID == "DetailTableWhsTo_CFL")
                    {
                        SAPbouiCOM.Grid oGrid = oForm.Items.Item("WhsTblDt").Specific;
                        SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("WhsTblDt");

                        int dTableRow = oGrid.GetDataTableRowIndex(pVal.Row);
                        string prjFr = oDataTable.GetValue("PrjFr", dTableRow);
                        //string whsType = oDataTable.GetValue("WhseTypeFr", dTableRow);
                        string whsFr = oDataTable.GetValue("WhsFr", dTableRow);
                        setWarehousesConditions(prjFr, null, whsFr, oCFL);
                    }
                    else if (sCFL_ID == "Category_CFL")
                    {
                        SAPbouiCOM.ComboBox oComboBoxCatLevel = (SAPbouiCOM.ComboBox)oForm.Items.Item("CatLvlE").Specific;
                        string catLevel = oComboBoxCatLevel.Value;

                        SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();
                        SAPbouiCOM.Condition oCon = oCons.Add();
                        oCon.Alias = "U_Level";
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = catLevel;
                        oCFL.SetConditions(oCons);
                    }
                    else if (sCFL_ID == "ItemFr_CFL")
                    {
                        SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();
                        SAPbouiCOM.Condition oCon = oCons.Add();
                        oCon.Alias = "ManBtchNum";
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = "N";
                        oCFL.SetConditions(oCons);
                    }
                    else if (sCFL_ID == "ItemTo_CFL")
                    {
                        SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();
                        SAPbouiCOM.Condition oCon = oCons.Add();
                        oCon.Alias = "ManBtchNum";
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = "N";
                        oCFL.SetConditions(oCons);
                    }
                }
            }
            catch (Exception ex)
            {
                Program.uiApp.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                GC.Collect();
            }
        }

        private static void checkUncheckTables(SAPbouiCOM.Form oForm, string CheckOperation, string tableName, out string errorText)
        {
            errorText = null;

            oForm.Freeze(true);

            SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item(tableName);

            string ChkBx = CheckOperation == "Check" ? "Y" : "N";
            for (int i = 0; i < oDataTable.Rows.Count; i++)
            {
                oDataTable.SetValue("ChkBx", i, ChkBx);
            }
            oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
            oForm.Freeze(false);
        }

        public static void setWarehousesConditions(string prjCode, string whsType, string whsCode, SAPbouiCOM.ChooseFromList oCFL)
        {
            SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();
            SAPbouiCOM.Condition oCon;

            if (!string.IsNullOrEmpty(prjCode))
            {
                oCon = oCons.Add();
                oCon.Alias = "U_BDOSPrjCod";
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCon.CondVal = prjCode;
                oCFL.SetConditions(oCons);

                oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                oCon = oCons.Add();
                oCon.Alias = "WhsCode";
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL;
                oCon.CondVal = whsCode;
                oCFL.SetConditions(oCons);

                oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
            }

            if (!string.IsNullOrEmpty(whsType))
            {
                oCon = oCons.Add();
                oCon.Alias = "U_BDOSWhType";
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCon.CondVal = whsType;
                oCFL.SetConditions(oCons);

                oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                oCon = oCons.Add();
                oCon.Alias = "WhsCode";
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL;
                oCon.CondVal = whsCode;
                oCFL.SetConditions(oCons);

                oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
            }

            oCon = oCons.Add();
            oCon.Alias = "Inactive";
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCon.CondVal = "N";
            oCFL.SetConditions(oCons);

            oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

            oCon = oCons.Add();
            oCon.Alias = "DropShip";
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCon.CondVal = "N";
            oCFL.SetConditions(oCons);

            oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

            oCon = oCons.Add();
            oCon.Alias = "Locked";
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCon.CondVal = "N";
            oCFL.SetConditions(oCons);
        }

        public static Dictionary<string, string> getItemGroupsList(out string errorText)
        {
            errorText = null;

            Dictionary<string, string> grpList = new Dictionary<string, string>();
            grpList.Add("", "");
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string query = @"SELECT * FROM ""OITB"" WHERE ""Locked""='N'";

            oRecordSet.DoQuery(query);

            while (!oRecordSet.EoF)
            {
                grpList.Add(oRecordSet.Fields.Item("ItmsGrpCod").Value.ToString(), oRecordSet.Fields.Item("ItmsGrpNam").Value.ToString());
                oRecordSet.MoveNext();
            }

            return grpList;
        }

        public static void setVisibleFormItems(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            SAPbouiCOM.ComboBox oComboBoxDocType = (SAPbouiCOM.ComboBox)oForm.Items.Item("DocTypeE").Specific;
            string docType = oComboBoxDocType.Value;

            //SAPbouiCOM.StaticText oStaticText = (SAPbouiCOM.StaticText)oForm.Items.Item("TableTxt").Specific;
            //oStaticText.Caption = (docType == "TransferToProjectWarehouse" || docType == "TransferWithoutType" ? BDOSResources.getTranslate("WareHouses") : BDOSResources.getTranslate("Project"));

            bool visibleWhsTo = oForm.PaneLevel == 1 && (docType == "TransferToProjectWarehouse" || docType == "ReturnToMainWarehouse" || docType == "TransferWithoutType");
            oForm.Items.Item("WhsToS").Visible = visibleWhsTo;
            oForm.Items.Item("WhsToE").Visible = visibleWhsTo;
            oForm.Items.Item("WhsToLB").Visible = visibleWhsTo;

            try
            {
                if (oForm.PaneLevel == 1)
                {
                    oForm.Items.Item("Prev").Enabled = false;
                }
                else
                {
                    oForm.Items.Item("Prev").Enabled = true;
                }

                if (oForm.PaneLevel == 2)
                {
                    oForm.Items.Item("Next").Specific.Caption = BDOSResources.getTranslate("CreateDoc");
                }
                else
                {
                    oForm.Items.Item("Next").Specific.Caption = BDOSResources.getTranslate("Next");
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
        }

        public static void fillFiltTable(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            oForm.Freeze(true);

            SAPbouiCOM.ComboBox oComboBoxDocType = (SAPbouiCOM.ComboBox)oForm.Items.Item("DocTypeE").Specific;
            string docType = oComboBoxDocType.Value;

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                SAPbouiCOM.Grid oGrid = ((SAPbouiCOM.Grid)(oForm.Items.Item("FiltTable").Specific));
                SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("FiltTable");
                oDataTable.Rows.Clear();

                StringBuilder Sbuilder = new StringBuilder();

                string XML = oDataTable.GetAsXML();
                XML = XML.Replace("<Rows/></DataTable>", "");
                Sbuilder.Append(XML);
                Sbuilder.Append("<Rows>");

                string query = @"SELECT 'N' AS ""ChkBx"",
                       ""WhsCode"" AS ""Code"",
                       ""WhsName"" AS ""Name"",
                       ""U_BDOSPrjCod"" AS ""ProjectCode"",
                       ""U_BDOSWhType"" AS ""WhseType""
                FROM ""OWHS""
                WHERE ""Inactive"" = 'N'
                  AND ""DropShip"" = 'N'
                  AND ""Locked"" = 'N'";

                if (docType == "TransferToProjectWarehouse") //გადაადგილება პროექტის საწყობზე
                {
                    query = query + @" AND ""OWHS"".""U_BDOSWhType"" = 'Main'";
                }
                else if (docType == "TransferToWriteOffWarehouse" || docType == "ReturnToMainWarehouse") //გადაადგილება ხარჯვის საწყობზე || დაბრუნება ცენტრალურ საწყობზე
                {
                    query = query + (@" AND ""OWHS"".""U_BDOSWhType"" = 'Project'");
                }
                else if (docType == "ReturnToProjectWarehouse") //დაბრუნება პროექტის საწყობზე
                {
                    query = query + (@" AND ""OWHS"".""U_BDOSWhType"" = 'WriteOff'");
                }
                //else if (docType != "TransferWithoutType") //გადაადგილება ტიპის განსაზღვრის გარეშე
                //{
                //}

                oDataTable.ExecuteQuery(query);

                oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
                oGrid.DataTable = oDataTable;
                SAPbouiCOM.GridColumns oColumns = oGrid.Columns;

                SAPbouiCOM.GridColumn oColumn = oColumns.Item("ChkBx");
                oColumn.Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
                oColumn.TitleObject.Caption = "";
                oColumn.Width = 20;

                SAPbouiCOM.EditTextColumn oEditTextColumn = (SAPbouiCOM.EditTextColumn)oColumns.Item("Code");
                oEditTextColumn.TitleObject.Caption = BDOSResources.getTranslate("Warehouse");
                oEditTextColumn.Editable = false;
                oEditTextColumn.LinkedObjectType = "64"; //პროექტი-63, საწყობი-64 (docType == "ReturnToProjectWarehouse" || docType == "ReturnToMainWarehouse" ? "63" : "64");

                oEditTextColumn = (SAPbouiCOM.EditTextColumn)oColumns.Item("Name");
                oEditTextColumn.TitleObject.Caption = BDOSResources.getTranslate("Name");
                oEditTextColumn.Editable = false;

                oEditTextColumn = (SAPbouiCOM.EditTextColumn)oColumns.Item("ProjectCode");
                oEditTextColumn.TitleObject.Caption = BDOSResources.getTranslate("Project");
                oEditTextColumn.Editable = false;
                oEditTextColumn.LinkedObjectType = "63"; //პროექტი-63

                oEditTextColumn = (SAPbouiCOM.EditTextColumn)oColumns.Item("WhseType");
                oEditTextColumn.TitleObject.Caption = BDOSResources.getTranslate("WhseType");
                oEditTextColumn.Editable = false;

                oColumns.Item("RowsHeader").Visible = false;

                setVisibleFormItems(oForm, out errorText);
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                oRecordSet = null;
                oForm.Freeze(false);
                GC.Collect();
            }
        }

        public static void fillWhsTables(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            oForm.Freeze(true);

            SAPbouiCOM.EditText oEditTextDate = (SAPbouiCOM.EditText)oForm.Items.Item("DateE").Specific;
            String dateStr = oEditTextDate.Value;
            if (string.IsNullOrEmpty(dateStr))
            {
                errorText = BDOSResources.getTranslate("Date") + " " + BDOSResources.getTranslate("YouCantLeaveEmpty");
                Program.uiApp.StatusBar.SetSystemMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return;
            }

            SAPbouiCOM.DataSource oDataSource = oForm.DataSources;
            SAPbouiCOM.UserDataSources oUserDataSources = oDataSource.UserDataSources;

            SAPbouiCOM.ComboBox oComboBoxDocType = (SAPbouiCOM.ComboBox)oForm.Items.Item("DocTypeE").Specific;
            string docType = oComboBoxDocType.Value;

            SAPbouiCOM.ComboBox oComboBoxItemGroup = (SAPbouiCOM.ComboBox)oForm.Items.Item("ItmGrpE").Specific;
            string itemGroup = oComboBoxItemGroup.Value;

            SAPbouiCOM.ComboBox oComboBoxCategoryLevel = (SAPbouiCOM.ComboBox)oForm.Items.Item("CatLvlE").Specific;
            string itemCategoryLevel = oComboBoxCategoryLevel.Value;

            string itemCodeFrom = oUserDataSources.Item("ItemFrE").ValueEx;
            string itemCodeTo = oUserDataSources.Item("ItemToE").ValueEx;
            string itemCategory = oUserDataSources.Item("CategoryE").ValueEx;
            string whsTo = "";
            string prjTo = "";
            if (docType == "TransferToProjectWarehouse" || docType == "ReturnToMainWarehouse" || docType == "TransferWithoutType")
            {
                whsTo = oUserDataSources.Item("WhsToE").ValueEx;
                prjTo = CommonFunctions.getValue("OWHS", "U_BDOSPrjCod", "WhsCode", whsTo).ToString();
            }

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                SAPbouiCOM.DataTable oDataTableFilter = oForm.DataSources.DataTables.Item("FiltTable");
                DateTime date = DateTime.ParseExact(dateStr, "yyyyMMdd", null);
                DateTime today = DateTime.Today;
                
                List<string> whseCodes = new List<string>();
                for (int i = 0; i < oDataTableFilter.Rows.Count; i++)
                {
                    if (oDataTableFilter.GetValue("ChkBx", i) == "Y")
                    {
                        whseCodes.Add("'" + oDataTableFilter.GetValue("Code", i).Trim() + "'");
                    }
                }
                string whseCodesStr = string.Join(",", whseCodes);

                StringBuilder queryBuilder = new StringBuilder();

                queryBuilder.Append(@"SELECT ""OIVL"".""LocCode"",
                       ""OIVL"".""ItemCode"",
                       ""OIVL"".""ItemName"",
                       ""OWHS"".""U_BDOSPrjCod"" AS ""PrjFr"",
                       ""OWHS"".""U_BDOSWhType"" AS ""WhseTypeFr"",
                       MIN(""OIVL"".""Qty"") AS ""Qty"",
                       MIN(""OIVL"".""Cost"") AS ""Cost""
                FROM
                  (SELECT ""OIVL"".""LocCode"",
                          ""OIVL"".""ItemCode"",
                          ""OITM"".""ItemName"",
                          SUM(""OIVL"".""InQty"" - ""OIVL"".""OutQty"") AS ""Qty"",
                          SUM(""OIVL"".""SumStock"") AS ""Cost""
                   FROM ""OIVL""
                   LEFT JOIN ""OITM"" ON ""OIVL"".""ItemCode"" = ""OITM"".""ItemCode""
                   WHERE
                     ""DocDate"" <= '" + date.ToString("yyyyMMdd") + @"'
                     AND ""OITM"".""ManBtchNum"" = 'N' AND ""OITM"".""ManSerNum"" = 'N' ");
                if (!string.IsNullOrEmpty(itemGroup))
                {
                    queryBuilder.Append(@" AND ""OITM"".""ItmsGrpCod"" = '");
                    queryBuilder.Append(itemGroup.Trim());
                    queryBuilder.Append("' ");
                }
                if (!string.IsNullOrEmpty(itemCategoryLevel) && !string.IsNullOrEmpty(itemCategory))
                {
                    queryBuilder.Append(@" AND ""OITM"".""U_BDOSCtg");
                    queryBuilder.Append(itemCategoryLevel.Trim());
                    queryBuilder.Append(@""" = '");
                    queryBuilder.Append(itemCategory);
                    queryBuilder.Append("' ");
                }
                if (!string.IsNullOrEmpty(itemCodeFrom))
                {
                    queryBuilder.Append(@" AND ""OIVL"".""ItemCode"" >= '");
                    queryBuilder.Append(itemCodeFrom.Trim());
                    queryBuilder.Append("' ");
                }
                if (!string.IsNullOrEmpty(itemCodeTo))
                {
                    queryBuilder.Append(@" AND ""OIVL"".""ItemCode"" <= '");
                    queryBuilder.Append(itemCodeTo.Trim());
                    queryBuilder.Append("' ");
                }
                if (!string.IsNullOrEmpty(whseCodesStr))
                {
                    queryBuilder.Append(@"AND ""OIVL"".""LocCode"" IN (");
                    queryBuilder.Append(whseCodesStr.Trim());
                    queryBuilder.Append(") ");
                }
                queryBuilder.Append(@"GROUP BY ""OIVL"".""LocCode"",
                            ""OIVL"".""ItemCode"",
                            ""OITM"".""ItemName""
                   UNION ALL SELECT ""OIVL"".""LocCode"",
                                    ""OIVL"".""ItemCode"",
                                    ""OITM"".""ItemName"",
                                    SUM(""OIVL"".""InQty"" - ""OIVL"".""OutQty"") AS ""Qty"",
                                    SUM(""OIVL"".""SumStock"") AS ""Cost""
                   FROM ""OIVL""
                   LEFT JOIN ""OITM"" ON ""OIVL"".""ItemCode"" = ""OITM"".""ItemCode""
                   WHERE
                     ""DocDate"" <= '" + today.ToString("yyyyMMdd") + @"'
                     AND ""OITM"".""ManBtchNum"" = 'N' AND ""OITM"".""ManSerNum"" = 'N' ");
                if (!string.IsNullOrEmpty(itemGroup))
                {
                    queryBuilder.Append(@" AND ""OITM"".""ItmsGrpCod"" = '");
                    queryBuilder.Append(itemGroup.Trim());
                    queryBuilder.Append("' ");
                }
                if (!string.IsNullOrEmpty(itemCategoryLevel) && !string.IsNullOrEmpty(itemCategory))
                {
                    queryBuilder.Append(@" AND ""OITM"".""U_BDOSCtg");
                    queryBuilder.Append(itemCategoryLevel.Trim());
                    queryBuilder.Append(@""" = '");
                    queryBuilder.Append(itemCategory);
                    queryBuilder.Append("' ");
                }
                if (!string.IsNullOrEmpty(itemCodeFrom))
                {
                    queryBuilder.Append(@" AND ""OIVL"".""ItemCode"" >= '");
                    queryBuilder.Append(itemCodeFrom.Trim());
                    queryBuilder.Append("' ");
                }
                if (!string.IsNullOrEmpty(itemCodeTo))
                {
                    queryBuilder.Append(@" AND ""OIVL"".""ItemCode"" <= '");
                    queryBuilder.Append(itemCodeTo.Trim());
                    queryBuilder.Append("' ");
                }
                if (!string.IsNullOrEmpty(whseCodesStr))
                {
                    queryBuilder.Append(@"AND ""OIVL"".""LocCode"" IN (");
                    queryBuilder.Append(whseCodesStr.Trim());
                    queryBuilder.Append(") ");
                }
                queryBuilder.Append(@"GROUP BY ""OIVL"".""LocCode"",
                            ""OIVL"".""ItemCode"",
                            ""OITM"".""ItemName"") AS ""OIVL""
                LEFT JOIN ""OWHS"" ON ""OIVL"".""LocCode"" = ""OWHS"".""WhsCode""
                GROUP BY ""OIVL"".""LocCode"",
                         ""OIVL"".""ItemCode"",
                         ""OIVL"".""ItemName"",
                         ""OWHS"".""U_BDOSPrjCod"",
                         ""OWHS"".""U_BDOSWhType""
                ORDER BY ""OIVL"".""LocCode""");

                string query = queryBuilder.ToString();

                SAPbouiCOM.Grid oGrid = ((SAPbouiCOM.Grid)(oForm.Items.Item("WhsTable").Specific));
                SAPbouiCOM.DataTable oDataTableWhs = oForm.DataSources.DataTables.Item("WhsTable");
                oDataTableWhs.Rows.Clear();
                StringBuilder SbuilderWhs = new StringBuilder();

                TableWhsItemsForDetail.Clear();
                TableWhsItemsForDetail.AcceptChanges();

                string XMLWhs = oDataTableWhs.GetAsXML();
                XMLWhs = XMLWhs.Replace("<Rows/></DataTable>", "");
                SbuilderWhs.Append(XMLWhs);
                SbuilderWhs.Append("<Rows>");

                oRecordSet.DoQuery(query);
                decimal qty = 0;
                decimal totalQty = 0;
                decimal totalCost = 0;
                int row = 0;
                int rowDetail = 0;
                string LocCode = "";
                string tmpLocCode = "";

                while (!oRecordSet.EoF)
                {
                    qty = Convert.ToDecimal(oRecordSet.Fields.Item("Qty").Value);

                    if (qty > 0)
                    {
                        tmpLocCode = oRecordSet.Fields.Item("LocCode").Value;
                        rowDetail++;
                        if (LocCode != tmpLocCode)
                        {
                            row++;
                            if (row > 1)
                            {
                                //წინა სტრიქონში ჯამები და დახურვა
                                SbuilderWhs = CommonFunctions.AddCellXML(SbuilderWhs, "Position", (rowDetail - 1).ToString());
                                SbuilderWhs = CommonFunctions.AddCellXML(SbuilderWhs, "Qty", FormsB1.ConvertDecimalToString(totalQty));
                                SbuilderWhs = CommonFunctions.AddCellXML(SbuilderWhs, "Cost", FormsB1.ConvertDecimalToString(totalCost));
                                SbuilderWhs.Append("</Row>");

                                totalQty = 0;
                                totalCost = 0;
                                rowDetail = 1;
                            }

                            SbuilderWhs.Append("<Row>");
                            SbuilderWhs = CommonFunctions.AddCellXML(SbuilderWhs, "ChkBx", "N");
                            SbuilderWhs = CommonFunctions.AddCellXML(SbuilderWhs, "LineID", row.ToString());
                            SbuilderWhs = CommonFunctions.AddCellXML(SbuilderWhs, "WhsFr", tmpLocCode);
                            SbuilderWhs = CommonFunctions.AddCellXML(SbuilderWhs, "WhseTypeFr", oRecordSet.Fields.Item("WhseTypeFr").Value);
                            SbuilderWhs = CommonFunctions.AddCellXML(SbuilderWhs, "WhsTo", whsTo);
                            SbuilderWhs = CommonFunctions.AddCellXML(SbuilderWhs, "PrjFr", oRecordSet.Fields.Item("PrjFr").Value);
                            SbuilderWhs = CommonFunctions.AddCellXML(SbuilderWhs, "PrjTo", prjTo);
                        }

                        totalQty = totalQty + qty;
                        totalCost = totalCost + Convert.ToDecimal(oRecordSet.Fields.Item("Cost").Value);

                        DataRow dataRow = TableWhsItemsForDetail.Rows.Add();
                        dataRow["LineIDWhsTable"] = row;
                        dataRow["LineID"] = rowDetail;
                        dataRow["ChkBx"] = 'Y';
                        dataRow["WhsFr"] = tmpLocCode;
                        dataRow["WhseTypeFr"] = oRecordSet.Fields.Item("WhseTypeFr").Value;
                        dataRow["WhsTo"] = whsTo;
                        dataRow["PrjFr"] = oRecordSet.Fields.Item("PrjFr").Value;
                        dataRow["PrjTo"] = prjTo;
                        dataRow["ItemCode"] = oRecordSet.Fields.Item("ItemCode").Value;
                        dataRow["ItemName"] = oRecordSet.Fields.Item("ItemName").Value;
                        dataRow["UomCode"] = "";
                        dataRow["InStock"] = qty;
                        dataRow["Qty"] = qty;
                        dataRow["Cost"] = oRecordSet.Fields.Item("Cost").Value;
                        dataRow["CostInStock"] = oRecordSet.Fields.Item("Cost").Value;
                    }
                    oRecordSet.MoveNext();
                    LocCode = tmpLocCode;
                }

                //ბოლო სტრიქონში ჯამები და დახურვა
                SbuilderWhs = CommonFunctions.AddCellXML(SbuilderWhs, "Position", rowDetail.ToString());
                SbuilderWhs = CommonFunctions.AddCellXML(SbuilderWhs, "Qty", FormsB1.ConvertDecimalToString(totalQty));
                SbuilderWhs = CommonFunctions.AddCellXML(SbuilderWhs, "Cost", FormsB1.ConvertDecimalToString(totalCost));
                SbuilderWhs.Append("</Row>");

                SbuilderWhs.Append("</Rows>");
                SbuilderWhs.Append("</DataTable>");

                XMLWhs = SbuilderWhs.ToString();
                oDataTableWhs.LoadFromXML(XMLWhs);

                TableWhsItemsForDetail.AcceptChanges();

                oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
                oGrid.DataTable = oDataTableWhs;
                SAPbouiCOM.GridColumns oColumns = oGrid.Columns;

                SAPbouiCOM.GridColumn oColumn = oColumns.Item("ChkBx");
                oColumn.Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
                oColumn.TitleObject.Caption = "";
                oColumn.Width = 20;

                SAPbouiCOM.EditTextColumn oEditTextColumn = (SAPbouiCOM.EditTextColumn)oColumns.Item("LineID");
                oEditTextColumn.TitleObject.Caption = "#";
                oEditTextColumn.Editable = false;
                oColumn.Width = 20;

                string objType = "64";
                oEditTextColumn = (SAPbouiCOM.EditTextColumn)oColumns.Item("WhsFr");
                oEditTextColumn.TitleObject.Caption = BDOSResources.getTranslate("WhsFr");
                oEditTextColumn.Editable = false;
                oEditTextColumn.LinkedObjectType = objType;

                oEditTextColumn = (SAPbouiCOM.EditTextColumn)oColumns.Item("WhseTypeFr");
                oEditTextColumn.TitleObject.Caption = BDOSResources.getTranslate("WhseTypeFrom");
                oEditTextColumn.Editable = false;

                oEditTextColumn = (SAPbouiCOM.EditTextColumn)oColumns.Item("WhsTo");
                oEditTextColumn.TitleObject.Caption = BDOSResources.getTranslate("WhsTo");
                oEditTextColumn.Editable = (string.IsNullOrEmpty(whsTo) ? true : false);
                oEditTextColumn.LinkedObjectType = objType;
                oEditTextColumn.ChooseFromListUID = "TableWhsTo_CFL";
                oEditTextColumn.ChooseFromListAlias = "WhsCode";

                oEditTextColumn = (SAPbouiCOM.EditTextColumn)oColumns.Item("PrjFr");
                oEditTextColumn.TitleObject.Caption = BDOSResources.getTranslate("ProjectFrom");
                oEditTextColumn.LinkedObjectType = "63";
                oEditTextColumn.Editable = false;

                oEditTextColumn = (SAPbouiCOM.EditTextColumn)oColumns.Item("PrjTo");
                oEditTextColumn.TitleObject.Caption = BDOSResources.getTranslate("ProjectTo");
                oEditTextColumn.LinkedObjectType = "63";
                oEditTextColumn.Editable = false;

                oEditTextColumn = (SAPbouiCOM.EditTextColumn)oColumns.Item("Position");
                oEditTextColumn.TitleObject.Caption = BDOSResources.getTranslate("Position");
                oEditTextColumn.Editable = false;

                oEditTextColumn = (SAPbouiCOM.EditTextColumn)oColumns.Item("Qty");
                oEditTextColumn.TitleObject.Caption = BDOSResources.getTranslate("Quantity");
                oEditTextColumn.Editable = false;

                oEditTextColumn = (SAPbouiCOM.EditTextColumn)oColumns.Item("Cost");
                oEditTextColumn.TitleObject.Caption = BDOSResources.getTranslate("Cost");
                oEditTextColumn.Editable = false;

                oEditTextColumn = (SAPbouiCOM.EditTextColumn)oColumns.Item("DocID");
                oEditTextColumn.TitleObject.Caption = BDOSResources.getTranslate("InventoryTransfer");
                oEditTextColumn.LinkedObjectType = "67";
                oEditTextColumn.Editable = false;

                oColumns.Item("RowsHeader").Visible = false;

                setVisibleFormItems(oForm, out errorText);
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                oRecordSet = null;
                oForm.Freeze(false);
                GC.Collect();
            }
        }

        public static void fillDetailTable(SAPbouiCOM.Form oForm, string LineIDWhsTable, string whsFr, string whseTypeFr, string whsTo, string docType, out string errorText)
        {
            errorText = null;

            oForm.Freeze(true);

            try
            {
                SAPbouiCOM.Grid oGrid = ((SAPbouiCOM.Grid)(oForm.Items.Item("WhsTblDt").Specific));
                SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("WhsTblDt");
                oDataTable.Rows.Clear();
                StringBuilder SbuilderWhs = new StringBuilder();

                string XMLWhs = oDataTable.GetAsXML();
                XMLWhs = XMLWhs.Replace("<Rows/></DataTable>", "");
                SbuilderWhs.Append(XMLWhs);
                SbuilderWhs.Append("<Rows>");

                string expression = "LineIDWhsTable = '" + LineIDWhsTable + "'";
                DataRow[] foundRows = TableWhsItemsForDetail.Select(expression);

                string tmpWhsTo;
                string tmpPrjTo;
                int row = 1;
                if (foundRows.Count() > 0)
                {
                    for (int i = 0; i < foundRows.Count(); i++)
                    {
                        if (string.IsNullOrEmpty(foundRows[i]["WhsTo"].ToString()))
                        {
                            tmpWhsTo = whsTo;
                            tmpPrjTo = CommonFunctions.getValue("OWHS", "U_BDOSPrjCod", "WhsCode", whsTo).ToString();
                        }
                        else
                        {
                            tmpWhsTo = foundRows[i]["WhsTo"].ToString();
                            tmpPrjTo = foundRows[i]["PrjTo"].ToString();
                        }

                        SbuilderWhs.Append("<Row>");
                        SbuilderWhs = CommonFunctions.AddCellXML(SbuilderWhs, "ChkBx", foundRows[i]["ChkBx"].ToString());
                        SbuilderWhs = CommonFunctions.AddCellXML(SbuilderWhs, "LineID", row.ToString());
                        SbuilderWhs = CommonFunctions.AddCellXML(SbuilderWhs, "WhsFr", whsFr);
                        SbuilderWhs = CommonFunctions.AddCellXML(SbuilderWhs, "WhseTypeFr", whseTypeFr);
                        SbuilderWhs = CommonFunctions.AddCellXML(SbuilderWhs, "WhsTo", tmpWhsTo);
                        SbuilderWhs = CommonFunctions.AddCellXML(SbuilderWhs, "PrjFr", foundRows[i]["PrjFr"].ToString());
                        SbuilderWhs = CommonFunctions.AddCellXML(SbuilderWhs, "PrjTo", tmpPrjTo);
                        SbuilderWhs = CommonFunctions.AddCellXML(SbuilderWhs, "ItemCode", foundRows[i]["ItemCode"].ToString());
                        SbuilderWhs = CommonFunctions.AddCellXML(SbuilderWhs, "ItemName", foundRows[i]["ItemName"].ToString());
                        SbuilderWhs = CommonFunctions.AddCellXML(SbuilderWhs, "UomCode", foundRows[i]["UomCode"].ToString());
                        SbuilderWhs = CommonFunctions.AddCellXML(SbuilderWhs, "InStock", FormsB1.ConvertDecimalToString(Convert.ToDecimal(foundRows[i]["InStock"].ToString())));
                        SbuilderWhs = CommonFunctions.AddCellXML(SbuilderWhs, "Qty", FormsB1.ConvertDecimalToString(Convert.ToDecimal(foundRows[i]["Qty"].ToString())));
                        SbuilderWhs = CommonFunctions.AddCellXML(SbuilderWhs, "Cost", FormsB1.ConvertDecimalToString(Convert.ToDecimal(foundRows[i]["Cost"].ToString())));
                        SbuilderWhs = CommonFunctions.AddCellXML(SbuilderWhs, "CostInStock", FormsB1.ConvertDecimalToString(Convert.ToDecimal(foundRows[i]["CostInStock"].ToString())));
                        SbuilderWhs.Append("</Row>");

                        row++;
                    }
                }

                SbuilderWhs.Append("</Rows>");
                SbuilderWhs.Append("</DataTable>");

                XMLWhs = SbuilderWhs.ToString();
                oDataTable.LoadFromXML(XMLWhs);

                oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
                oGrid.DataTable = oDataTable;
                SAPbouiCOM.GridColumns oColumns = oGrid.Columns;

                SAPbouiCOM.GridColumn oColumn = oColumns.Item("ChkBx");
                oColumn.Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
                oColumn.TitleObject.Caption = "";
                oColumn.Width = 20;

                SAPbouiCOM.EditTextColumn oEditTextColumn = (SAPbouiCOM.EditTextColumn)oColumns.Item("LineID");
                oEditTextColumn.TitleObject.Caption = "#";
                oEditTextColumn.Editable = false;
                oColumn.Width = 20;

                string objType = "64";
                oEditTextColumn = (SAPbouiCOM.EditTextColumn)oColumns.Item("WhsFr");
                oEditTextColumn.TitleObject.Caption = BDOSResources.getTranslate("WhsFr");
                oEditTextColumn.Editable = false;
                oEditTextColumn.LinkedObjectType = objType;

                oEditTextColumn = (SAPbouiCOM.EditTextColumn)oColumns.Item("WhseTypeFr");
                oEditTextColumn.TitleObject.Caption = BDOSResources.getTranslate("WhseTypeFrom");
                oEditTextColumn.Editable = false;

                oEditTextColumn = (SAPbouiCOM.EditTextColumn)oColumns.Item("WhsTo");
                oEditTextColumn.TitleObject.Caption = BDOSResources.getTranslate("WhsTo");
                oEditTextColumn.Editable = (docType == "TransferToProjectWarehouse" || docType == "ReturnToMainWarehouse" || docType == "TransferWithoutType" ? false : true);
                oEditTextColumn.LinkedObjectType = objType;
                oEditTextColumn.ChooseFromListUID = "DetailTableWhsTo_CFL";
                oEditTextColumn.ChooseFromListAlias = "WhsCode";

                oEditTextColumn = (SAPbouiCOM.EditTextColumn)oColumns.Item("ItemCode");
                oEditTextColumn.TitleObject.Caption = BDOSResources.getTranslate("ItemCode");
                oEditTextColumn.Editable = false;
                oEditTextColumn.LinkedObjectType = "4";

                oEditTextColumn = (SAPbouiCOM.EditTextColumn)oColumns.Item("ItemName");
                oEditTextColumn.TitleObject.Caption = BDOSResources.getTranslate("ItemName");
                oEditTextColumn.Editable = false;

                oEditTextColumn = (SAPbouiCOM.EditTextColumn)oColumns.Item("UomCode");
                oEditTextColumn.TitleObject.Caption = BDOSResources.getTranslate("UomCode");
                oEditTextColumn.Editable = false;
                oEditTextColumn.LinkedObjectType = "10000199";

                oEditTextColumn = (SAPbouiCOM.EditTextColumn)oColumns.Item("PrjFr");
                oEditTextColumn.TitleObject.Caption = BDOSResources.getTranslate("ProjectFrom");
                oEditTextColumn.LinkedObjectType = "63";
                oEditTextColumn.Editable = false;

                oEditTextColumn = (SAPbouiCOM.EditTextColumn)oColumns.Item("PrjTo");
                oEditTextColumn.TitleObject.Caption = BDOSResources.getTranslate("ProjectTo");
                oEditTextColumn.LinkedObjectType = "63";
                oEditTextColumn.Editable = false;

                oEditTextColumn = (SAPbouiCOM.EditTextColumn)oColumns.Item("InStock");
                oEditTextColumn.TitleObject.Caption = BDOSResources.getTranslate("InStock");
                oEditTextColumn.Editable = false;

                oEditTextColumn = (SAPbouiCOM.EditTextColumn)oColumns.Item("Qty");
                oEditTextColumn.TitleObject.Caption = BDOSResources.getTranslate("Quantity");

                oEditTextColumn = (SAPbouiCOM.EditTextColumn)oColumns.Item("Cost");
                oEditTextColumn.TitleObject.Caption = BDOSResources.getTranslate("Cost");
                oEditTextColumn.Editable = false; 

                oEditTextColumn = (SAPbouiCOM.EditTextColumn)oColumns.Item("CostInStock");
                oEditTextColumn.TitleObject.Caption = BDOSResources.getTranslate("CostInStock");
                oEditTextColumn.Editable = false;
                oEditTextColumn.Visible = false;
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
            finally
            {
                oForm.Freeze(false);
                GC.Collect();
            }
        }

        public static void updateDetailTable(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            //oForm.Freeze(true);
            oFormWizard.Freeze(true);

            try
            {
                SAPbouiCOM.Grid oGrid = ((SAPbouiCOM.Grid)(oForm.Items.Item("WhsTblDt").Specific));
                SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("WhsTblDt");

                string expression = "LineIDWhsTable = '" + currentLineIDWhsTable + "'";
                DataRow[] foundRows = TableWhsItemsForDetail.Select(expression);

                if (foundRows.Count() > 0)
                {
                    for (int i = 0; i < foundRows.Count(); i++)
                    {
                        foundRows[i].Delete();
                    }
                    TableWhsItemsForDetail.AcceptChanges();
                }

                decimal totalQty = 0;
                decimal totalCost = 0;
                int position = 0;
                decimal tmpQty;
                decimal tmpInStock;
                decimal tmpCost;

                for (int i = 0; i < oDataTable.Rows.Count; i++)
                {
                    tmpQty = Convert.ToDecimal(oDataTable.GetValue("Qty", i));
                    tmpInStock = Convert.ToDecimal(oDataTable.GetValue("InStock", i));
                    tmpCost = Convert.ToDecimal(oDataTable.GetValue("Cost", i));

                    DataRow dataRow = TableWhsItemsForDetail.Rows.Add();
                    dataRow["LineIDWhsTable"] = currentLineIDWhsTable;
                    dataRow["LineID"] = Convert.ToInt32(oDataTable.GetValue("LineID", i));
                    dataRow["ChkBx"] = oDataTable.GetValue("ChkBx", i).ToString();
                    dataRow["WhsFr"] = oDataTable.GetValue("WhsFr", i).ToString();
                    dataRow["WhsTo"] = oDataTable.GetValue("WhsTo", i).ToString();
                    dataRow["PrjFr"] = oDataTable.GetValue("PrjFr", i).ToString();
                    dataRow["PrjTo"] = oDataTable.GetValue("PrjTo", i).ToString();
                    dataRow["ItemCode"] = oDataTable.GetValue("ItemCode", i).ToString();
                    dataRow["ItemName"] = oDataTable.GetValue("ItemName", i).ToString();
                    dataRow["UomCode"] = oDataTable.GetValue("UomCode", i).ToString();
                    dataRow["InStock"] = tmpInStock;
                    dataRow["Qty"] = tmpQty;
                    dataRow["Cost"] = Convert.ToDecimal(oDataTable.GetValue("Cost", i));
                    dataRow["CostInStock"] = Convert.ToDecimal(oDataTable.GetValue("CostInStock", i));
                    
                    if (oDataTable.GetValue("ChkBx", i).ToString() == "Y")
                    {
                        totalQty = totalQty + tmpQty;
                        totalCost = totalCost + tmpCost;
                        position++;
                    }
                }

                TableWhsItemsForDetail.AcceptChanges();

                int whsLineID = Convert.ToInt32(currentLineIDWhsTable) - 1;
                SAPbouiCOM.DataTable oDataTableWHS = oFormWizard.DataSources.DataTables.Item("WhsTable");
                oDataTableWHS.SetValue("ChkBx", whsLineID, "Y");
                oDataTableWHS.SetValue("Position", whsLineID, position);
                oDataTableWHS.SetValue("Qty", whsLineID, Convert.ToDouble(totalQty, CultureInfo.InvariantCulture));
                oDataTableWHS.SetValue("Cost", whsLineID, Convert.ToDouble(totalCost, CultureInfo.InvariantCulture));

                //oFormWizard.Update();
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
            finally
            {
                //oForm.Freeze(false);
                oFormWizard.Freeze(false);
                GC.Collect();
            }
        }

        public static void findNextRow(SAPbouiCOM.Form oForm, string tableName, string codeColName, string nameColName, bool findNext, out string errorText)
        {
            errorText = null;

            SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item(tableName);
            SAPbouiCOM.Grid oGrid = ((SAPbouiCOM.Grid)(oForm.Items.Item(tableName).Specific));

            SAPbouiCOM.SelectedRows selectedRows = oGrid.Rows.SelectedRows;
            int rowIndex = 0;
            if (selectedRows.Count > 0)
            {
                rowIndex = selectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder);
            }
            string findNextVal = oForm.Items.Item("FindNextE").Specific.Value;

            string tmpCode;
            string tmpName;
            for (int i = 0; i < oDataTable.Rows.Count; i++)
            {
                tmpCode = oDataTable.GetValue(codeColName, i);
                tmpName = oDataTable.GetValue(nameColName, i);

                if ((tmpCode.Contains(findNextVal) || tmpName.Contains(findNextVal)) && (!findNext || rowIndex < i))
                {
                    selectedRows.Clear();
                    selectedRows.Add(i);
                    break;
                }
            }
        }

        public static void splitDetailRow(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            //oForm.Freeze(true);

            decimal newQty = FormsB1.cleanStringOfNonDigits(oForm.Items.Item("newQty").Specific.Value);

            SAPbouiCOM.DataTable oDataTable = oFormDetailWizard.DataSources.DataTables.Item("WhsTblDt");
            SAPbouiCOM.Grid oGrid = ((SAPbouiCOM.Grid)(oFormDetailWizard.Items.Item("WhsTblDt").Specific));

            SAPbouiCOM.SelectedRows selectedRows = oGrid.Rows.SelectedRows;
            int rowIndex = selectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder);
            int dTableRow = oGrid.GetDataTableRowIndex(rowIndex);

            decimal inStock = Convert.ToDecimal(oDataTable.GetValue("InStock", rowIndex));
            if (inStock <= newQty)
            {
                Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("QuantityShouldBeLessThan") + ": " + inStock, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            else
            {
                oFormDetailWizard.Freeze(true);

                oDataTable.Rows.Add();
                int i = oDataTable.Rows.Count - 1;

                //oDataTable.SetValue("LineIDWhsTable", i, currentLineIDWhsTable);
                oDataTable.SetValue("ChkBx", i, oDataTable.GetValue("ChkBx", rowIndex));
                oDataTable.SetValue("LineID", i, oDataTable.Rows.Count);
                oDataTable.SetValue("WhsFr", i, oDataTable.GetValue("WhsFr", rowIndex));
                oDataTable.SetValue("WhsTo", i, oDataTable.GetValue("WhsTo", rowIndex));
                oDataTable.SetValue("PrjFr", i, oDataTable.GetValue("PrjFr", rowIndex));
                oDataTable.SetValue("PrjTo", i, oDataTable.GetValue("PrjTo", rowIndex));
                oDataTable.SetValue("ItemCode", i, oDataTable.GetValue("ItemCode", rowIndex));
                oDataTable.SetValue("ItemName", i, oDataTable.GetValue("ItemName", rowIndex));
                oDataTable.SetValue("UomCode", i, oDataTable.GetValue("UomCode", rowIndex));
                oDataTable.SetValue("InStock", i, Convert.ToDouble(newQty));
                oDataTable.SetValue("Qty", i, Convert.ToDouble(newQty));
                oDataTable.SetValue("Cost", i, Convert.ToDouble(oDataTable.GetValue("Cost", rowIndex), CultureInfo.InvariantCulture));
                oDataTable.SetValue("InStock", rowIndex, Convert.ToDouble(inStock - newQty, CultureInfo.InvariantCulture));
                oDataTable.SetValue("Qty", rowIndex, Convert.ToDouble(inStock - newQty, CultureInfo.InvariantCulture));

                oFormDetailWizard.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                oFormDetailWizard.Freeze(false);
                oForm.Close();
            }

            //oForm.Freeze(false);
        }

        private static void createStockTransferDocuments(SAPbouiCOM.Form oForm)
        {
            int answer = Program.uiApp.MessageBox(BDOSResources.getTranslate("CreateStockTransferDocuments") + "?", 1, BDOSResources.getTranslate("Yes"), BDOSResources.getTranslate("No"), "");

            if (answer == 2)
            {
                return;
            }

            string errorText;

            SAPbouiCOM.EditText oEditTextDate = (SAPbouiCOM.EditText)oForm.Items.Item("DateE").Specific;
            DateTime date = DateTime.ParseExact(oEditTextDate.Value, "yyyyMMdd", null);
            string headWhsTo;
            string lineWhsTo;

            SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("WhsTable");
            for (int i = 0; i < oDataTable.Rows.Count; i++)
            {
                if (oDataTable.GetValue("ChkBx", i) == "Y")
                {
                    string lineIDWhsTable = oDataTable.GetValue("LineID", i);

                    if (!string.IsNullOrEmpty(oDataTable.GetValue("DocID", i)))
                    {
                        Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("DocumentAlreadyCreated") + "! " + BDOSResources.getTranslate("TableRow") + ": " + lineIDWhsTable, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);                 
                        continue;
                    }

                    SAPbobsCOM.StockTransfer oStockTransfer = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer);
                    SAPbobsCOM.StockTransfer oStockTransferNew = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer);
              
                    string expression = "LineIDWhsTable = '" + lineIDWhsTable + "'";
                    DataRow[] foundRows = TableWhsItemsForDetail.Select(expression);

                    if (foundRows.Count() > 0)
                    {
                        string prjFr = oDataTable.GetValue("PrjFr", i);
                        headWhsTo = oDataTable.GetValue("WhsTo", i);

                        oStockTransfer.FromWarehouse = oDataTable.GetValue("WhsFr", i);
                        oStockTransfer.ToWarehouse = headWhsTo;
                        oStockTransfer.DocDate = date;
                        oStockTransfer.TaxDate = date;
                        oStockTransfer.UserFields.Fields.Item("U_BDOSFrPrj").Value = prjFr;

                        for (int j = 0; j < foundRows.Count(); j++)
                        {
                            if (foundRows[j]["ChkBx"].ToString() == "Y")
                            {
                                lineWhsTo = foundRows[j]["WhsTo"].ToString();

                                oStockTransfer.Lines.ItemCode = foundRows[j]["ItemCode"].ToString();
                                //oStockTransfer.Lines.UoMEntry = Convert.ToInt32(foundRows[j]["UomCode"]);
                                oStockTransfer.Lines.Quantity = Convert.ToDouble(foundRows[j]["Qty"], CultureInfo.InvariantCulture);
                                oStockTransfer.Lines.ProjectCode = foundRows[j]["PrjTo"].ToString();
                                oStockTransfer.Lines.FromWarehouseCode = foundRows[j]["WhsFr"].ToString();
                                oStockTransfer.Lines.WarehouseCode = String.IsNullOrEmpty(lineWhsTo) ? headWhsTo : lineWhsTo;

                                oStockTransfer.Lines.Add();
                            }
                        }

                        CommonFunctions.StartTransaction();

                        bool DocumentCreated = false;
                        int retvals = oStockTransfer.Add();
                        if (retvals == 0)
                        {
                            bool newDoc = oStockTransferNew.GetByKey(Convert.ToInt32(Program.oCompany.GetNewObjectKey()));
                            if (newDoc == true)
                            {
                                StockTransfer.UpdateJournalEntry(oStockTransferNew.DocEntry.ToString(), "67", prjFr, out errorText);
                                if (string.IsNullOrEmpty(errorText))
                                {
                                    CommonFunctions.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);

                                    DocumentCreated = true;
                                    oDataTable.SetValue("DocID", i, oStockTransferNew.DocEntry);
                                    Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("OperationCompletedSuccessfully") + "! " + BDOSResources.getTranslate("TableRow") + ": " + lineIDWhsTable, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                                }
                            }
                        }

                        if (DocumentCreated == false)
                        {
                            CommonFunctions.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);

                            int errorCode;
                            Program.oCompany.GetLastError(out errorCode, out errorText);
                            Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("OperationCompletedUnSuccessfully") + "! " + errorText + "! " + BDOSResources.getTranslate("TableRow") + ": " + lineIDWhsTable, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        }
                    }
                }
            }
        }
    }
}
