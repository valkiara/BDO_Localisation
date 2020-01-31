using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace BDO_Localisation_AddOn
{
    class BDOSFuelTransferWizard
    {
        public static void createForm()
        {
            string errorText;
            int formHeight = Program.uiApp.Desktop.Width;
            int formWidth = Program.uiApp.Desktop.Width;

            //ფორმის აუცილებელი თვისებები
            Dictionary<string, object> formProperties = new Dictionary<string, object>();
            formProperties.Add("UniqueID", "BDOSFuelTransferWizard");
            formProperties.Add("BorderStyle", SAPbouiCOM.BoFormBorderStyle.fbs_Sizable);
            formProperties.Add("Title", BDOSResources.getTranslate("FuelTransferWizard"));
            formProperties.Add("ClientWidth", formWidth);
            formProperties.Add("ClientHeight", formHeight);

            SAPbouiCOM.Form oForm;
            bool newForm;
            bool formExist = FormsB1.createForm(formProperties, out oForm, out newForm, out errorText);

            if (errorText != null)
            {
                return;
            }

            if (formExist)
            {
                if (newForm)
                {
                    Dictionary<string, object> formItems;
                    string itemName;

                    int left_s = 6;
                    int left_e = 180;
                    int height = 15;
                    int top = 10;
                    int width_s = 160;
                    int width_e = 140;

                    int left_s2 = 400;
                    int left_e2 = left_s2 + 200;
                    int top2 = 10;

                    formItems = new Dictionary<string, object>();
                    itemName = "ForFilterS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", width_s * 2);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("DataForFilter"));
                    formItems.Add("TextStyle", 4);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    top += height + 7;
                    left_s = left_s + 10;

                    FormsB1.addChooseFromList(oForm, false, "4", "ItemCodeCFL"); //Items

                    formItems = new Dictionary<string, object>();
                    itemName = "ItemCodeS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("AssetCode"));
                    formItems.Add("LinkTo", "ItemCodeE");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "ItemCodeE"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("Length", 50);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("Left", left_e);
                    formItems.Add("Width", width_e);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("ChooseFromListUID", "ItemCodeCFL");
                    formItems.Add("ChooseFromListAlias", "ItemCode");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "ItemCodeLB"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    formItems.Add("Left", left_e - 20);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("LinkTo", "ItemCodeE");
                    formItems.Add("LinkedObjectType", "4");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    top += height + 1;

                    FormsB1.addChooseFromList(oForm, false, "171", "EmpIDCFL"); //Employees

                    formItems = new Dictionary<string, object>();
                    itemName = "EmpIDS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("Employee"));
                    formItems.Add("LinkTo", "EmpIDE");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "EmpIDE"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("Length", 50);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("Left", left_e);
                    formItems.Add("Width", width_e);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("ChooseFromListUID", "EmpIDCFL");
                    formItems.Add("ChooseFromListAlias", "empID");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "EmpIDLB"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    formItems.Add("Left", left_e - 20);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("LinkTo", "EmpIDE");
                    formItems.Add("LinkedObjectType", "171"); //Employees

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    top += height + 1;

                    FormsB1.addChooseFromList(oForm, false, "UDO_F_BDOSFUTP_D", "FuelTypeCodeCFL"); //Fuel Types

                    formItems = new Dictionary<string, object>();
                    itemName = "FuTpCodeS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("FuelType"));
                    formItems.Add("LinkTo", "FuTpCodeE");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "FuTpCodeE"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("Length", 50);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("Left", left_e);
                    formItems.Add("Width", width_e);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("ChooseFromListUID", "FuelTypeCodeCFL");
                    formItems.Add("ChooseFromListAlias", "Code");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "FuTpCodeLB"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    formItems.Add("Left", left_e - 20);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("LinkTo", "FuTpCodeE");
                    formItems.Add("LinkedObjectType", "UDO_F_BDOSFUTP_D");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    top += height + 1;

                    FormsB1.addChooseFromList(oForm, false, "4", "FuelCodeCFL"); //Items

                    formItems = new Dictionary<string, object>();
                    itemName = "FuelCodeS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("FuelCode"));
                    formItems.Add("LinkTo", "FuelCodeE");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "FuelCodeE"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("Length", 50);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("Left", left_e);
                    formItems.Add("Width", width_e);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("ChooseFromListUID", "FuelCodeCFL");
                    formItems.Add("ChooseFromListAlias", "ItemCode");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "FuelCodeLB"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    formItems.Add("Left", left_e - 20);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("LinkTo", "FuelCodeE");
                    formItems.Add("LinkedObjectType", "4");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    top += height + 1;

                    formItems = new Dictionary<string, object>();
                    itemName = "FuGroupS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("FuelGroup"));
                    formItems.Add("LinkTo", "FuGroupCB");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    Dictionary<string, string> listValidValuesItemGroups = getItemGroupsList();

                    formItems = new Dictionary<string, object>();
                    itemName = "FuGroupCB";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("Length", 20);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                    formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("Left", left_e);
                    formItems.Add("Width", width_e);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("ValidValues", listValidValuesItemGroups);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    //----------------------------------------------------------------------------------------------------------

                    formItems = new Dictionary<string, object>();
                    itemName = "ForCreateS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s2);
                    formItems.Add("Width", width_s * 2);
                    formItems.Add("Top", top2);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("DataForCreateDocument"));
                    formItems.Add("TextStyle", 4);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    top2 += height + 7;
                    left_s2 = left_s2 + 10;

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

                    DateTime endDate = DateTime.Today;
                    string endDateTxt = endDate.ToString("yyyyMMdd");

                    formItems = new Dictionary<string, object>();
                    itemName = "DocDateE";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_DATE);
                    formItems.Add("Length", 1);
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Left", left_e2);
                    formItems.Add("Width", width_e);
                    formItems.Add("Top", top2);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("ValueEx", endDateTxt);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    top2 += height + 1;

                    formItems = new Dictionary<string, object>();
                    itemName = "ReturnCH"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("Length", 1);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
                    formItems.Add("Width", width_e);
                    formItems.Add("Left", left_e2);
                    formItems.Add("Top", top2);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("Return"));
                    formItems.Add("ValOff", "N");
                    formItems.Add("ValOn", "Y");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    top2 += height + 1;

                    FormsB1.addChooseFromList(oForm, false, "64", "WarehouseFromCodeCFL"); //Warehouses

                    formItems = new Dictionary<string, object>();
                    itemName = "WhsFromS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s2);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top2);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("FromWarehouse"));
                    formItems.Add("LinkTo", "WhsFromE");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "WhsFromE"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("Length", 50);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("Left", left_e2);
                    formItems.Add("Width", width_e);
                    formItems.Add("Top", top2);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("ChooseFromListUID", "WarehouseFromCodeCFL");
                    formItems.Add("ChooseFromListAlias", "WhsCode");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "WhsFromLB"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    formItems.Add("Left", left_e2 - 20);
                    formItems.Add("Top", top2);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("LinkTo", "WhsFromE");
                    formItems.Add("LinkedObjectType", "64"); //Warehouses

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    top2 += height + 1;

                    FormsB1.addChooseFromList(oForm, false, "64", "WarehouseToCodeCFL"); //Warehouses

                    formItems = new Dictionary<string, object>();
                    itemName = "WhsToS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s2);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top2);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("ToWarehouse"));
                    formItems.Add("LinkTo", "WhsToE");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "WhsToE"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("Length", 50);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("Left", left_e2);
                    formItems.Add("Width", width_e);
                    formItems.Add("Top", top2);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("ChooseFromListUID", "WarehouseToCodeCFL");
                    formItems.Add("ChooseFromListAlias", "WhsCode");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "WhsToLB"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    formItems.Add("Left", left_e2 - 20);
                    formItems.Add("Top", top2);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("LinkTo", "WhsToE");
                    formItems.Add("LinkedObjectType", "64"); //Warehouses

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }
                    top2 += height + 1;
                    FormsB1.addChooseFromList(oForm, false, "63", "ProjectFromCodeCFL"); //Project Codes

                    formItems = new Dictionary<string, object>();
                    itemName = "PrjFromS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s2);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top2);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("FromProject"));
                    formItems.Add("LinkTo", "PrjFromE");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "PrjFromE"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("Length", 50);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("Left", left_e2);
                    formItems.Add("Width", width_e);
                    formItems.Add("Top", top2);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("ChooseFromListUID", "ProjectFromCodeCFL");
                    formItems.Add("ChooseFromListAlias", "PrjCode");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "PrjFromLB"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    formItems.Add("Left", left_e2 - 20);
                    formItems.Add("Top", top2);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("LinkTo", "PrjFromE");
                    formItems.Add("LinkedObjectType", "63");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    top2 += height + 1;

                    FormsB1.addChooseFromList(oForm, false, "63", "ProjectToCodeCFL"); //Project Codes

                    formItems = new Dictionary<string, object>();
                    itemName = "PrjToS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s2);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top2);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("ToProject"));
                    formItems.Add("LinkTo", "PrjToE");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "PrjToE"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("Length", 50);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("Left", left_e2);
                    formItems.Add("Width", width_e);
                    formItems.Add("Top", top2);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("ChooseFromListUID", "ProjectToCodeCFL");
                    formItems.Add("ChooseFromListAlias", "PrjCode");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "PrjToLB"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    formItems.Add("Left", left_e2 - 20);
                    formItems.Add("Top", top2);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("LinkTo", "PrjToE");
                    formItems.Add("LinkedObjectType", "63");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    top += height + 1;

                    //----------------------------------------------------------------------------------------------------------

                    top = top + 2 * height + 1;

                    itemName = "checkB";
                    formItems = new Dictionary<string, object>();
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", 19);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Image", "HANA_CHECKBOX_CH");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    left_s = left_s + 20;

                    itemName = "unCheckB";
                    formItems = new Dictionary<string, object>();
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", 19);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Image", "HANA_CHECKBOX_UH");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    left_s = left_s + 20;

                    itemName = "fillB";
                    formItems = new Dictionary<string, object>();
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", 65);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("Fill"));

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    itemName = "createDocB";
                    formItems = new Dictionary<string, object>();
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left_s + 65 + 2);
                    formItems.Add("Width", 65 * 2);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("CreateDocuments"));

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    top = top + height + 1;
                    left_s = 6;

                    formItems = new Dictionary<string, object>();
                    itemName = "AssetMTR"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", oForm.ClientWidth);
                    formItems.Add("Top", top);
                    formItems.Add("Height", (oForm.ClientHeight - 25 - top));
                    formItems.Add("UID", itemName);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("AffectsFormMode", false);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("AssetMTR").Specific));
                    oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
                    SAPbouiCOM.Columns oColumns = oMatrix.Columns;
                    SAPbouiCOM.Column oColumn;

                    SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Add("AssetMTR");

                    oDataTable.Columns.Add("CheckBox", SAPbouiCOM.BoFieldsType.ft_Text, 1);
                    oDataTable.Columns.Add("LineNum", SAPbouiCOM.BoFieldsType.ft_Integer, 50);
                    oDataTable.Columns.Add("DocEntryST", SAPbouiCOM.BoFieldsType.ft_Integer, 11);
                    oDataTable.Columns.Add("ItemCode", SAPbouiCOM.BoFieldsType.ft_Text, 50);
                    oDataTable.Columns.Add("ItemName", SAPbouiCOM.BoFieldsType.ft_Text, 100);
                    oDataTable.Columns.Add("IUoMEntry", SAPbouiCOM.BoFieldsType.ft_Integer, 11);
                    oDataTable.Columns.Add("InvntryUom", SAPbouiCOM.BoFieldsType.ft_Text, 100);
                    oDataTable.Columns.Add("AssetClass", SAPbouiCOM.BoFieldsType.ft_Text, 20);
                    oDataTable.Columns.Add("FuTpCode", SAPbouiCOM.BoFieldsType.ft_Text, 50);
                    oDataTable.Columns.Add("FuelCode", SAPbouiCOM.BoFieldsType.ft_Text, 50);
                    oDataTable.Columns.Add("FuelName", SAPbouiCOM.BoFieldsType.ft_Text, 100);
                    oDataTable.Columns.Add("FuelUom", SAPbouiCOM.BoFieldsType.ft_Text, 100);
                    oDataTable.Columns.Add("FuGrpCod", SAPbouiCOM.BoFieldsType.ft_Integer, 6);
                    oDataTable.Columns.Add("FuGrpNam", SAPbouiCOM.BoFieldsType.ft_Text, 20);
                    oDataTable.Columns.Add("Employee", SAPbouiCOM.BoFieldsType.ft_Integer, 11);
                    oDataTable.Columns.Add("FirstName", SAPbouiCOM.BoFieldsType.ft_Text, 50);
                    oDataTable.Columns.Add("LastName", SAPbouiCOM.BoFieldsType.ft_Text, 50);
                    oDataTable.Columns.Add("Quantity", SAPbouiCOM.BoFieldsType.ft_Quantity);

                    for (int i = 1; i <= 5; i++)
                    {
                        oDataTable.Columns.Add("Dimension" + i, SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 50);
                    }

                    string UID = "AssetMTR";
                    SAPbouiCOM.LinkedButton oLink;

                    oColumn = oColumns.Add("LineNum", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = "#";
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "LineNum");

                    oColumn = oColumns.Add("CheckBox", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
                    oColumn.TitleObject.Caption = "";
                    oColumn.Editable = true;
                    oColumn.ValOff = "N";
                    oColumn.ValOn = "Y";
                    oColumn.DataBind.Bind(UID, "CheckBox");

                    oColumn = oColumns.Add("DocEntryST", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("DocEntry") + " (" + BDOSResources.getTranslate("InventoryTransfer") + ")";
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "DocEntryST");
                    oLink = oColumn.ExtendedObject;
                    oLink.LinkedObjectType = "67"; //Inventory Transfer

                    oColumn = oColumns.Add("ItemCode", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("AssetCode");
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "ItemCode");
                    oLink = oColumn.ExtendedObject;
                    oLink.LinkedObjectType = "4"; //Items

                    oColumn = oColumns.Add("ItemName", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("AssetName");
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "ItemName");

                    oColumn = oColumns.Add("IUoMEntry", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("UomEntry");
                    oColumn.Visible = false;
                    oColumn.DataBind.Bind(UID, "IUoMEntry");
                    oLink = oColumn.ExtendedObject;
                    oLink.LinkedObjectType = "10000199"; //UoM Master Data

                    oColumn = oColumns.Add("InvntryUom", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("UomCode");
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "InvntryUom");

                    oColumn = oColumns.Add("AssetClass", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("AssetClass");
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "AssetClass");
                    oLink = oColumn.ExtendedObject;
                    oLink.LinkedObjectType = "1470000032"; //Asset Classes

                    oColumn = oColumns.Add("Employee", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("Employee");
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "Employee");
                    oLink = oColumn.ExtendedObject;
                    oLink.LinkedObjectType = "171"; //Employees

                    oColumn = oColumns.Add("FirstName", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("FirstName");
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "FirstName");

                    oColumn = oColumns.Add("LastName", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("LastName");
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "LastName");

                    oColumn = oColumns.Add("FuGrpNam", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("FuelGroupName");
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "FuGrpNam");

                    oColumn = oColumns.Add("FuTpCode", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("FuelType");
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "FuTpCode");
                    oLink = oColumn.ExtendedObject;
                    oLink.LinkedObjectType = "UDO_F_BDOSFUTP_D"; //Fuel Types

                    oColumn = oColumns.Add("FuelCode", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("Fuel");
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "FuelCode");
                    oLink = oColumn.ExtendedObject;
                    oLink.LinkedObjectType = "4"; //Items

                    oColumn = oColumns.Add("FuelName", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("FuelDescription");
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "FuelName");

                    oColumn = oColumns.Add("FuelUom", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("FuelUomCode");
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "FuelUom");

                    oColumn = oColumns.Add("FuGrpCod", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("FuelGroupCode");
                    oColumn.Editable = false;
                    oColumn.Visible = false;
                    oColumn.DataBind.Bind(UID, "FuGrpCod");
                    oLink = oColumn.ExtendedObject;
                    oLink.LinkedObjectType = "1470000046"; //Asset Groups

                    oColumn = oColumns.Add("Quantity", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("Quantity");
                    oColumn.DataBind.Bind(UID, "Quantity");

                    for (int i = 1; i <= 5; i++)
                    {
                        FormsB1.addChooseFromList(oForm, false, "62", "Dimension" + i + "CFL");

                        oColumn = oColumns.Add("Dimension" + i, SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                        oColumn.TitleObject.Caption = BDOSResources.getTranslate("Dimension" + i);
                        oColumn.DataBind.Bind(UID, "Dimension" + i);
                        oColumn.ChooseFromListUID = "Dimension" + i + "CFL";
                        oColumn.ChooseFromListAlias = "OcrCode";
                        oLink = oColumn.ExtendedObject;
                        oLink.LinkedObjectType = "62"; //Cost Rate
                    }
                }
                oForm.Visible = true;
                oForm.Select();
            }
        }

        public static void addMenus()
        {
            string enableFuelMng = (string)CommonFunctions.getOADM("U_BDOSEnbFlM");

            if (enableFuelMng == "Y")
            {
                try
                {
                    SAPbouiCOM.MenuItem fatherMenuItem = Program.uiApp.Menus.Item("3072");
                    // Add a pop-up menu item
                    SAPbouiCOM.MenuCreationParams oCreationPackage = (SAPbouiCOM.MenuCreationParams)Program.uiApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                    oCreationPackage.Checked = false;
                    oCreationPackage.Enabled = true;
                    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                    oCreationPackage.UniqueID = "BDOSFuelTransferWizard";
                    oCreationPackage.String = BDOSResources.getTranslate("FuelTransferWizard");
                    oCreationPackage.Position = fatherMenuItem.SubMenus.Count - 1;

                    SAPbouiCOM.MenuItem menuItem = fatherMenuItem.SubMenus.AddEx(oCreationPackage);
                }
                catch
                {
                    //Program.uiApp.MessageBox(ex.Message);
                }
            }
        }

        public static void uiApp_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_CLOSE && pVal.BeforeAction == true)
                {
                    int answer = Program.uiApp.MessageBox(BDOSResources.getTranslate("CloseFuelTransferWizard") + "?", 1, BDOSResources.getTranslate("Yes"), BDOSResources.getTranslate("No"), "");

                    if (answer != 1)
                    {
                        BubbleEvent = false;
                    }
                }

                else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK)
                {
                    if ((pVal.ItemUID == "checkB" || pVal.ItemUID == "unCheckB") && !pVal.BeforeAction)
                    {
                        checkUncheckMTR(oForm, pVal.ItemUID);
                    }
                    else if (pVal.ItemUID == "fillB" && !pVal.BeforeAction)
                    {
                        fillMTR(oForm);
                    }
                    else if (pVal.ItemUID == "createDocB" && !pVal.BeforeAction)
                    {
                        int answer = Program.uiApp.MessageBox(BDOSResources.getTranslate("CreateDocumentInventoryTransfer") + "?", 1, BDOSResources.getTranslate("Yes"), BDOSResources.getTranslate("No"), "");

                        if (answer == 1)
                        {
                            createDocuments(oForm);
                        }
                        return;
                    }
                }

                else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_RESIZE && !pVal.BeforeAction)
                {
                    resizeForm(oForm);
                }

                else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                {
                    SAPbouiCOM.IChooseFromListEvent oCFLEvento = (SAPbouiCOM.IChooseFromListEvent)pVal;
                    chooseFromList(oForm, pVal, oCFLEvento);
                }
            }
        }

        public static void resizeForm(SAPbouiCOM.Form oForm)
        {
            oForm.Freeze(true);
            try
            {
                SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("AssetMTR").Specific));
                int mtrWidth = oForm.ClientWidth;
                oForm.Items.Item("AssetMTR").Width = mtrWidth;
                int columnsCount = oMatrix.Columns.Count - 4;

                oMatrix.Columns.Item("LineNum").Width = 19;
                oMatrix.Columns.Item("CheckBox").Width = 19;
                mtrWidth -= 38;
                mtrWidth /= columnsCount;

                foreach (SAPbouiCOM.Column column in oMatrix.Columns)
                {
                    if (column.UniqueID == "LineNum" || column.UniqueID == "CheckBox" || column.UniqueID == "IUoMEntry" || column.UniqueID == "FuGrpCod")
                        continue;
                    column.Width = mtrWidth;
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

        public static void checkUncheckMTR(SAPbouiCOM.Form oForm, string checkOperation)
        {
            try
            {
                oForm.Freeze(true);

                SAPbouiCOM.CheckBox oCheckBox;
                SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("AssetMTR").Specific));

                for (int j = 1; j <= oMatrix.RowCount; j++)
                {
                    oCheckBox = oMatrix.Columns.Item("CheckBox").Cells.Item(j).Specific;

                    oCheckBox.Checked = (checkOperation == "checkB");
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                oForm.Freeze(false);
                GC.Collect();
            }
        }

        public static void chooseFromList(SAPbouiCOM.Form oForm, SAPbouiCOM.ItemEvent pVal, SAPbouiCOM.IChooseFromListEvent oCFLEvento)
        {
            try
            {
                oForm.Freeze(true);

                if (pVal.BeforeAction)
                {
                    if (oCFLEvento.ChooseFromListUID == "ProjectToCodeCFL" || oCFLEvento.ChooseFromListUID == "ProjectFromCodeCFL")
                    {
                        SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID);
                        SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();

                        SAPbouiCOM.Condition oCon = oCons.Add();
                        oCon.Alias = "Active";
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = "Y";
                        oCFL.SetConditions(oCons);
                    }
                    else if (oCFLEvento.ChooseFromListUID == "ItemCodeCFL")
                    {
                        SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID);
                        SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();

                        SAPbouiCOM.Condition oCon = oCons.Add();
                        oCon.Alias = "ItemType";
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = "F";
                        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                        oCon = oCons.Add();
                        oCon.Alias = "validFor";
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = "Y";
                        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                        SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        oRecordSet.DoQuery(@"SELECT ""Code"" FROM ""OACS"" WHERE ""U_BDOSVhcle"" = 'Y'");
                        int recordCount = oRecordSet.RecordCount;
                        int i = 0;

                        while (!oRecordSet.EoF)
                        {
                            string assetClassCode = oRecordSet.Fields.Item("Code").Value;
                            oCon = oCons.Add();
                            oCon.Alias = "AssetClass";
                            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                            oCon.CondVal = assetClassCode;
                            if (i == 0)
                                oCon.BracketOpenNum = 1;
                            if (i < recordCount - 1)
                            {
                                oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR;
                            }
                            if (i == recordCount - 1)
                            {
                                oCon.BracketCloseNum = 1;
                                oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_NONE;
                            }
                            i++;
                            oRecordSet.MoveNext();
                        }
                        oCFL.SetConditions(oCons);
                    }
                    else if (oCFLEvento.ChooseFromListUID == "FuelCodeCFL")
                    {
                        SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID);
                        SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();

                        SAPbouiCOM.Condition oCon = oCons.Add();
                        oCon.Alias = "InvntItem";
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = "Y";
                        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                        oCon = oCons.Add();
                        oCon.Alias = "validFor";
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = "Y";
                        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                        SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        oRecordSet.DoQuery(@"SELECT ""ItmsGrpCod"" FROM ""OITB"" WHERE ""U_BDOSFuel"" = 'Y'");
                        int recordCount = oRecordSet.RecordCount;
                        int i = 0;

                        while (!oRecordSet.EoF)
                        {
                            int itmsGrpCod = Convert.ToInt32(oRecordSet.Fields.Item("ItmsGrpCod").Value);
                            oCon = oCons.Add();
                            oCon.Alias = "ItmsGrpCod";
                            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                            oCon.CondVal = itmsGrpCod.ToString();
                            if (i == 0)
                                oCon.BracketOpenNum = 1;
                            if (i < recordCount - 1)
                            {
                                oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR;
                            }
                            if (i == recordCount - 1)
                            {
                                oCon.BracketCloseNum = 1;
                                oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_NONE;
                            }
                            i++;
                            oRecordSet.MoveNext();
                        }
                        oCFL.SetConditions(oCons);
                    }
                    else if (oCFLEvento.ChooseFromListUID == "WarehouseToCodeCFL")
                    {
                        SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID);
                        SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();

                        if (oForm.DataSources.UserDataSources.Item("ReturnCH").ValueEx == "N" || string.IsNullOrEmpty(oForm.DataSources.UserDataSources.Item("ReturnCH").ValueEx))
                        {
                            SAPbouiCOM.Condition oCon = oCons.Add();
                            oCon.Alias = "U_BDOSWhType";
                            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                            oCon.CondVal = "Fuel";
                        }

                        oCFL.SetConditions(oCons);
                    }
                    else if (oCFLEvento.ChooseFromListUID == "WarehouseFromCodeCFL")
                    {
                        SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID);
                        SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();

                        if (oForm.DataSources.UserDataSources.Item("ReturnCH").ValueEx == "Y")
                        {
                            SAPbouiCOM.Condition oCon = oCons.Add();
                            oCon.Alias = "U_BDOSWhType";
                            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                            oCon.CondVal = "Fuel";
                        }

                        oCFL.SetConditions(oCons);
                    }
                    else if (oCFLEvento.ChooseFromListUID.StartsWith("Dimension"))
                    {
                        SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID);
                        SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();

                        string dimCode = oCFLEvento.ChooseFromListUID.Substring("Dimension".Length, 1);

                        SAPbouiCOM.Condition oCon = oCons.Add();
                        oCon.Alias = "DimCode";
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = dimCode;
                        oCFL.SetConditions(oCons);
                    }
                }
                else
                {
                    SAPbouiCOM.DataTable oDataTable = oCFLEvento.SelectedObjects;

                    if (oDataTable != null)
                    {
                        if (oCFLEvento.ChooseFromListUID == "ProjectToCodeCFL")
                        {
                            string prjCode = oDataTable.GetValue("PrjCode", 0);
                            LanguageUtils.IgnoreErrors<string>(() => oForm.Items.Item("PrjToE").Specific.Value = prjCode);
                        }
                        else if (oCFLEvento.ChooseFromListUID == "ProjectFromCodeCFL")
                        {
                            string prjCode = oDataTable.GetValue("PrjCode", 0);
                            LanguageUtils.IgnoreErrors<string>(() => oForm.Items.Item("PrjFromE").Specific.Value = prjCode);
                        }

                        else if (oCFLEvento.ChooseFromListUID == "ItemCodeCFL")
                        {
                            string itemCode = oDataTable.GetValue("ItemCode", 0);
                            LanguageUtils.IgnoreErrors<string>(() => oForm.Items.Item("ItemCodeE").Specific.Value = itemCode);
                        }
                        else if (oCFLEvento.ChooseFromListUID == "FuelCodeCFL")
                        {
                            string fuelCode = oDataTable.GetValue("ItemCode", 0);
                            LanguageUtils.IgnoreErrors<string>(() => oForm.Items.Item("FuelCodeE").Specific.Value = fuelCode);
                        }
                        else if (oCFLEvento.ChooseFromListUID == "FuelTypeCodeCFL")
                        {
                            string fuTpCode = oDataTable.GetValue("Code", 0);
                            LanguageUtils.IgnoreErrors<string>(() => oForm.Items.Item("FuTpCodeE").Specific.Value = fuTpCode);
                        }
                        else if (oCFLEvento.ChooseFromListUID == "WarehouseToCodeCFL")
                        {
                            string whsCode = oDataTable.GetValue("WhsCode", 0);
                            string prjCode = oDataTable.GetValue("U_BDOSPrjCod", 0);
                            LanguageUtils.IgnoreErrors<string>(() => oForm.Items.Item("WhsToE").Specific.Value = whsCode);
                            LanguageUtils.IgnoreErrors<string>(() => oForm.Items.Item("PrjToE").Specific.Value = prjCode);
                        }
                        else if (oCFLEvento.ChooseFromListUID == "WarehouseFromCodeCFL")
                        {
                            string whsCode = oDataTable.GetValue("WhsCode", 0);
                            string prjCode = oDataTable.GetValue("U_BDOSPrjCod", 0);
                            LanguageUtils.IgnoreErrors<string>(() => oForm.Items.Item("WhsFromE").Specific.Value = whsCode);
                            LanguageUtils.IgnoreErrors<string>(() => oForm.Items.Item("PrjFromE").Specific.Value = prjCode);
                        }
                        else if (oCFLEvento.ChooseFromListUID == "EmpIDCFL")
                        {
                            int empID = oDataTable.GetValue("empID", 0);
                            LanguageUtils.IgnoreErrors<string>(() => oForm.Items.Item("EmpIDE").Specific.Value = empID.ToString());
                        }
                        else if (oCFLEvento.ChooseFromListUID.StartsWith("Dimension"))
                        {
                            string dimension = oDataTable.GetValue("OcrCode", 0);
                            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("AssetMTR").Specific;
                            LanguageUtils.IgnoreErrors<string>(() => oMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value = dimension);
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

        public static void fillMTR(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("AssetMTR");
            oDataTable.Rows.Clear();
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string empID = oForm.DataSources.UserDataSources.Item("EmpIDE").ValueEx;
            string fuTpCode = oForm.DataSources.UserDataSources.Item("FuTpCodeE").ValueEx;
            string fuGroup = oForm.DataSources.UserDataSources.Item("FuGroupCB").ValueEx;
            string itemCode = oForm.DataSources.UserDataSources.Item("ItemCodeE").ValueEx;
            string fuelCode = oForm.DataSources.UserDataSources.Item("FuelCodeE").ValueEx;

            StringBuilder query = new StringBuilder();
            query.Append("SELECT \n");
            query.Append("\"OITM\".\"ItemCode\", \n");
            query.Append("\"OITM\".\"ItemName\", \n");
            query.Append("\"OITM\".\"IUoMEntry\", \n");
            query.Append("\"OITM\".\"InvntryUom\", \n");
            query.Append("\"OITM\".\"AssetClass\", \n");
            query.Append("\"OITM\".\"U_BDOSFuTp\" AS \"FuTpCode\", \n");
            query.Append("\"@BDOSFUTP\".\"U_ItemCode\" AS \"FuelCode\", \n");
            query.Append("\"@BDOSFUTP\".\"U_ItemName\" AS \"FuelName\", \n");
            query.Append("\"@BDOSFUTP\".\"U_UomCode\" AS \"FuelUomCode\", \n");
            query.Append("\"@BDOSFUTP\".\"FuGrpCod\", \n");
            query.Append("\"@BDOSFUTP\".\"FuGrpNam\", \n");
            query.Append("\"OITM\".\"Employee\", \n");
            query.Append("\"OHEM\".\"firstName\" AS \"FirstName\", \n");
            query.Append("\"OHEM\".\"lastName\" AS \"LastName\", \n");
            query.Append("\"OPRC\".\"PrcCode\" \n");
            query.Append("FROM \"OITM\" \n");
            query.Append("INNER JOIN \"OACS\" ON \"OITM\".\"AssetClass\" = \"OACS\".\"Code\" \n");
            query.Append("INNER JOIN \n");
            query.Append("(SELECT \"@BDOSFUTP\".*, \n");
            query.Append("\"OITM\".\"ItmsGrpCod\" AS \"FuGrpCod\", \n");
            query.Append("\"OITB\".\"ItmsGrpNam\" AS \"FuGrpNam\" \n");
            query.Append("FROM \"@BDOSFUTP\" \n");
            query.Append("INNER JOIN \"OITM\" ON \"OITM\".\"ItemCode\" = \"@BDOSFUTP\".\"U_ItemCode\" \n");
            query.Append("INNER JOIN \"OITB\" ON \"OITM\".\"ItmsGrpCod\" = \"OITB\".\"ItmsGrpCod\") AS \"@BDOSFUTP\" \n");
            query.Append("ON \"OITM\".\"U_BDOSFuTp\" = \"@BDOSFUTP\".\"Code\" \n");
            query.Append("LEFT JOIN \"OHEM\" ON \"OITM\".\"Employee\" = \"OHEM\".\"empID\" \n");
            query.Append("INNER JOIN \"OPRC\" ON \"OITM\".\"ItemCode\" = \"OPRC\".\"U_BDOSFACode\" \n");
            query.Append("WHERE \"OITM\".\"ItemType\" = 'F' \n");
            query.Append("AND \"OITM\".\"validFor\" = 'Y' \n");
            query.Append("AND \"OACS\".\"U_BDOSVhcle\" = 'Y' \n");

            if (!string.IsNullOrEmpty(empID))
            {
                query.Append("AND \"OITM\".\"Employee\" = '" + empID + "' \n");
            }
            if (!string.IsNullOrEmpty(itemCode))
            {
                query.Append("AND \"OITM\".\"ItemCode\" = '" + itemCode + "' \n");
            }
            if (!string.IsNullOrEmpty(fuTpCode))
            {
                query.Append("AND \"OITM\".\"U_BDOSFuTp\" = '" + fuTpCode + "' \n");
            }
            if (!string.IsNullOrEmpty(fuelCode))
            {
                query.Append("AND \"@BDOSFUTP\".\"U_ItemCode\" = '" + fuelCode + "' \n");
            }
            if (!string.IsNullOrEmpty(fuGroup))
            {
                query.Append("AND \"@BDOSFUTP\".\"FuGrpCod\" = '" + fuGroup + "' \n");
            }

            query.Append("ORDER BY \"OITM\".\"ItemCode\"");

            oRecordSet.DoQuery(query.ToString());

            try
            {
                int rowIndex = 0;

                string dimensionNbr = (string)CommonFunctions.getOADM("U_BDOSFADim");

                while (!oRecordSet.EoF)
                {
                    oDataTable.Rows.Add();
                    oDataTable.SetValue("LineNum", rowIndex, rowIndex + 1);
                    oDataTable.SetValue("CheckBox", rowIndex, "N");
                    oDataTable.SetValue("ItemCode", rowIndex, oRecordSet.Fields.Item("ItemCode").Value);
                    oDataTable.SetValue("ItemName", rowIndex, oRecordSet.Fields.Item("ItemName").Value);
                    oDataTable.SetValue("IUoMEntry", rowIndex, oRecordSet.Fields.Item("IUoMEntry").Value);
                    oDataTable.SetValue("InvntryUom", rowIndex, oRecordSet.Fields.Item("InvntryUom").Value);
                    oDataTable.SetValue("AssetClass", rowIndex, oRecordSet.Fields.Item("AssetClass").Value);
                    oDataTable.SetValue("FuTpCode", rowIndex, oRecordSet.Fields.Item("FuTpCode").Value);
                    oDataTable.SetValue("FuelCode", rowIndex, oRecordSet.Fields.Item("FuelCode").Value);
                    oDataTable.SetValue("FuelName", rowIndex, oRecordSet.Fields.Item("FuelName").Value);
                    oDataTable.SetValue("FuelUom", rowIndex, oRecordSet.Fields.Item("FuelUomCode").Value);
                    oDataTable.SetValue("FuGrpCod", rowIndex, oRecordSet.Fields.Item("FuGrpCod").Value);
                    oDataTable.SetValue("FuGrpNam", rowIndex, oRecordSet.Fields.Item("FuGrpNam").Value);
                    oDataTable.SetValue("Employee", rowIndex, oRecordSet.Fields.Item("Employee").Value);
                    oDataTable.SetValue("FirstName", rowIndex, oRecordSet.Fields.Item("FirstName").Value);
                    oDataTable.SetValue("LastName", rowIndex, oRecordSet.Fields.Item("LastName").Value);
                    oDataTable.SetValue("Dimension" + dimensionNbr, rowIndex, oRecordSet.Fields.Item("PrcCode").Value);

                    oRecordSet.MoveNext();
                    rowIndex++;
                }

                oForm.Freeze(true);
                SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("AssetMTR").Specific));
                oMatrix.Clear();
                oMatrix.LoadFromDataSource();
                oMatrix.AutoResizeColumns();
                oForm.Update();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                oForm.Freeze(false);
                Marshal.ReleaseComObject(oRecordSet);
            }
        }

        public static void createDocuments(SAPbouiCOM.Form oForm)
        {
            string errorText;
            string docDateStr = oForm.DataSources.UserDataSources.Item("DocDateE").ValueEx;
            string whsToCode = oForm.DataSources.UserDataSources.Item("WhsToE").ValueEx;
            string whsFromCode = oForm.DataSources.UserDataSources.Item("WhsFromE").ValueEx;
            string prjToCode = oForm.DataSources.UserDataSources.Item("PrjToE").ValueEx;
            string prjFromCode = oForm.DataSources.UserDataSources.Item("PrjFromE").ValueEx;

            if (string.IsNullOrEmpty(docDateStr) || string.IsNullOrEmpty(whsToCode) || string.IsNullOrEmpty(whsFromCode))
            {
                errorText = BDOSResources.getTranslate("TheFollowingFieldsAreMandatory") + " : \"" + oForm.Items.Item("DocDateS").Specific.caption + "\", \"" + oForm.Items.Item("WhsFromS").Specific.caption + "\", \"" + oForm.Items.Item("WhsToS").Specific.caption + "\"";
                Program.uiApp.SetStatusBarMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                return;
            }

            DateTime docDate = Convert.ToDateTime(DateTime.ParseExact(docDateStr, "yyyyMMdd", CultureInfo.InvariantCulture));

            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("AssetMTR").Specific;
            oMatrix.FlushToDataSource();

            SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("AssetMTR");
            string checkBox = "N";
            string fuTpCode;

            SAPbobsCOM.StockTransfer oStockTransfer = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer);
            oStockTransfer.DocDate = docDate;
            oStockTransfer.TaxDate = docDate;
            oStockTransfer.FromWarehouse = whsFromCode;
            oStockTransfer.ToWarehouse = whsToCode;
            oStockTransfer.UserFields.Fields.Item("U_BDOSFrPrj").Value = prjFromCode;

            bool existLine = false;

            for (int i = 0; i < oDataTable.Rows.Count; i++)
            {
                checkBox = oDataTable.GetValue("CheckBox", i);
                fuTpCode = oDataTable.GetValue("FuTpCode", i);

                if (checkBox == "Y" && !string.IsNullOrEmpty(fuTpCode) && oDataTable.GetValue("DocEntryST", i) == 0 && oDataTable.GetValue("Quantity", i) > 0)
                {
                    oStockTransfer.Lines.ItemCode = oDataTable.GetValue("FuelCode", i);
                    oStockTransfer.Lines.ItemDescription = oDataTable.GetValue("FuelName", i);
                    oStockTransfer.Lines.FromWarehouseCode = whsFromCode;
                    oStockTransfer.Lines.WarehouseCode = whsToCode;
                    oStockTransfer.Lines.ProjectCode = prjToCode;
                    double quantity = Convert.ToDouble(oDataTable.GetValue("Quantity", i), CultureInfo.InvariantCulture);
                    oStockTransfer.Lines.Quantity = quantity;
                    oStockTransfer.Lines.DistributionRule = oDataTable.GetValue("Dimension1", i);
                    oStockTransfer.Lines.DistributionRule2 = oDataTable.GetValue("Dimension2", i);
                    oStockTransfer.Lines.DistributionRule3 = oDataTable.GetValue("Dimension3", i);
                    oStockTransfer.Lines.DistributionRule4 = oDataTable.GetValue("Dimension4", i);
                    oStockTransfer.Lines.DistributionRule5 = oDataTable.GetValue("Dimension5", i);

                    oStockTransfer.Lines.Add();
                    existLine = true;
                }
            }

            if (existLine)
            {
                CommonFunctions.StartTransaction();
                int resultCode = oStockTransfer.Add();
                if (resultCode != 0)
                {
                    int errCode;
                    string errMsg;
                    Program.oCompany.GetLastError(out errCode, out errMsg);
                    Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("DocumentNotCreated") + ". " + BDOSResources.getTranslate("ReasonIs") + ": " + errMsg, SAPbouiCOM.BoMessageTime.bmt_Short);
                }
                else
                {
                    int docEntryST;
                    bool newDoc = oStockTransfer.GetByKey(Convert.ToInt32(Program.oCompany.GetNewObjectKey()));
                    if (newDoc == true)
                    {
                        StockTransfer.UpdateJournalEntry(oStockTransfer.DocEntry.ToString(), "67", prjFromCode, out errorText);
                        if (string.IsNullOrEmpty(errorText))
                        {
                            CommonFunctions.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);

                            docEntryST = oStockTransfer.DocEntry;
                            Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("DocumentCreatedSuccesfully") + ": " + docEntryST, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);

                            for (int i = 0; i < oDataTable.Rows.Count; i++)
                            {
                                checkBox = oDataTable.GetValue("CheckBox", i);
                                if (docEntryST > 0 && checkBox == "Y" && oDataTable.GetValue("DocEntryST", i) == 0)
                                {
                                    oDataTable.SetValue("DocEntryST", i, docEntryST);
                                }
                            }
                        }
                    }
                    else
                    {
                        CommonFunctions.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                        Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("DocumentNotCreated"), SAPbouiCOM.BoMessageTime.bmt_Short);
                        //return;
                    }
                    Marshal.ReleaseComObject(oStockTransfer);
                }

                oForm.Freeze(true);
                oMatrix.Clear();
                oMatrix.LoadFromDataSource();
                oForm.Update();
                oForm.Freeze(false);
            }
            else
            {
                Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("DocumentNotCreated") + ". " + BDOSResources.getTranslate("ReasonIs") + ": " + BDOSResources.getTranslate("TheTableCanNotBeEmpty"), SAPbouiCOM.BoMessageTime.bmt_Short);
            }
            //else
            //{
            //    Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("DocumentAlreadyCreated") + "! " + BDOSResources.getTranslate("ToCreateNewDocumentPressFillButton"), SAPbouiCOM.BoMessageTime.bmt_Short);
            //}
        }

        public static Dictionary<string, string> getItemGroupsList()
        {
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                Dictionary<string, string> itemGroupsList = new Dictionary<string, string>();
                itemGroupsList.Add("", "");

                oRecordSet.DoQuery(@"SELECT ""ItmsGrpCod"", ""ItmsGrpNam"" FROM ""OITB"" WHERE ""Locked""='N'");
                while (!oRecordSet.EoF)
                {
                    itemGroupsList.Add(oRecordSet.Fields.Item("ItmsGrpCod").Value.ToString(), oRecordSet.Fields.Item("ItmsGrpNam").Value.ToString());
                    oRecordSet.MoveNext();
                }
                return itemGroupsList;
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
    }
}
