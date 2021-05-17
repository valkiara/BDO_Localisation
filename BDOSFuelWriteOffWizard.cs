using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace BDO_Localisation_AddOn
{
    class BDOSFuelWriteOffWizard
    {
        public static void createForm()
        {
            string errorText;
            int formHeight = Program.uiApp.Desktop.Width;
            int formWidth = Program.uiApp.Desktop.Width;

            //ფორმის აუცილებელი თვისებები
            Dictionary<string, object> formProperties = new Dictionary<string, object>();
            formProperties.Add("UniqueID", "BDOSFuelWOForm");
            formProperties.Add("BorderStyle", SAPbouiCOM.BoFormBorderStyle.fbs_Sizable);
            formProperties.Add("Title", BDOSResources.getTranslate("FuelWriteOffWizard"));
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

                    formItems = new Dictionary<string, object>();
                    itemName = "PeriodS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("Period"));
                    formItems.Add("LinkTo", "DateFromE");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    DateTime startDate = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);
                    string startDateTxt = startDate.ToString("yyyyMMdd");

                    DateTime endDate = DateTime.Today;
                    string endDateTxt = endDate.ToString("yyyyMMdd");

                    formItems = new Dictionary<string, object>();
                    itemName = "DateFromE";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_DATE);
                    formItems.Add("Length", 1);
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Left", left_e);
                    formItems.Add("Width", width_e / 2);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("ValueEx", startDateTxt);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "DateToE";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_DATE);
                    formItems.Add("Length", 1);
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Left", left_e + width_e / 2);
                    formItems.Add("Width", width_e / 2);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("ValueEx", endDateTxt);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    top += height + 1;

                    FormsB1.addChooseFromList(oForm, false, "63", "ProjectCodeCFL"); //Project Codes

                    formItems = new Dictionary<string, object>();
                    itemName = "PrjCodeS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("Project"));
                    formItems.Add("LinkTo", "PrjCodeE");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "PrjCodeE"; //10 characters
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
                    formItems.Add("ChooseFromListUID", "ProjectCodeCFL");
                    formItems.Add("ChooseFromListAlias", "PrjCode");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "PrjCodeLB"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    formItems.Add("Left", left_e - 20);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("LinkTo", "PrjCodeE");
                    formItems.Add("LinkedObjectType", "63");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    top += height + 1;

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

                    FormsB1.addChooseFromList(oForm, false, "1", "AccountCodeCFL"); //G/L Accounts

                    formItems = new Dictionary<string, object>();
                    itemName = "AccountS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s2);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top2);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("GLAccountCode"));
                    formItems.Add("LinkTo", "AccountE");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "AccountE"; //10 characters
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
                    formItems.Add("ChooseFromListUID", "AccountCodeCFL");
                    formItems.Add("ChooseFromListAlias", "AcctCode");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "AccountLB"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    formItems.Add("Left", left_e2 - 20);
                    formItems.Add("Top", top2);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("LinkTo", "AccountE");
                    formItems.Add("LinkedObjectType", "1");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    top2 += height + 1;

                    formItems = new Dictionary<string, object>();
                    itemName = "WhsCodeS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s2);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top2);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("FuelWarehouse"));
                    formItems.Add("LinkTo", "WhsCodeE");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "WhsCodeE"; //10 characters
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
                    formItems.Add("Enabled", false);
                    formItems.Add("ValueEx", (string)CommonFunctions.getOADM("U_BDOSFlWhs"));

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "WhsCodeLB"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    formItems.Add("Left", left_e2 - 20);
                    formItems.Add("Top", top2);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("LinkTo", "WhsCodeE");
                    formItems.Add("LinkedObjectType", "64"); //Warehouses

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }
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
                    oDataTable.Columns.Add("DocEntry", SAPbouiCOM.BoFieldsType.ft_Integer, 50);
                    oDataTable.Columns.Add("DocEntryGI", SAPbouiCOM.BoFieldsType.ft_Integer, 11);
                    oDataTable.Columns.Add("DocDate", SAPbouiCOM.BoFieldsType.ft_Date, 50);
                    oDataTable.Columns.Add("DateFrom", SAPbouiCOM.BoFieldsType.ft_Date, 50);
                    oDataTable.Columns.Add("DateTo", SAPbouiCOM.BoFieldsType.ft_Date, 50);
                    oDataTable.Columns.Add("PrjCode", SAPbouiCOM.BoFieldsType.ft_Text, 20);
                    oDataTable.Columns.Add("FuNrCode", SAPbouiCOM.BoFieldsType.ft_Text, 50);
                    oDataTable.Columns.Add("LineId", SAPbouiCOM.BoFieldsType.ft_Integer);
                    oDataTable.Columns.Add("ItemCode", SAPbouiCOM.BoFieldsType.ft_Text, 50);
                    oDataTable.Columns.Add("ItemName", SAPbouiCOM.BoFieldsType.ft_Text, 100);
                    oDataTable.Columns.Add("FuTpCode", SAPbouiCOM.BoFieldsType.ft_Text, 50);
                    oDataTable.Columns.Add("FuelCode", SAPbouiCOM.BoFieldsType.ft_Text, 50);
                    oDataTable.Columns.Add("FuelName", SAPbouiCOM.BoFieldsType.ft_Text, 100);
                    oDataTable.Columns.Add("FuUomEntry", SAPbouiCOM.BoFieldsType.ft_Integer, 11);
                    oDataTable.Columns.Add("FuUomCode", SAPbouiCOM.BoFieldsType.ft_Text, 20);
                    oDataTable.Columns.Add("FuPerKm", SAPbouiCOM.BoFieldsType.ft_Sum);
                    oDataTable.Columns.Add("FuPerHr", SAPbouiCOM.BoFieldsType.ft_Sum);
                    oDataTable.Columns.Add("OdmtrStart", SAPbouiCOM.BoFieldsType.ft_Sum);
                    oDataTable.Columns.Add("OdmtrEnd", SAPbouiCOM.BoFieldsType.ft_Sum);
                    oDataTable.Columns.Add("HrsWorked", SAPbouiCOM.BoFieldsType.ft_Sum);
                    oDataTable.Columns.Add("NormCn", SAPbouiCOM.BoFieldsType.ft_Sum);
                    oDataTable.Columns.Add("ActuallyCn", SAPbouiCOM.BoFieldsType.ft_Sum);

                    for (int i = 1; i <= 5; i++)
                    {
                        oDataTable.Columns.Add("Dimension" + i, SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 50);
                    }
                    oDataTable.Columns.Add("AcctCode", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 15);

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

                    oColumn = oColumns.Add("DocEntry", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("DocEntry");
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "DocEntry");
                    oLink = oColumn.ExtendedObject;
                    oLink.LinkedObjectType = "UDO_F_BDOSFUCN_D"; //Fuel Consumption Act

                    oColumn = oColumns.Add("DocEntryGI", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("DocEntry") + " (" + BDOSResources.getTranslate("GoodsIssue") + ")";
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "DocEntryGI");
                    oLink = oColumn.ExtendedObject;
                    oLink.LinkedObjectType = "60"; //Goods Issue

                    oColumn = oColumns.Add("DocDate", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("PostingDate");
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "DocDate");

                    oColumn = oColumns.Add("DateFrom", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("DateFrom");
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "DateFrom");

                    oColumn = oColumns.Add("DateTo", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("DateTo");
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "DateTo");

                    oColumn = oColumns.Add("PrjCode", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("Project");
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "PrjCode");
                    oLink = oColumn.ExtendedObject;
                    oLink.LinkedObjectType = "63"; //Project Codes

                    oColumn = oColumns.Add("FuNrCode", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("SpecificationOfFuelNorm");
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "FuNrCode");
                    oLink = oColumn.ExtendedObject;
                    oLink.LinkedObjectType = "UDO_F_BDOSFUNR_D"; //Specification of Fuel Norm

                    oColumn = oColumns.Add("LineId", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.Editable = false;
                    oColumn.Visible = false;
                    oColumn.DataBind.Bind(UID, "LineId");

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

                    oColumn = oColumns.Add("FuUomEntry", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("UomEntry");
                    oColumn.Visible = false;
                    oColumn.DataBind.Bind(UID, "FuUomEntry");
                    oLink = oColumn.ExtendedObject;
                    oLink.LinkedObjectType = "10000199"; //UoM Master Data

                    oColumn = oColumns.Add("FuUomCode", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("UomCode");
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "FuUomCode");

                    oColumn = oColumns.Add("FuPerKm", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("PerKm");
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "FuPerKm");

                    oColumn = oColumns.Add("FuPerHr", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("PerHr");
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "FuPerHr");

                    oColumn = oColumns.Add("OdmtrStart", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("StartingValueOfOdometer");
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "OdmtrStart");

                    oColumn = oColumns.Add("OdmtrEnd", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("EndingValueOfOdometer");
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "OdmtrEnd");

                    oColumn = oColumns.Add("HrsWorked", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("HoursWorked");
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "HrsWorked");

                    oColumn = oColumns.Add("NormCn", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("NormConsumption");
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "NormCn");

                    oColumn = oColumns.Add("ActuallyCn", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("ActuallyConsumption");
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "ActuallyCn");

                    for (int i = 1; i <= 5; i++)
                    {
                        oColumn = oColumns.Add("Dimension" + i, SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                        oColumn.TitleObject.Caption = BDOSResources.getTranslate("Dimension" + i);
                        oColumn.Editable = false;
                        oColumn.DataBind.Bind(UID, "Dimension" + i);
                        oLink = oColumn.ExtendedObject;
                        oLink.LinkedObjectType = "62"; //Cost Rate
                    }

                    FormsB1.addChooseFromList(oForm, false, "1", "AccountCodeMTRCFL"); //G/L Accounts
                    oColumn = oColumns.Add("AcctCode", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("GLAccountCode");
                    oColumn.DataBind.Bind(UID, "AcctCode");
                    oColumn.ChooseFromListUID = "AccountCodeMTRCFL";
                    oColumn.ChooseFromListAlias = "AcctCode";
                    oLink = oColumn.ExtendedObject;
                    oLink.LinkedObjectType = "1"; //G/L Accounts
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
                    oCreationPackage.UniqueID = "BDOSFuelWOForm";
                    oCreationPackage.String = BDOSResources.getTranslate("FuelWriteOffWizard");
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
                    int answer = Program.uiApp.MessageBox(BDOSResources.getTranslate("CloseFuelWriteOffWizard") + "?", 1, BDOSResources.getTranslate("Yes"), BDOSResources.getTranslate("No"), "");

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
                        int answer = Program.uiApp.MessageBox(BDOSResources.getTranslate("CreateDocumentGoodsIssue") + "?", 1, BDOSResources.getTranslate("Yes"), BDOSResources.getTranslate("No"), "");

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
                    if (column.UniqueID == "LineNum" || column.UniqueID == "CheckBox" || column.UniqueID == "FuUomEntry" || column.UniqueID == "LineId")
                        continue;
                    column.Width = mtrWidth;
                }
            }
            catch (Exception ex)
            {
                throw ex;
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
                oForm.Freeze(false);

            }
            catch (Exception ex)
            {
                throw ex;
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
                    if (oCFLEvento.ChooseFromListUID == "ProjectCodeCFL")
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
                    else if (oCFLEvento.ChooseFromListUID == "AccountCodeCFL" || oCFLEvento.ChooseFromListUID == "AccountCodeMTRCFL")
                    {
                        SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID);
                        SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();

                        SAPbouiCOM.Condition oCon = oCons.Add();
                        oCon.Alias = "Postable"; //Active Account, (Title Account)
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = "Y";

                        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                        oCon = oCons.Add();
                        oCon.Alias = "FrozenFor"; //Inactive
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = "N";

                        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                        oCon = oCons.Add();
                        oCon.Alias = "LocManTran"; //Lock Manual Transaction (Control Account)
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = "N";

                        oCFL.SetConditions(oCons);
                    }
                }
                else
                {
                    SAPbouiCOM.DataTable oDataTable = oCFLEvento.SelectedObjects;

                    if (oDataTable != null)
                    {
                        if (oCFLEvento.ChooseFromListUID == "ProjectCodeCFL")
                        {
                            string prjCode = oDataTable.GetValue("PrjCode", 0);
                            LanguageUtils.IgnoreErrors<string>(() => oForm.Items.Item("PrjCodeE").Specific.Value = prjCode);
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
                        else if (oCFLEvento.ChooseFromListUID == "AccountCodeCFL")
                        {
                            string account = oDataTable.GetValue("AcctCode", 0);
                            LanguageUtils.IgnoreErrors<string>(() => oForm.Items.Item("AccountE").Specific.Value = account);
                        }
                        else if (oCFLEvento.ChooseFromListUID == "AccountCodeMTRCFL")
                        {
                            string account = oDataTable.GetValue("AcctCode", 0);
                            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("AssetMTR").Specific;
                            LanguageUtils.IgnoreErrors<string>(() => oMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value = account);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
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

            string docDate = oForm.DataSources.UserDataSources.Item("DocDateE").ValueEx;
            string dateFrom = oForm.DataSources.UserDataSources.Item("DateFromE").ValueEx;
            string dateTo = oForm.DataSources.UserDataSources.Item("DateToE").ValueEx;
            string prjCode = oForm.DataSources.UserDataSources.Item("PrjCodeE").ValueEx;
            string fuTpCode = oForm.DataSources.UserDataSources.Item("FuTpCodeE").ValueEx;
            string itemCode = oForm.DataSources.UserDataSources.Item("ItemCodeE").ValueEx;
            string fuelCode = oForm.DataSources.UserDataSources.Item("FuelCodeE").ValueEx;
            string account = oForm.DataSources.UserDataSources.Item("AccountE").ValueEx;
            string fuGroup = oForm.DataSources.UserDataSources.Item("FuGroupCB").ValueEx;

            if (string.IsNullOrEmpty(dateFrom) || string.IsNullOrEmpty(dateTo))
            {
                string errorText = BDOSResources.getTranslate("TheFollowingFieldsAreMandatory") + " : \"" + oForm.Items.Item("PeriodS").Specific.caption + "\"";
                Program.uiApp.SetStatusBarMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                return;
            }

            StringBuilder query = new StringBuilder();
            query.Append("SELECT \"@BDOSFUCN\".\"DocEntry\" AS \"Entry\", \n");
            query.Append("\"@BDOSFUCN\".\"U_DocDate\", \n");
            query.Append("\"@BDOSFUCN\".\"U_DateFrom\", \n");
            query.Append("\"@BDOSFUCN\".\"U_DateTo\", \n");
            query.Append("\"@BDOSFUCN\".\"U_PrjCode\", \n");
            query.Append("\"@BDOSFUCN\".\"U_FuNrCode\", \n");
            query.Append("\"@BDOSFUC1\".*, \n");
            query.Append("\"OITM\".\"ItemName\" AS \"FuelName\" \n");
            query.Append("FROM \"@BDOSFUCN\" \n");
            query.Append("INNER JOIN \"@BDOSFUC1\" \n");
            query.Append("ON \"@BDOSFUCN\".\"DocEntry\" = \"@BDOSFUC1\".\"DocEntry\" \n");
            query.Append("INNER JOIN \"OITM\" \n");
            query.Append("ON \"@BDOSFUC1\".\"U_FuelCode\" = \"OITM\".\"ItemCode\" \n");
            query.Append("WHERE  \"@BDOSFUCN\".\"Canceled\" = 'N' \n");
            query.Append("AND \"@BDOSFUCN\".\"U_DocDate\" >= '" + dateFrom + "' \n");
            query.Append("AND \"@BDOSFUCN\".\"U_DocDate\" <= '" + dateTo + "' \n");

            if (!string.IsNullOrEmpty(prjCode))
            {
                query.Append("AND \"@BDOSFUCN\".\"U_PrjCode\" = '" + prjCode + "' \n");
            }
            if (!string.IsNullOrEmpty(itemCode))
            {
                query.Append("AND \"@BDOSFUC1\".\"U_ItemCode\" = '" + itemCode + "' \n");
            }
            if (!string.IsNullOrEmpty(fuTpCode))
            {
                query.Append("AND \"@BDOSFUC1\".\"U_FuTpCode\" = '" + fuTpCode + "' \n");
            }
            if (!string.IsNullOrEmpty(fuelCode))
            {
                query.Append("AND \"@BDOSFUC1\".\"U_FuelCode\" = '" + fuelCode + "' \n");
            }
            if (!string.IsNullOrEmpty(fuGroup))
            {
                query.Append("AND \"OITM\".\"ItmsGrpCod\" = '" + fuGroup + "' \n");
            }

            query.Append("ORDER BY \"@BDOSFUCN\".\"DocEntry\"");

            oRecordSet.DoQuery(query.ToString());

            try
            {
                int rowIndex = 0;

                while (!oRecordSet.EoF)
                {
                    oDataTable.Rows.Add();
                    oDataTable.SetValue("LineNum", rowIndex, rowIndex + 1);
                    oDataTable.SetValue("CheckBox", rowIndex, "N");
                    oDataTable.SetValue("DocEntry", rowIndex, oRecordSet.Fields.Item("DocEntry").Value);
                    if (oRecordSet.Fields.Item("U_DocEntryGI").Value != 0)
                        oDataTable.SetValue("DocEntryGI", rowIndex, oRecordSet.Fields.Item("U_DocEntryGI").Value);
                    oDataTable.SetValue("DocDate", rowIndex, oRecordSet.Fields.Item("U_DocDate").Value.ToString("yyyyMMdd") == "18991230" ? "" : oRecordSet.Fields.Item("U_DocDate").Value.ToString("yyyyMMdd"));
                    oDataTable.SetValue("DateFrom", rowIndex, oRecordSet.Fields.Item("U_DateFrom").Value.ToString("yyyyMMdd") == "18991230" ? "" : oRecordSet.Fields.Item("U_DateFrom").Value.ToString("yyyyMMdd"));
                    oDataTable.SetValue("DateTo", rowIndex, oRecordSet.Fields.Item("U_DateTo").Value.ToString("yyyyMMdd") == "18991230" ? "" : oRecordSet.Fields.Item("U_DateTo").Value.ToString("yyyyMMdd"));
                    oDataTable.SetValue("PrjCode", rowIndex, oRecordSet.Fields.Item("U_PrjCode").Value);
                    oDataTable.SetValue("FuNrCode", rowIndex, oRecordSet.Fields.Item("U_FuNrCode").Value);
                    oDataTable.SetValue("LineId", rowIndex, oRecordSet.Fields.Item("LineId").Value);
                    oDataTable.SetValue("ItemCode", rowIndex, oRecordSet.Fields.Item("U_ItemCode").Value);
                    oDataTable.SetValue("ItemName", rowIndex, oRecordSet.Fields.Item("U_ItemName").Value);
                    oDataTable.SetValue("FuTpCode", rowIndex, oRecordSet.Fields.Item("U_FuTpCode").Value);
                    oDataTable.SetValue("FuelCode", rowIndex, oRecordSet.Fields.Item("U_FuelCode").Value);
                    oDataTable.SetValue("FuelName", rowIndex, oRecordSet.Fields.Item("FuelName").Value);
                    oDataTable.SetValue("FuUomEntry", rowIndex, oRecordSet.Fields.Item("U_FuUomEntry").Value);
                    oDataTable.SetValue("FuUomCode", rowIndex, oRecordSet.Fields.Item("U_FuUomCode").Value);
                    oDataTable.SetValue("FuPerKm", rowIndex, oRecordSet.Fields.Item("U_FuPerKm").Value);
                    oDataTable.SetValue("FuPerHr", rowIndex, oRecordSet.Fields.Item("U_FuPerHr").Value);
                    oDataTable.SetValue("OdmtrStart", rowIndex, oRecordSet.Fields.Item("U_OdmtrStart").Value);
                    oDataTable.SetValue("OdmtrEnd", rowIndex, oRecordSet.Fields.Item("U_OdmtrEnd").Value);
                    oDataTable.SetValue("HrsWorked", rowIndex, oRecordSet.Fields.Item("U_HrsWorked").Value);
                    oDataTable.SetValue("NormCn", rowIndex, oRecordSet.Fields.Item("U_NormCn").Value);
                    oDataTable.SetValue("ActuallyCn", rowIndex, oRecordSet.Fields.Item("U_ActuallyCn").Value);

                    for (int i = 1; i <= 5; i++)
                    {
                        oDataTable.SetValue("Dimension" + i, rowIndex, oRecordSet.Fields.Item("U_Dimension" + i).Value);
                    }
                    oDataTable.SetValue("AcctCode", rowIndex, account);

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
                throw ex;
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
            string whsCode = oForm.DataSources.UserDataSources.Item("WhsCodeE").ValueEx;

            if (string.IsNullOrEmpty(docDateStr) || string.IsNullOrEmpty(whsCode))
            {
                errorText = BDOSResources.getTranslate("TheFollowingFieldsAreMandatory") + " : \"" + oForm.Items.Item("DocDateS").Specific.caption + "\", \"" + oForm.Items.Item("WhsCodeS").Specific.caption + "\"";
                Program.uiApp.SetStatusBarMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                return;
            }

            DateTime docDate = Convert.ToDateTime(DateTime.ParseExact(docDateStr, "yyyyMMdd", CultureInfo.InvariantCulture));

            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("AssetMTR").Specific;
            oMatrix.FlushToDataSource();

            SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("AssetMTR");
            string checkBox;
            string fuTpCode;

            SAPbobsCOM.Documents oGoodsIssue = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit);
            oGoodsIssue.DocDate = docDate;

            for (int i = 0; i < oDataTable.Rows.Count; i++)
            {
                checkBox = oDataTable.GetValue("CheckBox", i);
                fuTpCode = oDataTable.GetValue("FuTpCode", i);

                if (checkBox == "Y" && !string.IsNullOrEmpty(fuTpCode) && oDataTable.GetValue("DocEntryGI", i) == 0 && oDataTable.GetValue("ActuallyCn", i) > 0)
                {
                    oGoodsIssue.Lines.ItemCode = oDataTable.GetValue("FuelCode", i);
                    oGoodsIssue.Lines.ItemDescription = oDataTable.GetValue("FuelName", i);
                    oGoodsIssue.Lines.WarehouseCode = whsCode;
                    double actuallyCn = Convert.ToDouble(oDataTable.GetValue("ActuallyCn", i), CultureInfo.InvariantCulture);
                    oGoodsIssue.Lines.Quantity = actuallyCn;
                    oGoodsIssue.Lines.AccountCode = oDataTable.GetValue("AcctCode", i);
                    oGoodsIssue.Lines.ProjectCode = oDataTable.GetValue("PrjCode", i);
                    oGoodsIssue.Lines.CostingCode = oDataTable.GetValue("Dimension1", i);
                    oGoodsIssue.Lines.CostingCode2 = oDataTable.GetValue("Dimension2", i);
                    oGoodsIssue.Lines.CostingCode3 = oDataTable.GetValue("Dimension3", i);
                    oGoodsIssue.Lines.CostingCode4 = oDataTable.GetValue("Dimension4", i);
                    oGoodsIssue.Lines.CostingCode5 = oDataTable.GetValue("Dimension5", i);

                    oGoodsIssue.Lines.Add();
                }
            }

            int resultCode = oGoodsIssue.Add();

            if (resultCode != 0)
            {
                int errCode;
                string errMsg;
                Program.oCompany.GetLastError(out errCode, out errMsg);
                Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("DocumentNotCreated") + ". " + BDOSResources.getTranslate("ReasonIs") + ": " + errMsg, SAPbouiCOM.BoMessageTime.bmt_Short);
            }
            else
            {
                int docEntryGI;
                bool newDoc = oGoodsIssue.GetByKey(Convert.ToInt32(Program.oCompany.GetNewObjectKey()));
                if (newDoc == true)
                {
                    docEntryGI = oGoodsIssue.DocEntry;
                    Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("DocumentCreatedSuccesfully") + ": " + docEntryGI, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                }
                else
                {
                    Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("DocumentNotCreated"), SAPbouiCOM.BoMessageTime.bmt_Short);
                    return;
                }

                Marshal.ReleaseComObject(oGoodsIssue);

                for (int i = 0; i < oDataTable.Rows.Count; i++)
                {
                    checkBox = oDataTable.GetValue("CheckBox", i);
                    if (docEntryGI > 0 && checkBox == "Y" && oDataTable.GetValue("DocEntryGI", i) == 0)
                    {
                        oDataTable.SetValue("DocEntryGI", i, docEntryGI);
                    }
                }

                for (int i = 0; i < oDataTable.Rows.Count; i++)
                {
                    checkBox = oDataTable.GetValue("CheckBox", i);
                    if (docEntryGI > 0 && checkBox == "Y")
                    {
                        int docEntry = oDataTable.GetValue("DocEntry", i);
                        int lineId = oDataTable.GetValue("LineId", i);
                        updateDocumentFuelConsumptionAct(docEntry, lineId, docEntryGI);
                    }
                }
            }

            oForm.Freeze(true);
            oMatrix.Clear();
            oMatrix.LoadFromDataSource();
            oForm.Update();
            oForm.Freeze(false);
        }

        public static void updateDocumentFuelConsumptionAct(int docEntry, int lineId, int docEntryGI)
        {
            SAPbobsCOM.CompanyService oCompanyService = null;
            SAPbobsCOM.GeneralService oGeneralService = null;
            SAPbobsCOM.GeneralData oGeneralData = null;
            SAPbobsCOM.GeneralDataParams oGeneralParams = null;

            try
            {
                oCompanyService = Program.oCompany.GetCompanyService();
                oGeneralService = oCompanyService.GetGeneralService("UDO_F_BDOSFUCN_D");
                oGeneralData = ((SAPbobsCOM.GeneralData)(oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)));

                oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                oGeneralParams.SetProperty("DocEntry", docEntry);
                oGeneralData = oGeneralService.GetByParams(oGeneralParams);

                int oChildCount = oGeneralData.Child("BDOSFUC1").Count;
                if (oChildCount > 0)
                {
                    oGeneralData.Child("BDOSFUC1").Item(lineId - 1).SetProperty("U_DocEntryGI", docEntryGI);
                }

                oGeneralService.Update(oGeneralData);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                Marshal.ReleaseComObject(oGeneralParams);
                Marshal.ReleaseComObject(oGeneralData);
                Marshal.ReleaseComObject(oGeneralService);
                Marshal.ReleaseComObject(oCompanyService);
            }
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
                throw ex;
            }
            finally
            {
                Marshal.ReleaseComObject(oRecordSet);
            }
        }
    }
}
