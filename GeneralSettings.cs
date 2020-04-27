using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;
using System.Reflection;
using System.Resources;
using System.Runtime.InteropServices;
using System.Threading;
using SAPbobsCOM;
using static BDO_Localisation_AddOn.Program;

namespace BDO_Localisation_AddOn
{
    static partial class GeneralSettings
    {
        public static void createUserFields(out string errorText)
        {
            errorText = null;
            Dictionary<string, object> fieldskeysMap;

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSAtPayA");
            fieldskeysMap.Add("TableName", "OADM");
            fieldskeysMap.Add("Description", "Account for Down Automatic Payments in Internet bank");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 15);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSRcDPPr");
            fieldskeysMap.Add("TableName", "OADM");
            fieldskeysMap.Add("Description", "Received Down Payment Purpose");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 100);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSBTCNPR");
            fieldskeysMap.Add("TableName", "OADM");
            fieldskeysMap.Add("Description", "Batch Number Prefix");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 2);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSFADim");
            fieldskeysMap.Add("TableName", "OADM");
            fieldskeysMap.Add("Description", "Fixed Asset Dimension");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 100);

            UDO.addUserTableFields(fieldskeysMap, out errorText);
            
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "StockRevWh");
            fieldskeysMap.Add("TableName", "OADM");
            fieldskeysMap.Add("Description", "Stock Revaluation Whs");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 8);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSStock");
            fieldskeysMap.Add("TableName", "OADM");
            fieldskeysMap.Add("Description", "Stock Revaluation");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 1);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            GC.Collect();
        }

        public static void createFormItems(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            Dictionary<string, object> formItems;
            string itemName = "";

            int pane = 21;
            int height = 14;
            SAPbouiCOM.Item oItem;

            //Construction-ც ემატება
            try
            {
                oItem = oForm.Items.Item("BDOSLIC");
            }
            catch
            {
                //ლიცენზირება (ჩანართი)
                SAPbouiCOM.Item oFolder = oForm.Items.Item("46");
                formItems = new Dictionary<string, object>();
                itemName = "BDOSLIC";
                formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_FOLDER);
                formItems.Add("Left", oFolder.Left + oFolder.Width);
                formItems.Add("Width", oFolder.Width);
                formItems.Add("Top", oFolder.Top);
                formItems.Add("Height", oFolder.Height);
                formItems.Add("UID", itemName);
                formItems.Add("Caption", BDOSResources.getTranslate("License"));
                formItems.Add("Pane", pane);
                formItems.Add("ValOn", "0");
                formItems.Add("ValOff", itemName);
                formItems.Add("GroupWith", "46");
                formItems.Add("AffectsFormMode", false);
                formItems.Add("Description", BDOSResources.getTranslate("AddonLicense"));

                FormsB1.createFormItem(oForm, formItems, out errorText);
                if (errorText != null)
                {
                    return;
                }
            }

            int top = 37;

            oItem = oForm.Items.Item("234000016");
            int left = oForm.Items.Item("234000015").Left;
            int width = 125;

            Dictionary<string, string> CompanyLicenseInfo = CommonFunctions.getCompanyLicenseInfo();
            string licenseStatus = CompanyLicenseInfo["LicenseStatus"];
            string licenseUpdateDate = CompanyLicenseInfo["LicenseUpdateDate"];
            string licenseQuantity = CompanyLicenseInfo["LicenseQuantity"];

            formItems = new Dictionary<string, object>();
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            formItems.Add("Left", left);
            formItems.Add("Width", 2 * width);
            formItems.Add("Top", top);
            formItems.Add("Height", oItem.Height);
            formItems.Add("Caption", BDOSResources.getTranslate("LicenseLocalisationProgram"));
            formItems.Add("UID", "BDOSLLicPr");
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            left = oItem.Left;
            int left_e = oItem.Left + width + 5;
            top = top + 20;

            formItems = new Dictionary<string, object>();
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left);
            formItems.Add("Width", width);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("Caption", BDOSResources.getTranslate("LicenseStatus"));
            formItems.Add("UID", "BDOSLLcStN");
            formItems.Add("LinkTo", "BDOSLLicSt");
            formItems.Add("RightJustified", false);
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", "BDOSLLicSt");
            formItems.Add("DisplayDesc", true);
            formItems.Add("IsPassword", false);
            formItems.Add("Description", BDOSResources.getTranslate("LicenseStatus"));
            formItems.Add("RightJustified", false);
            formItems.Add("Enabled", false);
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);
            formItems.Add("isDataSource", true);
            formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
            formItems.Add("DataSource", "UserDataSources");
            formItems.Add("Length", 150);
            formItems.Add("Value", licenseStatus);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + 20;

            formItems = new Dictionary<string, object>();
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left);
            formItems.Add("Width", width);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("Caption", BDOSResources.getTranslate("LicenseUpdateDate"));
            formItems.Add("UID", "BDOSLLicDN");
            formItems.Add("LinkTo", "BDOSLLicD");
            formItems.Add("RightJustified", false);
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", "BDOSLLicD");
            formItems.Add("DisplayDesc", true);
            formItems.Add("IsPassword", false);
            formItems.Add("Description", BDOSResources.getTranslate("LicenseUpdateDate"));
            formItems.Add("RightJustified", false);
            formItems.Add("Enabled", false);
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);
            formItems.Add("isDataSource", true);
            formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
            formItems.Add("DataSource", "UserDataSources");
            formItems.Add("Length", 150);
            formItems.Add("Value", licenseUpdateDate);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            //ლიცენზირება

            string objectType = "1"; //Account

            //Cash Flow  
            oItem = oForm.Items.Item("247");

            pane = 14;
            left = oItem.Left;
            width = 270;
            top = oItem.Top + 2 * oItem.Height + 2 * 25;

            pane = 14;

            itemName = "ActAutPayS";
            formItems = new Dictionary<string, object>();
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left);
            formItems.Add("Width", width);
            formItems.Add("Top", top);
            formItems.Add("Height", 15);
            formItems.Add("Caption", BDOSResources.getTranslate("AcctForAutoPaymentInternetBank"));
            formItems.Add("UID", itemName);
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);
            formItems.Add("LinkTo", "ActAutPayE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            //left = left + 70 + 10;
            bool multiSelection = false;
            string uniqueID_lf_Project = "GLAccount_CFL";
            FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_lf_Project);

            formItems = new Dictionary<string, object>();
            itemName = "ActAutPayE";
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
            formItems.Add("Length", 30);
            formItems.Add("Size", 20);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("TableName", "OADM");
            formItems.Add("Alias", "U_BDOSAtPayA");
            formItems.Add("Bound", true);
            formItems.Add("Left", left + width + 5 + 10);
            formItems.Add("Width", width);
            formItems.Add("Top", top);
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);
            formItems.Add("Height", 19);
            formItems.Add("UID", itemName);
            formItems.Add("ChooseFromListUID", uniqueID_lf_Project);
            formItems.Add("ChooseFromListAlias", "AcctCode");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            //golden arrow
            formItems = new Dictionary<string, object>();
            itemName = "ActAutPayL"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            formItems.Add("Left", left + width + 5 - 10);
            formItems.Add("Top", top);
            formItems.Add("Height", 14);
            formItems.Add("UID", itemName);
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);
            formItems.Add("LinkTo", "ActAutPayE");
            formItems.Add("LinkedObjectType", objectType);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + 19 + 5;

            //RecDPPurp
            itemName = "RecDPPurpS";
            formItems = new Dictionary<string, object>();
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left);
            formItems.Add("Width", width);
            formItems.Add("Top", top);
            formItems.Add("Height", 15);
            formItems.Add("Caption", BDOSResources.getTranslate("ReceivedDownPaymentPurpose"));
            formItems.Add("UID", itemName);
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);
            formItems.Add("LinkTo", "RecDPPurpE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "RecDPPurpE";
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
            formItems.Add("Length", 30);
            formItems.Add("Size", 100);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("TableName", "OADM");
            formItems.Add("Alias", "U_BDOSRcDPPr");
            formItems.Add("Bound", true);
            formItems.Add("Left", left + width + 5 + 10);
            formItems.Add("Width", width);
            formItems.Add("Top", top);
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);
            formItems.Add("Height", 19);
            formItems.Add("UID", itemName);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }


            pane = 15;

            left = 300;
            top = oForm.Items.Item("540002042").Top;



            itemName = "BDOSFADimS";
            formItems = new Dictionary<string, object>();
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left);
            formItems.Add("Width", width);
            formItems.Add("Top", top);
            formItems.Add("Height", 15);
            formItems.Add("Caption", BDOSResources.getTranslate("FixedAssetsDimension"));
            formItems.Add("UID", itemName);
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);
            formItems.Add("LinkTo", "BDOSFADimE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            left_e = left + width + 5 + 30;

            Dictionary<string, string> activeDimensionsList = CommonFunctions.getActiveDimensionsList(out errorText);

            //objectType = "251";
            //multiSelection = false;
            //uniqueID_lf_Project = "Dim_CFL";
            //FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_lf_Project);

            formItems = new Dictionary<string, object>();
            itemName = "BDOSFADimE";
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
            formItems.Add("Length", 30);
            formItems.Add("Size", 100);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            formItems.Add("TableName", "OADM");
            formItems.Add("Alias", "U_BDOSFADim");
            formItems.Add("Bound", true);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width);
            formItems.Add("Top", top);
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);
            formItems.Add("Height", 19);
            formItems.Add("UID", itemName);
            //formItems.Add("ChooseFromListUID", uniqueID_lf_Project);
            //formItems.Add("ChooseFromListAlias", "DimCode");
            formItems.Add("ValidValues", activeDimensionsList);
            formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_ValueDescription);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            //---------------------------------
            pane = 22;
            try
            {
                oItem = oForm.Items.Item("BDOSBTCHN");
            }
            catch
            {
                //Batch Number (ჩანართი)
                SAPbouiCOM.Item oFolder = oForm.Items.Item("46");
                formItems = new Dictionary<string, object>();
                itemName = "BDOSBTCHN";
                formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_FOLDER);
                formItems.Add("Left", oFolder.Left + 2 * oFolder.Width);
                formItems.Add("Width", oFolder.Width);
                formItems.Add("Top", oFolder.Top);
                formItems.Add("Height", oFolder.Height);
                formItems.Add("UID", itemName);
                formItems.Add("Caption", BDOSResources.getTranslate("BatchNumber"));
                formItems.Add("Pane", pane);
                formItems.Add("ValOn", "0");
                formItems.Add("ValOff", itemName);
                formItems.Add("GroupWith", "46");
                formItems.Add("AffectsFormMode", false);
                formItems.Add("Description", BDOSResources.getTranslate("BatchNumber"));

                FormsB1.createFormItem(oForm, formItems, out errorText);
                if (errorText != null)
                {
                    return;
                }
            }


            //-----
            left = oForm.Items.Item("234000016").Left;
            width = oForm.Items.Item("BDOSLLicSt").Width;

            left_e = oForm.Items.Item("BDOSLLicSt").Left;

            formItems = new Dictionary<string, object>();
            itemName = "BDOSBTCHPS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left);
            formItems.Add("Width", width);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("BatchNumberPrefix"));
            formItems.Add("LinkTo", "cardCodeE");
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }


            formItems = new Dictionary<string, object>();
            itemName = "BDOSBTCHPE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OADM");
            formItems.Add("Alias", "U_BDOSBTCNPR");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            pane = 8;

            left = oForm.Items.Item("20").Left;
            top = oForm.Items.Item("20").Top + 15;

            formItems = new Dictionary<string, object>();
            itemName = "BDOSStock";
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OADM");
            formItems.Add("Alias", "U_BDOSStock");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
            formItems.Add("Length", 1);
            formItems.Add("Left", left);
            formItems.Add("Width", 160);
            formItems.Add("Top", top);
            formItems.Add("Height", 14);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("StockOnOff"));
            formItems.Add("ValOff", "N");
            formItems.Add("ValOn", "Y");
            formItems.Add("DisplayDesc", true);
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = oForm.Items.Item("BDOSStock").Top + 15;

            formItems = new Dictionary<string, object>();
            itemName = "StRevWhST";
            formItems.Add("Size", 20);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left);
            formItems.Add("Width", 120);
            formItems.Add("Top", top);
            formItems.Add("Caption", BDOSResources.getTranslate("Stock Revaluation Whs"));
            formItems.Add("UID", itemName);
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            multiSelection = false;
            objectType = "64";
            string uniqueID_Whs = "WhsFr_CFL";
            FormsB1.addChooseFromList(oForm, true, objectType, uniqueID_Whs);

            left = oForm.Items.Item("172").Left;

            formItems = new Dictionary<string, object>();
            itemName = "StockRevWh";
            
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OADM");
            formItems.Add("Alias", "U_StockRevWh");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            //formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
            formItems.Add("Left", left);
            formItems.Add("Width", 100);
            formItems.Add("Top", top);
            formItems.Add("Height", 19);
            formItems.Add("UID", itemName);
            formItems.Add("ChooseFromListUID", uniqueID_Whs);
            formItems.Add("ChooseFromListAlias", "WhsCode");
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);
            
            //formItems.Add("Length", 30);
            //formItems.Add("Size", 20);
            
            //

            //formItems.Add("Editable", true);
            //formItems.Add("Enabled", true);
            

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }
            //oForm.Items.Item(itemName).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True);
            //oForm.Items.Item(itemName).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True);

            formItems = new Dictionary<string, object>();
            itemName = "BDOSSRWLB"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            formItems.Add("Left", left - 20);
            formItems.Add("Top", top);
            formItems.Add("Height", 19);
            formItems.Add("UID", itemName);
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);
            formItems.Add("LinkTo", "StockRevWh");
            formItems.Add("LinkedObjectType", objectType);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }
            


            ////Batch Number ცხრილი
            //itemName = "BTCHMatrix";
            //formItems = new Dictionary<string, object>();
            //formItems.Add("isDataSource", true);
            //formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
            //formItems.Add("Left", left);
            //formItems.Add("Width", oForm.Width - 50);
            //formItems.Add("Top", top);
            //formItems.Add("Height", 100);
            //formItems.Add("UID", itemName);
            //formItems.Add("FromPane", pane);
            //formItems.Add("ToPane", pane);

            //FormsB1.createFormItem(oForm, formItems, out errorText);
            //if (errorText != null)
            //{
            //    return;
            //}

            ////-------------------


            ////----- Batch Number Table
            //SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Add("BTCHTable");
            //oDataTable.Columns.Add("#", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20); //0
            //oDataTable.Columns.Add("Type", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100); //1
            //oDataTable.Columns.Add("Value", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);//2
            ////-----



            ////------
            //SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("BTCHMatrix").Specific));
            //SAPbouiCOM.Columns oColumns = oMatrix.Columns;


            //SAPbouiCOM.Column oColumn = oColumns.Add("#", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            //oColumn.TitleObject.Caption = "#";
            //oColumn.Width = 40;
            //oColumn.Editable = false;
            //oColumn.DataBind.Bind("BTCHTable", "#");

            //oColumn = oColumns.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            //oColumn.TitleObject.Caption = BDOSResources.getTranslate("Type");
            //oColumn.Width = 40;
            //oColumn.Editable = true;
            //oColumn.DataBind.Bind("BTCHTable", "Type");
            //oColumn.DisplayDesc = true;
            //oColumn.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;

            //oColumn.ValidValues.Add("Custom", "Custom");
            //oColumn.ValidValues.Add("Year", "Year");
            //oColumn.ValidValues.Add("Month", "Month");
            //oColumn.ValidValues.Add("Day", "Day");



            //oColumn = oColumns.Add("Value", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            //oColumn.TitleObject.Caption = "Value";
            //oColumn.Width = 40;
            //oColumn.Editable = true;
            //oColumn.DataBind.Bind("BTCHTable", "Value");
            ////--------


        }

        public static void setValueCFLEvent(SAPbouiCOM.Form oForm, SAPbouiCOM.ChooseFromListEvent oCFL, out string errorText)
        {
            errorText = null;

            if (oCFL.ChooseFromListUID == "Dimension_CFL")
            {
                SAPbouiCOM.DataTable oDataTableSelectedObjects = oCFL.SelectedObjects;
                string OcrCode = oDataTableSelectedObjects.GetValue("DimCode", 0);

                try
                {
                    SAPbouiCOM.EditText oEdit = oForm.Items.Item("BDOSFADimE").Specific;
                    oEdit.Value = OcrCode;
                }
                catch { }
                finally
                {
                    if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                    {
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                    }
                    GC.Collect();
                }
            }

            if (oCFL.ChooseFromListUID == "GLAccount_CFL")
            {
                try
                {
                    SAPbouiCOM.DataTable oDataTableSelectedObjects = oCFL.SelectedObjects;
                    string AcctCode = oDataTableSelectedObjects.GetValue("AcctCode", 0);

                    try
                    {
                        SAPbouiCOM.EditText oEdit = oForm.Items.Item("ActAutPayE").Specific;
                        oEdit.Value = AcctCode;
                    }
                    catch { }

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
                    if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                    {
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                    }
                    GC.Collect();
                }
            }
            if (oCFL.ChooseFromListUID == "WhsFr_CFL")
            {
                SAPbouiCOM.DataTable oDataTableSelectedObjects = oCFL.SelectedObjects;
                string WhsCode = oDataTableSelectedObjects.GetValue("WhsCode", 0);
                LanguageUtils.IgnoreErrors<string>(() => oForm.Items.Item("StockRevWh").Specific.Value = WhsCode);
                SAPbouiCOM.EditText oEdit = oForm.Items.Item("StockRevWh").Specific;
                oEdit.Value = WhsCode;
            }
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
                    oForm.PaneLevel = 3;
                    oForm.Items.Item("45").Click();
                }

                if (pVal.ItemUID == "BDOSLIC" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED & pVal.BeforeAction == true)
                {
                    oForm.PaneLevel = 21;
                }

                if (pVal.ItemUID == "BDOSBTCHN" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED & pVal.BeforeAction == true)
                {
                    oForm.PaneLevel = 22;
                }

                if (pVal.ItemUID == "BDOSLLicPr" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK & pVal.BeforeAction == false)
                {
                    License.createAddOnLicenseForm(out errorText);
                }

                if ((pVal.ItemUID == "ActAutPayE") & pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                {
                    if (pVal.BeforeAction == false)
                    {
                        SAPbouiCOM.ChooseFromListEvent oCFLEvento = null;
                        oCFLEvento = ((SAPbouiCOM.ChooseFromListEvent)(pVal));

                        setValueCFLEvent(oForm, oCFLEvento, out errorText);
                    }
                    else
                    {
                        string sCFL_ID = "GLAccount_CFL";
                        SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(sCFL_ID);

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

                if ((pVal.ItemUID == "StockRevWh") & pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                {
                    if (pVal.BeforeAction == false)
                    {
                        SAPbouiCOM.ChooseFromListEvent oCFLEvento = null;
                        oCFLEvento = ((SAPbouiCOM.ChooseFromListEvent)(pVal));

                        setValueCFLEvent(oForm, oCFLEvento, out errorText);
                    }
                }
            

                    //if (pVal.ItemUID == "1" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK & pVal.BeforeAction == false)
                    //{
                    //var w = Program.uiApp.Forms.ActiveForm.TypeEx;
                    //    string StockWhs = "";
                    //    StockWhs = oForm.Items.Item("StockRevWh").Specific.Value;

                    //}
                }
        }
    }
}