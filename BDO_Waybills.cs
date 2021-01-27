using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Resources;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using SAPbobsCOM;
using static BDO_Localisation_AddOn.Program;

namespace BDO_Localisation_AddOn
{
    static partial class BDO_Waybills
    {
        public static void createDocumentUDO(out string errorText)
        {
            errorText = null;
            string tableName = "BDO_WBLD";
            string description = "Waybill";

            int result = UDO.addUserTable(tableName, description, BoUTBTableType.bott_Document, out errorText);

            if (result != 0)
            {
                return;
            }

            Dictionary<string, object> fieldskeysMap;
            List<string> listValidValues;

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "wblID");
            fieldskeysMap.Add("TableName", "BDO_WBLD");
            fieldskeysMap.Add("Description", "Waybill ID");
            fieldskeysMap.Add("Type", BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 20);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "number");
            fieldskeysMap.Add("TableName", "BDO_WBLD");
            fieldskeysMap.Add("Description", "Waybill Number");
            fieldskeysMap.Add("Type", BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 20);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            listValidValues = new List<string>();
            listValidValues.Add("With Transport"); //0 //ტრანსპორტირებით
            listValidValues.Add("Without Transport"); //1 /ტრანსპორტირების გარეშე

            fieldskeysMap.Add("Name", "type");
            fieldskeysMap.Add("TableName", "BDO_WBLD");
            fieldskeysMap.Add("Description", "Waybill Type");
            fieldskeysMap.Add("Type", BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);
            fieldskeysMap.Add("ValidValues", listValidValues);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            listValidValues = new List<string>();
            listValidValues.Add(""); //-1
            listValidValues.Add("Saved"); //1
            listValidValues.Add("Active"); //2
            listValidValues.Add("Finished"); //3
            listValidValues.Add("Deleted"); //4
            listValidValues.Add("Canceled"); //5
            listValidValues.Add("Sent To Transporter"); //6

            fieldskeysMap.Add("Name", "status");
            fieldskeysMap.Add("TableName", "BDO_WBLD");
            fieldskeysMap.Add("Description", "Waybill Status");
            fieldskeysMap.Add("Type", BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);
            fieldskeysMap.Add("ValidValues", listValidValues);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //გააქტიურების თარიღი
            fieldskeysMap.Add("Name", "actDate");
            fieldskeysMap.Add("TableName", "BDO_WBLD");
            fieldskeysMap.Add("Description", "Activate Date");
            fieldskeysMap.Add("Type", BoFieldTypes.db_Date);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //ტრანსპორტირების დაწყების თარიღი
            fieldskeysMap.Add("Name", "begDate");
            fieldskeysMap.Add("TableName", "BDO_WBLD");
            fieldskeysMap.Add("Description", "Begin Date");
            fieldskeysMap.Add("Type", BoFieldTypes.db_Date);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //ტრანსპორტირების დაწყების თარიღი (საათები/წუთები)
            fieldskeysMap.Add("Name", "beginTime");
            fieldskeysMap.Add("TableName", "BDO_WBLD");
            fieldskeysMap.Add("Description", "BeginTime");
            fieldskeysMap.Add("Type", BoFieldTypes.db_Date);
            fieldskeysMap.Add("SubType", BoFldSubTypes.st_Time);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "strAddrs");
            fieldskeysMap.Add("TableName", "BDO_WBLD");
            fieldskeysMap.Add("Description", "Start Address");
            fieldskeysMap.Add("Type", BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 100);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "endAddrs");
            fieldskeysMap.Add("TableName", "BDO_WBLD");
            fieldskeysMap.Add("Description", "End Address");
            fieldskeysMap.Add("Type", BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 100);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //გამშვები  //ჩამბარებელი თანამშრომელი
            fieldskeysMap.Add("Name", "recpInfo");
            fieldskeysMap.Add("TableName", "BDO_WBLD");
            fieldskeysMap.Add("Description", "Reception Info");
            fieldskeysMap.Add("Type", BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //გამშვები  //ჩამბარებელი თანამშრომელი
            fieldskeysMap.Add("Name", "recpInfN");
            fieldskeysMap.Add("TableName", "BDO_WBLD");
            fieldskeysMap.Add("Description", "Reception Info Name");
            fieldskeysMap.Add("Type", BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 100);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //მიმღები //საკონტაქტო პირი
            fieldskeysMap.Add("Name", "recvInfo");
            fieldskeysMap.Add("TableName", "BDO_WBLD");
            fieldskeysMap.Add("Description", "Receiver Info");
            fieldskeysMap.Add("Type", BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //მიმღები //საკონტაქტო პირი
            fieldskeysMap.Add("Name", "recvInfN");
            fieldskeysMap.Add("TableName", "BDO_WBLD");
            fieldskeysMap.Add("Description", "Receiver Info Name");
            fieldskeysMap.Add("Type", BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "comment");
            fieldskeysMap.Add("TableName", "BDO_WBLD");
            fieldskeysMap.Add("Description", "Comment");
            fieldskeysMap.Add("Type", BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 254);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //მიწოდების თარიღი
            fieldskeysMap.Add("Name", "delvDate");
            fieldskeysMap.Add("TableName", "BDO_WBLD");
            fieldskeysMap.Add("Description", "Delivery Date");
            fieldskeysMap.Add("Type", BoFieldTypes.db_Date);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            listValidValues = new List<string>();
            listValidValues.Add(""); //-1
            listValidValues.Add("Auto"); //1 //საავტომობილო
            listValidValues.Add("Railway"); //2 //სარკინიგზო
            listValidValues.Add("Aviation"); //3 //საავიაციო
            listValidValues.Add("Other"); //4 //სხვა
            listValidValues.Add("Auto Other Country"); //5 //საავტომობილო - უცხო ქვეყნის
            listValidValues.Add("Auto Transporter"); //6 //გადამზიდავი - საავტომობილო

            fieldskeysMap.Add("Name", "trnsType");
            fieldskeysMap.Add("TableName", "BDO_WBLD");
            fieldskeysMap.Add("Description", "Transport Type");
            fieldskeysMap.Add("Type", BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);
            fieldskeysMap.Add("ValidValues", listValidValues);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "vehicle");
            fieldskeysMap.Add("TableName", "BDO_WBLD");
            fieldskeysMap.Add("Description", "Vehicle");
            fieldskeysMap.Add("Type", BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("LinkedTable", "BDO_VECL");

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "vehicNum");
            fieldskeysMap.Add("TableName", "BDO_WBLD");
            fieldskeysMap.Add("Description", "Vehicle Number");
            fieldskeysMap.Add("Type", BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 20);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "trailNum");
            fieldskeysMap.Add("TableName", "BDO_WBLD");
            fieldskeysMap.Add("Description", "Trailer Number");
            fieldskeysMap.Add("Type", BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 20);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "drvCode");
            fieldskeysMap.Add("TableName", "BDO_WBLD");
            fieldskeysMap.Add("Description", "Driver Code");
            fieldskeysMap.Add("Type", BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "drivTin");
            fieldskeysMap.Add("TableName", "BDO_WBLD");
            fieldskeysMap.Add("Description", "Driver TIN");
            fieldskeysMap.Add("Type", BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 11);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "notRsdnt");
            fieldskeysMap.Add("TableName", "BDO_WBLD");
            fieldskeysMap.Add("Description", "Driver Not Resident");
            fieldskeysMap.Add("Type", BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 1);
            fieldskeysMap.Add("DefaultValue", "N");

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "tporter");
            fieldskeysMap.Add("TableName", "BDO_WBLD");
            fieldskeysMap.Add("Description", "Transporter");
            fieldskeysMap.Add("Type", BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 15);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "tporterN");
            fieldskeysMap.Add("TableName", "BDO_WBLD");
            fieldskeysMap.Add("Description", "Transporter Name");
            fieldskeysMap.Add("Type", BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 100);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "tporterT");
            fieldskeysMap.Add("TableName", "BDO_WBLD");
            fieldskeysMap.Add("Description", "Transporter TIN");
            fieldskeysMap.Add("Type", BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 32);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "trnsExpn");
            fieldskeysMap.Add("TableName", "BDO_WBLD");
            fieldskeysMap.Add("Description", "Transportation Expense");
            fieldskeysMap.Add("Type", BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            listValidValues = new List<string>();
            listValidValues.Add(""); //-1
            listValidValues.Add("Buyer"); //1 //მყიდველი
            listValidValues.Add("Seller"); //2 //გამყიდველი

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "payForTr");
            fieldskeysMap.Add("TableName", "BDO_WBLD");
            fieldskeysMap.Add("Description", "Payment For Transportation");
            fieldskeysMap.Add("Type", BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 10);
            fieldskeysMap.Add("ValidValues", listValidValues);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "baseDoc");
            fieldskeysMap.Add("TableName", "BDO_WBLD");
            fieldskeysMap.Add("Description", "Base Document");
            fieldskeysMap.Add("Type", BoFieldTypes.db_Numeric);
            fieldskeysMap.Add("EditSize", 11);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "baseDTxt");
            fieldskeysMap.Add("TableName", "BDO_WBLD");
            fieldskeysMap.Add("Description", "Base Document");
            fieldskeysMap.Add("Type", BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 11);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "baseDocT");
            fieldskeysMap.Add("TableName", "BDO_WBLD");
            fieldskeysMap.Add("Description", "Base Document Type");
            fieldskeysMap.Add("Type", BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "cardCode");
            fieldskeysMap.Add("TableName", "BDO_WBLD");
            fieldskeysMap.Add("Description", "Customer Code");
            fieldskeysMap.Add("Type", BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 15);

            UDO.addUserTableFields(fieldskeysMap, out errorText);
        }

        public static void registerUDO(out string errorText)
        {
            errorText = null;
            string code = "UDO_F_BDO_WBLD_D"; //20 characters (must include at least one alphabetical character).
            Dictionary<string, object> formProperties;

            formProperties = new Dictionary<string, object>();
            formProperties.Add("Name", "Waybill"); //100 characters
            formProperties.Add("TableName", "BDO_WBLD");
            formProperties.Add("ObjectType", BoUDOObjType.boud_Document);
            formProperties.Add("CanCancel", BoYesNoEnum.tYES);
            formProperties.Add("CanClose", BoYesNoEnum.tYES);
            formProperties.Add("CanCreateDefaultForm", BoYesNoEnum.tYES);
            formProperties.Add("CanDelete", BoYesNoEnum.tNO);
            formProperties.Add("CanFind", BoYesNoEnum.tYES);
            formProperties.Add("CanYearTransfer", BoYesNoEnum.tYES);
            formProperties.Add("ManageSeries", BoYesNoEnum.tYES);

            List<Dictionary<string, object>> listFindColumns = new List<Dictionary<string, object>>();
            List<Dictionary<string, object>> listFormColumns = new List<Dictionary<string, object>>();
            Dictionary<string, object> fieldskeysMap;

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "U_cardCode");
            fieldskeysMap.Add("ColumnDescription", "Customer Code"); //30 characters
            listFindColumns.Add(fieldskeysMap);
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "U_wblID");
            fieldskeysMap.Add("ColumnDescription", "Waybill ID"); //30 characters
            listFindColumns.Add(fieldskeysMap);
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "U_number");
            fieldskeysMap.Add("ColumnDescription", "Waybill Number"); //30 characters
            listFindColumns.Add(fieldskeysMap);
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "U_status");
            fieldskeysMap.Add("ColumnDescription", "Waybill Status"); //30 characters
            listFindColumns.Add(fieldskeysMap);
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "U_type");
            fieldskeysMap.Add("ColumnDescription", "Waybill Type"); //30 characters
            listFindColumns.Add(fieldskeysMap);
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "U_actDate");
            fieldskeysMap.Add("ColumnDescription", "Activate Date"); //30 characters
            listFindColumns.Add(fieldskeysMap);
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "U_delvDate");
            fieldskeysMap.Add("ColumnDescription", "Delivery Date"); //30 characters
            listFindColumns.Add(fieldskeysMap);
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "U_trnsType");
            fieldskeysMap.Add("ColumnDescription", "Transport Type"); //30 characters
            listFindColumns.Add(fieldskeysMap);
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "U_drivTin");
            fieldskeysMap.Add("ColumnDescription", "Driver TIN"); //30 characters
            listFindColumns.Add(fieldskeysMap);
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "U_tporter");
            fieldskeysMap.Add("ColumnDescription", "Transporter"); //30 characters
            listFindColumns.Add(fieldskeysMap);
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "U_tporterN");
            fieldskeysMap.Add("ColumnDescription", "Transporter Name"); //30 characters
            listFindColumns.Add(fieldskeysMap);
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "U_comment");
            fieldskeysMap.Add("ColumnDescription", "Comment"); //30 characters
            listFindColumns.Add(fieldskeysMap);
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "DocEntry");
            fieldskeysMap.Add("ColumnDescription", "DocEntry"); //30 characters
            listFindColumns.Add(fieldskeysMap);
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "DocNum");
            fieldskeysMap.Add("ColumnDescription", "Number"); //30 characters
            listFindColumns.Add(fieldskeysMap);
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "Status");
            fieldskeysMap.Add("ColumnDescription", "Status"); //30 characters
            listFindColumns.Add(fieldskeysMap);
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "CreateDate");
            fieldskeysMap.Add("ColumnDescription", "Create Date"); //30 characters
            listFindColumns.Add(fieldskeysMap);
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "UpdateDate");
            fieldskeysMap.Add("ColumnDescription", "Update Date"); //30 characters
            listFindColumns.Add(fieldskeysMap);
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "Remark");
            fieldskeysMap.Add("ColumnDescription", "Remark"); //30 characters
            listFindColumns.Add(fieldskeysMap);

            formProperties.Add("FindColumns", listFindColumns);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("FormColumnAlias", "DocEntry");
            fieldskeysMap.Add("FormColumnDescription", "DocEntry"); //30 characters
            listFormColumns.Add(fieldskeysMap);

            formProperties.Add("FormColumns", listFormColumns);

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
                fatherMenuItem = uiApp.Menus.Item("2048");
                // Add a pop-up menu item
                oCreationPackage = (SAPbouiCOM.MenuCreationParams)uiApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                oCreationPackage.Checked = false;
                oCreationPackage.Enabled = true;
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "UDO_F_BDO_WBLD_D";
                oCreationPackage.String = BDOSResources.getTranslate("Waybill");
                oCreationPackage.Position = fatherMenuItem.SubMenus.Count - 1;

                menuItem = fatherMenuItem.SubMenus.AddEx(oCreationPackage);
            }
            catch
            {

            }
        }

        public static void createFormItems(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            Dictionary<string, object> formItems;

            string addonName = "BDOS Localisation AddOn";
            string addonFormType = "UDO_FT_UDO_F_BDO_WBLD_D";
            oForm.ReportType = CrystalReports.getReportTypeCode(addonName, addonFormType, out errorText);

            string itemName = "";

            int left_s = 6;
            int left_e = 160;
            int height = 15;
            int top = 6;
            //int width_s = 121;
            //int width_e = 148;
            int width_s = 139;
            int width_e = 140;

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "1_U_S"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("TransBeginDate"));
            formItems.Add("LinkTo", "1_U_E");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "1_U_E"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_WBLD");
            formItems.Add("Alias", "U_begDate");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "beginTimeS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("TransBeginTime"));
            formItems.Add("LinkTo", "beginTimeE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "beginTimeE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_WBLD");
            formItems.Add("Alias", "U_beginTime");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "2_U_S"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("ActivateDate")); //გააქტიურების თარიღი
            formItems.Add("LinkTo", "2_U_E");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "2_U_E"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_WBLD");
            formItems.Add("Alias", "U_actDate");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "3_U_S"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("WaybillID"));
            formItems.Add("LinkTo", "3_U_E");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "3_U_E"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_WBLD");
            formItems.Add("Alias", "U_wblID");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "4_U_S"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("WaybillNumber"));
            formItems.Add("LinkTo", "4_U_E");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "4_U_E"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_WBLD");
            formItems.Add("Alias", "U_number");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "5_U_S"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("WaybillStatus"));
            formItems.Add("LinkTo", "5_U_C");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            List<string> listValidValues = new List<string>();
            listValidValues.Add(""); //-1
            listValidValues.Add(BDOSResources.getTranslate("Saved")); //1
            listValidValues.Add(BDOSResources.getTranslate("Active")); //2
            listValidValues.Add(BDOSResources.getTranslate("finished")); //3
            listValidValues.Add(BDOSResources.getTranslate("deleted")); //4
            listValidValues.Add(BDOSResources.getTranslate("Canceled")); //5
            listValidValues.Add(BDOSResources.getTranslate("SentToTransporter")); //6

            formItems = new Dictionary<string, object>();
            itemName = "5_U_C"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_WBLD");
            formItems.Add("Alias", "U_status");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Description", BDOSResources.getTranslate("WaybillStatus"));
            formItems.Add("Enabled", false);
            formItems.Add("ValidValues", listValidValues);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "6_U_S"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("WaybillType"));
            formItems.Add("LinkTo", "6_U_C");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            listValidValues = new List<string>();
            listValidValues.Add(BDOSResources.getTranslate("WithTransport")); //0 //ტრანსპორტირებით
            listValidValues.Add(BDOSResources.getTranslate("WithoutTransport")); //1 /ტრანსპორტირების გარეშე

            formItems = new Dictionary<string, object>();
            itemName = "6_U_C"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_WBLD");
            formItems.Add("Alias", "U_type");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Description", BDOSResources.getTranslate("WaybillType"));
            formItems.Add("ValidValues", listValidValues);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            left_s = 300;
            left_e = left_s + 121;
            top = 6;

            formItems = new Dictionary<string, object>();
            itemName = "7_U_S"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s / 3);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Number"));
            formItems.Add("LinkTo", "8_U_C");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "8_U_C"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_WBLD");
            formItems.Add("Alias", "Series");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            formItems.Add("Left", left_s + width_s / 3);
            formItems.Add("Width", width_s * 2 / 3);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Description", BDOSResources.getTranslate("Series"));

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "9_U_E"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_WBLD");
            formItems.Add("Alias", "DocNum");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "10_U_S"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Status"));
            formItems.Add("LinkTo", "10_U_C");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "10_U_C"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_WBLD");
            formItems.Add("Alias", "Status");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Description", BDOSResources.getTranslate("Status"));

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "11_U_S"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("CreateDate"));
            formItems.Add("LinkTo", "11_U_E");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "11_U_E"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_WBLD");
            formItems.Add("Alias", "CreateDate");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "12_U_S"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("UpdateDate"));
            formItems.Add("LinkTo", "12_U_E");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "12_U_E"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_WBLD");
            formItems.Add("Alias", "UpdateDate");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 1;

            //საფუძველი დოკუმენტის ტიპის მიხედვით CFL- ის დამატება ---->           
            bool multiSelection = false;
            string objectType = "13"; //A/R Invoice
            string uniqueID_BaseDocCFL = "BaseDoc_CFL" + objectType;
            FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_BaseDocCFL);

            objectType = "15"; //A/R Credit Memo
            uniqueID_BaseDocCFL = "BaseDoc_CFL" + objectType;
            FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_BaseDocCFL);

            objectType = "165"; //A/R Correction Invoice
            uniqueID_BaseDocCFL = "BaseDoc_CFL" + objectType;
            FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_BaseDocCFL);

            objectType = "67"; //Inventory Transfer
            uniqueID_BaseDocCFL = "BaseDoc_CFL" + objectType;
            FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_BaseDocCFL);

            objectType = "UDO_F_BDOSFASTRD_D"; //Inventory Transfer
            uniqueID_BaseDocCFL = "BaseDoc_CFL" + objectType;
            FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_BaseDocCFL);

            objectType = "14"; //A/R Credit Memo
            uniqueID_BaseDocCFL = "BaseDoc_CFL" + objectType;
            FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_BaseDocCFL);

            objectType = "60"; //Goods Issue
            uniqueID_BaseDocCFL = "BaseDoc_CFL" + objectType;
            FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_BaseDocCFL);
            //<----

            formItems = new Dictionary<string, object>();
            itemName = "13_U_E"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_WBLD");
            formItems.Add("Alias", "U_baseDocT");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Visible", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "13_U_S"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", "");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "14_U_E"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_WBLD");
            formItems.Add("Alias", "U_baseDTxt");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            //formItems.Add("ChooseFromListUID", uniqueID_BaseDocCFL);
            //formItems.Add("ChooseFromListAlias", "DocEntry");
            formItems.Add("Enabled", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "14_U_E1"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_WBLD");
            formItems.Add("Alias", "U_baseDoc");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            //formItems.Add("ChooseFromListUID", uniqueID_BaseDocCFL);
            //formItems.Add("ChooseFromListAlias", "DocEntry");
            //formItems.Add("Enabled", false);
            formItems.Add("Visible", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "14_U_LB"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            formItems.Add("Left", left_e - 20);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("LinkTo", "14_U_E");
            //formItems.Add("LinkedObjectType", objectType);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            left_s = 6;
            left_e = 160;

            //Address section ---->

            top = 150;

            formItems = new Dictionary<string, object>();
            itemName = "15_U_S"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s - 5);
            formItems.Add("Width", width_s + 50);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("AddressSection"));
            formItems.Add("TextStyle", 4);
            formItems.Add("FontSize", 12);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + 25;

            formItems = new Dictionary<string, object>();
            itemName = "16_U_S"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("StartAddress"));
            formItems.Add("LinkTo", "16_U_E");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "16_U_E"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_WBLD");
            formItems.Add("Alias", "U_strAddrs");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "34_U_E"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources"); //DBDataSources
            formItems.Add("TableName", "@BDO_WBLD");
            formItems.Add("Alias", "U_cardCode");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Visible", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            multiSelection = false;
            objectType = "171"; //SAPbouiCOM.BoLinkedObject.lf_Employee 
            string uniqueID_lf_EmployeeCFL = "Employee_CFL";
            FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_lf_EmployeeCFL);

            formItems = new Dictionary<string, object>(); //გამშვები  //ჩამბარებელი თანამშრომელი
            itemName = "17_U_S"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("ReceptionInfo"));
            formItems.Add("LinkTo", "17_U_E");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "17_U_E"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_WBLD");
            formItems.Add("Alias", "U_recpInfo");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e / 4);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("ChooseFromListUID", uniqueID_lf_EmployeeCFL);
            formItems.Add("ChooseFromListAlias", "empID");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "17_U_E1"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_WBLD");
            formItems.Add("Alias", "U_recpInfN");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e + width_e / 4);
            formItems.Add("Width", width_e * 3 / 4);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "17_U_LB"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            formItems.Add("Left", left_e - 20);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("LinkTo", "17_U_E");
            formItems.Add("LinkedObjectType", objectType);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            left_s = 300;
            left_e = left_s + 121;
            top = 150;
            top = top + 25;

            formItems = new Dictionary<string, object>();
            itemName = "19_U_S"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("EndAddress"));
            formItems.Add("LinkTo", "19_U_E");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "19_U_E"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_WBLD");
            formItems.Add("Alias", "U_endAddrs");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 1;

            multiSelection = false;
            objectType = "11";
            string uniqueID_lf_ContactCFL = "Contact_CFL";
            FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_lf_ContactCFL);

            formItems = new Dictionary<string, object>(); //მიმღები //საკონტაქტო პირი
            itemName = "20_U_S"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("ReceiverInfo"));
            formItems.Add("LinkTo", "20_U_E");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "20_U_E"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_WBLD");
            formItems.Add("Alias", "U_recvInfo");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e / 4);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            //formItems.Add("ChooseFromListUID", uniqueID_lf_ContactCFL);
            //formItems.Add("ChooseFromListAlias", "Name");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "20_U_E1"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_WBLD");
            formItems.Add("Alias", "U_recvInfN");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e + width_e / 4);
            formItems.Add("Width", width_e * 3 / 4);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "21_U_S"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("DeliveryDate")); //მიწოდების თარიღი
            formItems.Add("LinkTo", "21_U_E");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "21_U_E"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_WBLD");
            formItems.Add("Alias", "U_delvDate");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            left_s = 6;
            left_e = 160;

            //Transport section ---->

            top = 255;

            formItems = new Dictionary<string, object>();
            itemName = "22_U_S"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s - 5);
            formItems.Add("Width", width_s + 50);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("TransportSection"));
            formItems.Add("TextStyle", 4);
            formItems.Add("FontSize", 12);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + 25;

            formItems = new Dictionary<string, object>();
            itemName = "23_U_S"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("TransportType"));
            formItems.Add("LinkTo", "23_U_C");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            listValidValues = new List<string>();
            listValidValues.Add(""); //-1
            listValidValues.Add(BDOSResources.getTranslate("Auto")); //1 //საავტომობილო
            listValidValues.Add(BDOSResources.getTranslate("Railway")); //2 //სარკინიგზო
            listValidValues.Add(BDOSResources.getTranslate("Aviation")); //3 //საავიაციო
            listValidValues.Add(BDOSResources.getTranslate("other")); //4 //სხვა
            listValidValues.Add(BDOSResources.getTranslate("AutoOtherCountry")); //5 //საავტომობილო - უცხო ქვეყნის
            listValidValues.Add(BDOSResources.getTranslate("AutoTransporter")); //6 //გადამზიდავი - საავტომობილო

            formItems = new Dictionary<string, object>();
            itemName = "23_U_C"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_WBLD");
            formItems.Add("Alias", "U_trnsType");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Description", BDOSResources.getTranslate("TransportType"));
            formItems.Add("ValidValues", listValidValues);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "24_U_S"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Vehicle"));
            formItems.Add("LinkTo", "24_U_E");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            multiSelection = false;
            objectType = "UDO_F_BDO_VECL_D";
            string uniqueID_VehicleCodeCFL = "VehicleCode_CFL";
            FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_VehicleCodeCFL);

            formItems = new Dictionary<string, object>();
            itemName = "24_U_E"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_WBLD");
            formItems.Add("Alias", "U_vehicle");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e / 2);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("ChooseFromListUID", uniqueID_VehicleCodeCFL);
            formItems.Add("ChooseFromListAlias", "Code");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "24_U_LB"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            formItems.Add("Left", left_e - 20);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("LinkTo", "24_U_E");
            formItems.Add("LinkedObjectType", objectType);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "24_U_E1"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_WBLD");
            formItems.Add("Alias", "U_vehicNum");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e + width_e / 2);
            formItems.Add("Width", width_e / 2);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "25_U_S"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("TrailerNumber"));
            formItems.Add("LinkTo", "25_U_E");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "25_U_E"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_WBLD");
            formItems.Add("Alias", "U_trailNum");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            left_s = 300;
            left_e = left_s + 121;
            top = 255 + 25;

            multiSelection = false;
            objectType = "UDO_F_BDO_DRVS_D";
            string uniqueID_DriverCodeCFL = "DriverCode_CFL";
            FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_DriverCodeCFL);

            formItems = new Dictionary<string, object>();
            itemName = "26_U_S"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("DriverCode"));
            formItems.Add("LinkTo", "26_U_E");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "26_U_E"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources"); //DBDataSources
            formItems.Add("TableName", "@BDO_WBLD");
            formItems.Add("Alias", "U_drvCode");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("ChooseFromListUID", uniqueID_DriverCodeCFL);
            formItems.Add("ChooseFromListAlias", "Code");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "26_U_LB"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            formItems.Add("Left", left_e - 20);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("LinkTo", "26_U_E");
            formItems.Add("LinkedObjectType", objectType);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "26_U_B"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            formItems.Add("Left", left_e + width_e);
            formItems.Add("Width", 20);
            formItems.Add("Top", top - 2);
            formItems.Add("Height", 20);
            formItems.Add("UID", itemName);
            formItems.Add("Image", "CHOOSE_ICON");
            formItems.Add("DisplayDesc", true);
            formItems.Add("ChooseFromListUID", uniqueID_DriverCodeCFL);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "27_U_S"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("DriverTin"));
            formItems.Add("LinkTo", "27_U_E");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "27_U_E"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_WBLD");
            formItems.Add("Alias", "U_drivTin");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "28_U_CH"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_WBLD");
            formItems.Add("Alias", "U_notRsdnt");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_e + 45);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Caption", BDOSResources.getTranslate("DriverNotResident"));
            formItems.Add("ValOff", "N");
            formItems.Add("ValOn", "Y");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "29_U_S"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("TransportationExpense"));
            formItems.Add("LinkTo", "29_U_E");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "29_U_E"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_WBLD");
            formItems.Add("Alias", "U_trnsExpn");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e / 2);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            listValidValues = new List<string>();
            listValidValues.Add(""); //-1
            listValidValues.Add(BDOSResources.getTranslate("Buyer")); //1 //მყიდველი
            listValidValues.Add(BDOSResources.getTranslate("Seller")); //2 //გამყიდველი

            formItems = new Dictionary<string, object>();
            itemName = "29_U_C"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_WBLD");
            formItems.Add("Alias", "U_payForTr");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            formItems.Add("Left", left_e + width_e / 2);
            formItems.Add("Width", width_e / 2);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Description", BDOSResources.getTranslate("PaymentForTransportation"));
            formItems.Add("ValidValues", listValidValues);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "30_U_S"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Transporter"));
            formItems.Add("LinkTo", "30_U_E");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            multiSelection = false;
            objectType = "2"; //SAPbouiCOM.BoLinkedObject.lf_BusinessPartner, Business Partner object 
            string uniqueID_lf_BusinessPartnerCFL = "BusinessPartner_CFL";
            FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_lf_BusinessPartnerCFL);

            //პირობის დადება ბიზნესპარტნიორის არჩევის სიაზე
            SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(uniqueID_lf_BusinessPartnerCFL);
            SAPbouiCOM.Conditions oCons = oCFL.GetConditions();
            SAPbouiCOM.Condition oCon = oCons.Add();
            oCon.Alias = "CardType";
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCon.CondVal = "S"; //მომწოდებელი
            oCFL.SetConditions(oCons);

            formItems = new Dictionary<string, object>();
            itemName = "30_U_E"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_WBLD");
            formItems.Add("Alias", "U_tporter");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e / 4);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("ChooseFromListUID", uniqueID_lf_BusinessPartnerCFL);
            formItems.Add("ChooseFromListAlias", "CardCode");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "30_U_E1"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_WBLD");
            formItems.Add("Alias", "U_tporterN");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e + width_e / 4);
            formItems.Add("Width", width_e * 3 / 4);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "30_U_LB"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            formItems.Add("Left", left_e - 20);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("LinkTo", "30_U_E");
            formItems.Add("LinkedObjectType", objectType);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "31_U_S"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("TransporterTin"));
            formItems.Add("LinkTo", "31_U_E");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "31_U_E"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_WBLD");
            formItems.Add("Alias", "U_tporterT");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            //სარდაფი
            left_s = 6;
            left_e = 160;
            top = top + 25;

            formItems = new Dictionary<string, object>();
            itemName = "32_U_S"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Creator"));
            formItems.Add("LinkTo", "32_U_E");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "32_U_E"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_WBLD");
            formItems.Add("Alias", "Creator");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "33_U_S"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Remarks"));
            formItems.Add("LinkTo", "33_U_E");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "33_U_E"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_WBLD");
            formItems.Add("Alias", "Remark");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EXTEDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", 3 * height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("ScrollBars", SAPbouiCOM.BoScrollBars.sb_Vertical);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + 3 * height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "18_U_S"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Comment"));
            formItems.Add("LinkTo", "18_U_E");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "18_U_E"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_WBLD");
            formItems.Add("Alias", "U_comment");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e * 3);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            //ღილაკები

            left_s = 300;
            left_e = left_s + 121;

            top = top + height + 1;

            formItems = new Dictionary<string, object>();

            listValidValues = new List<string>();
            listValidValues.Add(BDOSResources.getTranslate("RSCreate"));
            listValidValues.Add(BDOSResources.getTranslate("RSActivation"));
            listValidValues.Add(BDOSResources.getTranslate("RSSendToTransporter"));
            listValidValues.Add(BDOSResources.getTranslate("RSCorrection"));
            listValidValues.Add(BDOSResources.getTranslate("RSFinish"));
            listValidValues.Add(BDOSResources.getTranslate("RSCancel"));
            listValidValues.Add(BDOSResources.getTranslate("RSUpdateStatus"));

            formItems = new Dictionary<string, object>();
            itemName = "33_U_BC";
            formItems.Add("Caption", BDOSResources.getTranslate("Operations"));
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON_COMBO);
            formItems.Add("Left", left_e);
            formItems.Add("Width", 100);
            formItems.Add("Top", top);
            formItems.Add("Height", 19);
            formItems.Add("UID", itemName);
            formItems.Add("ValidValues", listValidValues);
            formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
            formItems.Add("AffectsFormMode", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            GC.Collect();
        }

        public static void fillNewDocument(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            int docEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("@BDO_WBLD").GetValue("DocEntry", 0).Trim());

            if (docEntry != 0)
            {
                return;
            }

            Dictionary<string, string> rsSettings = CompanyDetails.getRSSettings(out errorText);
            if (errorText != null)
            {
                return;
            }

            string waybillType = rsSettings["WaybillType"];
            oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_type", 0, waybillType);
            oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_status", 0, "-1");
            oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_payForTr", 0, "-1");
            oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_trnsType", 0, waybillType == "0" ? "1" : "-1");
            oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_baseDocT", 0, ""); //SAPbouiCOM.BoLinkedObject.lf_Invoice
            oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_begDate", 0, DateTime.Today.ToString("yyyyMMdd"));
        }

        public static void createDocument(string objectType, int baseDocEntry, string vehicleCode, string driverCode, string trnsType, string trnsprter, out int newDocEntry, out string errorText)
        {
            errorText = null;
            newDocEntry = 0;
            if (canCreateDocument(baseDocEntry, objectType))
            {
                if (objectType == "13")  //A/R Invoice
                {
                    createDocumentARInvoiceType(baseDocEntry, vehicleCode, driverCode, trnsType, trnsprter, out newDocEntry, out errorText);
                }
                if (objectType == "15")  //Delivery
                {
                    createDocumentDeliveryType(baseDocEntry, vehicleCode, driverCode, trnsType, trnsprter, out newDocEntry, out errorText);
                }

                else if (objectType == "67") //Inventory Transfer
                {
                    createDocumentInventoryTransferType(baseDocEntry, vehicleCode, driverCode, trnsType, trnsprter, out newDocEntry, out errorText);
                }

                else if (objectType == "UDO_F_BDOSFASTRD_D") //Fixed Asset Transfer
                {
                    createDocumentFixedAssetTransferType(baseDocEntry, vehicleCode, driverCode, trnsType, trnsprter, out newDocEntry, out errorText);
                }

                else if (objectType == "14") //A/R Credit Memo 
                {
                    createDocumentInvoiceCreditMemoType(baseDocEntry, vehicleCode, driverCode, trnsType, trnsprter, out newDocEntry, out errorText);
                }
                else if (objectType == "60")  //Goods Issue
                {
                    createDocumentGoodsIssueType(baseDocEntry, vehicleCode, driverCode, trnsType, trnsprter, out newDocEntry, out errorText);
                }
                else if (objectType == "165")  //A/R Correction Invoice
                {
                    CreateDocumentArCorrectionInvoiceType(baseDocEntry, vehicleCode, driverCode, trnsType, trnsprter, out newDocEntry, out errorText);
                }
            }
        }

        private static void createDocumentARInvoiceType(int baseDocEntry, string vehicleCode, string driverCode, string trnsType, string trnsprter, out int newDocEntry, out string errorText)
        {
            errorText = null;
            newDocEntry = 0;

            Recordset oRecordSet = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            string query = "SELECT \"OINV\".\"DocEntry\", \"OINV\".\"ObjType\", \"OINV\".\"CardCode\", \"OINV\".\"Address2\", \"OINV\".\"DocDate\", \"OINV\".\"CntctCode\" , \"OCPR\".\"Name\"" +
                               " FROM \"OINV\" AS \"OINV\"" +
                               " LEFT JOIN \"OCPR\" AS \"OCPR\"" +
                               " ON \"OINV\".\"CntctCode\" = \"OCPR\".\"CntctCode\"" +
                               " WHERE \"DocEntry\" = '" + baseDocEntry + "'";

            string strAddrs = "";
            string queryWhs = "SELECT DISTINCT \"INV1\".\"WhsCode\", \"OWHS\".\"Street\", \"OWHS\".\"ZipCode\", \"OWHS\".\"City\", \"OWHS\".\"County\"" +
                           " FROM \"INV1\" AS \"INV1\"" +
                           " INNER JOIN \"OWHS\" AS \"OWHS\"" +
                           " ON \"INV1\".\"WhsCode\" = \"OWHS\".\"WhsCode\"" +
                           " WHERE \"INV1\".\"DocEntry\" = '" + baseDocEntry + "'";
            try
            {
                oRecordSet.DoQuery(queryWhs);
                if (oRecordSet.RecordCount == 1)
                {
                    while (!oRecordSet.EoF)
                    {
                        strAddrs = strAddrs + oRecordSet.Fields.Item("Street").Value.ToString() + '\n' +
                            oRecordSet.Fields.Item("ZipCode").Value.ToString() + " " +
                            oRecordSet.Fields.Item("City").Value.ToString() + '\n' +
                            oRecordSet.Fields.Item("County").Value.ToString();
                        oRecordSet.MoveNext();
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
                return;
            }

            try
            {
                oRecordSet.DoQuery(query);

                while (!oRecordSet.EoF)
                {
                    CompanyService oCompanyService = null;
                    GeneralService oGeneralService = null;
                    GeneralData oGeneralData = null;
                    oCompanyService = oCompany.GetCompanyService();
                    oGeneralService = oCompanyService.GetGeneralService("UDO_F_BDO_WBLD_D");
                    oGeneralData = ((GeneralData)(oGeneralService.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralData)));

                    string empID;
                    string empName;
                    Users.getUserEmployee(out empID, out empName, out errorText);

                    string cntctCode = oRecordSet.Fields.Item("CntctCode").Value.ToString();

                    oGeneralData.SetProperty("U_endAddrs", oRecordSet.Fields.Item("Address2").Value.ToString());
                    oGeneralData.SetProperty("U_delvDate", oRecordSet.Fields.Item("DocDate").Value.ToString());
                    oGeneralData.SetProperty("U_recpInfo", empID == "0" ? "" : empID); //ჩამბარებელი
                    oGeneralData.SetProperty("U_recpInfN", empName);
                    oGeneralData.SetProperty("U_recvInfo", cntctCode == "0" ? "" : cntctCode); //მიმღები
                    oGeneralData.SetProperty("U_recvInfN", oRecordSet.Fields.Item("Name").Value.ToString());
                    oGeneralData.SetProperty("U_cardCode", oRecordSet.Fields.Item("CardCode").Value.ToString());
                    oGeneralData.SetProperty("U_baseDoc", oRecordSet.Fields.Item("DocEntry").Value);
                    oGeneralData.SetProperty("U_baseDTxt", oRecordSet.Fields.Item("DocEntry").Value.ToString());
                    oGeneralData.SetProperty("U_baseDocT", oRecordSet.Fields.Item("ObjType").Value.ToString());
                    oGeneralData.SetProperty("U_strAddrs", strAddrs);
                    oGeneralData.SetProperty("U_begDate", DateTime.Today);  //DateTime.Today.ToString("yyyyMMdd"));
                    oGeneralData.SetProperty("U_beginTime", DateTime.Now);

                    if (vehicleCode != null)
                    {
                        if (trnsType == "4")
                        {
                            oGeneralData.SetProperty("U_vehicNum", vehicleCode);
                        }
                        else
                        {
                            //-->სატრანსპორტო საშუალება
                            UserTable oUserTable = null;
                            oUserTable = oCompany.UserTables.Item("BDO_VECL");
                            oUserTable.GetByKey(vehicleCode);
                            oGeneralData.SetProperty("U_vehicle", vehicleCode);
                            string vehicleNumber = oUserTable.UserFields.Fields.Item("U_number").Value;
                            string vehicleTrailNum = oUserTable.UserFields.Fields.Item("U_trailNum").Value;
                            oGeneralData.SetProperty("U_vehicNum", vehicleNumber);
                            oGeneralData.SetProperty("U_trailNum", vehicleTrailNum);
                            //სატრანსპორტო საშუალება<--

                            //-->მძღოლი                     
                            if (driverCode == null || driverCode == "")
                            {
                                driverCode = oUserTable.UserFields.Fields.Item("U_drvCode").Value;
                            }
                            oUserTable = oCompany.UserTables.Item("BDO_DRVS");
                            oUserTable.GetByKey(driverCode);
                            string driverTin = oUserTable.UserFields.Fields.Item("U_tin").Value;
                            string driverNotResident = oUserTable.UserFields.Fields.Item("U_notRsdnt").Value;
                            oGeneralData.SetProperty("U_drvCode", driverCode);
                            oGeneralData.SetProperty("U_notRsdnt", driverNotResident);
                            oGeneralData.SetProperty("U_drivTin", driverTin);
                            //მძღოლი<--
                        }
                    }

                    if (trnsprter != null)
                    {
                        Recordset oRecordSetBP = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                        string queryBP = "SELECT \"OCRD\".\"CardName\" AS \"CardName\", \"OCRD\".\"LicTradNum\" AS \"LicTradNum\" FROM \"OCRD\" WHERE \"CardCode\" = N'" + trnsprter + "'";
                        oRecordSetBP.DoQuery(queryBP);

                        if (!oRecordSetBP.EoF)
                        {
                            oGeneralData.SetProperty("U_tporter", trnsprter);
                            oGeneralData.SetProperty("U_tporterN", oRecordSetBP.Fields.Item("CardName").Value);
                            oGeneralData.SetProperty("U_tporterT", oRecordSetBP.Fields.Item("LicTradNum").Value);
                        }
                    }

                    string waybillType = null;
                    if (trnsType == null)
                    {
                        Dictionary<string, string> rsSettings = CompanyDetails.getRSSettings(out errorText);

                        if (errorText == null)
                        {
                            waybillType = rsSettings["WaybillType"];
                        }
                        trnsType = waybillType == "0" ? "1" : "-1";
                    }
                    else if (trnsType == "7") //ტრანსპორტირების გარეშე
                    {
                        waybillType = "1";
                        trnsType = "-1";
                    }
                    else if (trnsType == "6") //გადამზიდავი საავტომობილო
                    {
                        waybillType = "0";
                        trnsType = "6";
                    }
                    else
                    {
                        waybillType = "0"; //ტრანსპორტირებით
                    }
                    oGeneralData.SetProperty("U_type", waybillType);
                    oGeneralData.SetProperty("U_status", "-1");
                    oGeneralData.SetProperty("U_payForTr", "-1");
                    oGeneralData.SetProperty("U_trnsType", trnsType);

                    try
                    {
                        var response = oGeneralService.Add(oGeneralData);
                        var docEntry = response.GetProperty("DocEntry");
                        newDocEntry = Convert.ToInt32(docEntry);
                    }
                    catch (Exception ex)
                    {
                        int errCode;
                        string errMsg;

                        oCompany.GetLastError(out errCode, out errMsg);
                        errorText = ex.Message;
                        errorText = BDOSResources.getTranslate("ErrorOfDocumentAdd") + " " + BDOSResources.getTranslate("ErrorDescription") + " " + errMsg + "! " + BDOSResources.getTranslate("Code") + " : " + errCode + "! " + errorText;
                    }

                    oRecordSet.MoveNext();
                    break;
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
                return;
            }
            finally
            {
                GC.Collect();
            }
        }

        private static void createDocumentDeliveryType(int baseDocEntry, string vehicleCode, string driverCode, string trnsType, string trnsprter, out int newDocEntry, out string errorText)
        {
            errorText = null;
            newDocEntry = 0;

            Recordset oRecordSet = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            string query = "SELECT \"ODLN\".\"DocEntry\", \"ODLN\".\"ObjType\", \"ODLN\".\"CardCode\", \"ODLN\".\"Address2\", \"ODLN\".\"DocDate\", \"ODLN\".\"CntctCode\" , \"OCPR\".\"Name\"" +
                               " FROM \"ODLN\" AS \"ODLN\"" +
                               " LEFT JOIN \"OCPR\" AS \"OCPR\"" +
                               " ON \"ODLN\".\"CntctCode\" = \"OCPR\".\"CntctCode\"" +
                               " WHERE \"DocEntry\" = '" + baseDocEntry + "'";

            string strAddrs = "";
            string queryWhs = "SELECT DISTINCT \"DLN1\".\"WhsCode\", \"OWHS\".\"Street\", \"OWHS\".\"ZipCode\", \"OWHS\".\"City\", \"OWHS\".\"County\"" +
                           " FROM \"DLN1\" AS \"DLN1\"" +
                           " INNER JOIN \"OWHS\" AS \"OWHS\"" +
                           " ON \"DLN1\".\"WhsCode\" = \"OWHS\".\"WhsCode\"" +
                           " WHERE \"DLN1\".\"DocEntry\" = '" + baseDocEntry + "'";
            try
            {
                oRecordSet.DoQuery(queryWhs);
                if (oRecordSet.RecordCount == 1)
                {
                    while (!oRecordSet.EoF)
                    {
                        strAddrs = strAddrs + oRecordSet.Fields.Item("Street").Value.ToString() + '\n' +
                            oRecordSet.Fields.Item("ZipCode").Value.ToString() + " " +
                            oRecordSet.Fields.Item("City").Value.ToString() + '\n' +
                            oRecordSet.Fields.Item("County").Value.ToString();
                        oRecordSet.MoveNext();
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
                return;
            }

            try
            {
                oRecordSet.DoQuery(query);

                while (!oRecordSet.EoF)
                {
                    CompanyService oCompanyService = null;
                    GeneralService oGeneralService = null;
                    GeneralData oGeneralData = null;
                    oCompanyService = oCompany.GetCompanyService();
                    oGeneralService = oCompanyService.GetGeneralService("UDO_F_BDO_WBLD_D");
                    oGeneralData = ((GeneralData)(oGeneralService.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralData)));

                    string empID;
                    string empName;
                    Users.getUserEmployee(out empID, out empName, out errorText);

                    string cntctCode = oRecordSet.Fields.Item("CntctCode").Value.ToString();

                    oGeneralData.SetProperty("U_endAddrs", oRecordSet.Fields.Item("Address2").Value.ToString());
                    oGeneralData.SetProperty("U_delvDate", oRecordSet.Fields.Item("DocDate").Value.ToString());
                    oGeneralData.SetProperty("U_recpInfo", empID == "0" ? "" : empID); //ჩამბარებელი
                    oGeneralData.SetProperty("U_recpInfN", empName);
                    oGeneralData.SetProperty("U_recvInfo", cntctCode == "0" ? "" : cntctCode); //მიმღები
                    oGeneralData.SetProperty("U_recvInfN", oRecordSet.Fields.Item("Name").Value.ToString());
                    oGeneralData.SetProperty("U_cardCode", oRecordSet.Fields.Item("CardCode").Value.ToString());
                    oGeneralData.SetProperty("U_baseDoc", oRecordSet.Fields.Item("DocEntry").Value);
                    oGeneralData.SetProperty("U_baseDTxt", oRecordSet.Fields.Item("DocEntry").Value.ToString());
                    oGeneralData.SetProperty("U_baseDocT", oRecordSet.Fields.Item("ObjType").Value.ToString());
                    oGeneralData.SetProperty("U_strAddrs", strAddrs);
                    oGeneralData.SetProperty("U_begDate", DateTime.Today);  //DateTime.Today.ToString("yyyyMMdd"));
                    oGeneralData.SetProperty("U_beginTime", DateTime.Now);

                    if (vehicleCode != null)
                    {
                        if (trnsType == "4")
                        {
                            oGeneralData.SetProperty("U_vehicNum", vehicleCode);
                        }
                        else
                        {
                            //-->სატრანსპორტო საშუალება
                            UserTable oUserTable = null;
                            oUserTable = oCompany.UserTables.Item("BDO_VECL");
                            oUserTable.GetByKey(vehicleCode);
                            oGeneralData.SetProperty("U_vehicle", vehicleCode);
                            string vehicleNumber = oUserTable.UserFields.Fields.Item("U_number").Value;
                            string vehicleTrailNum = oUserTable.UserFields.Fields.Item("U_trailNum").Value;
                            oGeneralData.SetProperty("U_vehicNum", vehicleNumber);
                            oGeneralData.SetProperty("U_trailNum", vehicleTrailNum);
                            //სატრანსპორტო საშუალება<--

                            //-->მძღოლი                     
                            if (driverCode == null || driverCode == "")
                            {
                                driverCode = oUserTable.UserFields.Fields.Item("U_drvCode").Value;
                            }
                            oUserTable = oCompany.UserTables.Item("BDO_DRVS");
                            oUserTable.GetByKey(driverCode);
                            string driverTin = oUserTable.UserFields.Fields.Item("U_tin").Value;
                            string driverNotResident = oUserTable.UserFields.Fields.Item("U_notRsdnt").Value;
                            oGeneralData.SetProperty("U_drvCode", driverCode);
                            oGeneralData.SetProperty("U_notRsdnt", driverNotResident);
                            oGeneralData.SetProperty("U_drivTin", driverTin);
                            //მძღოლი<--
                        }
                    }

                    if (trnsprter != null)
                    {
                        Recordset oRecordSetBP = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                        string queryBP = "SELECT \"OCRD\".\"CardName\" AS \"CardName\", \"OCRD\".\"LicTradNum\" AS \"LicTradNum\" FROM \"OCRD\" WHERE \"CardCode\" = N'" + trnsprter + "'";
                        oRecordSetBP.DoQuery(queryBP);

                        if (!oRecordSetBP.EoF)
                        {
                            oGeneralData.SetProperty("U_tporter", trnsprter);
                            oGeneralData.SetProperty("U_tporterN", oRecordSetBP.Fields.Item("CardName").Value);
                            oGeneralData.SetProperty("U_tporterT", oRecordSetBP.Fields.Item("LicTradNum").Value);
                        }
                    }

                    string waybillType = null;
                    if (trnsType == null)
                    {
                        Dictionary<string, string> rsSettings = CompanyDetails.getRSSettings(out errorText);

                        if (errorText == null)
                        {
                            waybillType = rsSettings["WaybillType"];
                        }
                        trnsType = waybillType == "0" ? "1" : "-1";
                    }
                    else if (trnsType == "7") //ტრანსპორტირების გარეშე
                    {
                        waybillType = "1";
                        trnsType = "-1";
                    }
                    else if (trnsType == "6") //გადამზიდავი საავტომობილო
                    {
                        waybillType = "0";
                        trnsType = "6";
                    }
                    else
                    {
                        waybillType = "0"; //ტრანსპორტირებით
                    }
                    oGeneralData.SetProperty("U_type", waybillType);
                    oGeneralData.SetProperty("U_status", "-1");
                    oGeneralData.SetProperty("U_payForTr", "-1");
                    oGeneralData.SetProperty("U_trnsType", trnsType);

                    try
                    {
                        var response = oGeneralService.Add(oGeneralData);
                        var docEntry = response.GetProperty("DocEntry");
                        newDocEntry = Convert.ToInt32(docEntry);
                    }
                    catch (Exception ex)
                    {
                        int errCode;
                        string errMsg;

                        oCompany.GetLastError(out errCode, out errMsg);
                        errorText = ex.Message;
                        errorText = BDOSResources.getTranslate("ErrorOfDocumentAdd") + " " + BDOSResources.getTranslate("ErrorDescription") + " " + errMsg + "! " + BDOSResources.getTranslate("Code") + " : " + errCode + "! " + errorText;
                    }

                    oRecordSet.MoveNext();
                    break;
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
                return;
            }
            finally
            {
                GC.Collect();
            }
        }

        private static void createDocumentInventoryTransferType(int baseDocEntry, string vehicleCode, string driverCode, string trnsType, string trnsprter, out int newDocEntry, out string errorText)
        {
            errorText = null;
            newDocEntry = 0;

            Recordset oRecordSet = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            string query = "SELECT \"OWTR\".\"DocEntry\", \"OWTR\".\"ObjType\", \"OWTR\".\"CardCode\", \"OWTR\".\"Address\", \"OWTR\".\"DocDate\", \"OWTR\".\"CntctCode\" , \"OCPR\".\"Name\"" +
                               " FROM \"OWTR\" AS \"OWTR\"" +
                               " LEFT JOIN \"OCPR\" AS \"OCPR\"" +
                               " ON \"OWTR\".\"CntctCode\" = \"OCPR\".\"CntctCode\"" +
                               " WHERE \"DocEntry\" = '" + baseDocEntry + "'";

            string strAddrs = "";
            string queryStrAddrs = "SELECT \"OWTR\".\"Filler\", \"OWHS\".\"Street\", \"OWHS\".\"ZipCode\", \"OWHS\".\"City\", \"OWHS\".\"County\"" +
                           " FROM \"OWTR\" AS \"OWTR\"" +
                           " INNER JOIN \"OWHS\" AS \"OWHS\"" +
                           " ON \"OWTR\".\"Filler\" = \"OWHS\".\"WhsCode\"" +
                           " WHERE \"OWTR\".\"DocEntry\" = '" + baseDocEntry + "'";
            try
            {
                oRecordSet.DoQuery(queryStrAddrs);
                while (!oRecordSet.EoF)
                {
                    strAddrs = strAddrs + oRecordSet.Fields.Item("Street").Value.ToString() + '\n' +
                        oRecordSet.Fields.Item("ZipCode").Value.ToString() + " " +
                        oRecordSet.Fields.Item("City").Value.ToString() + '\n' +
                        oRecordSet.Fields.Item("County").Value.ToString();
                    oRecordSet.MoveNext();
                    break;
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
                return;
            }

            string endAddrs = "";
            string queryEndAddrs = "SELECT \"OWTR\".\"ToWhsCode\", \"OWHS\".\"Street\", \"OWHS\".\"ZipCode\", \"OWHS\".\"City\", \"OWHS\".\"County\"" +
                           " FROM \"OWTR\" AS \"OWTR\"" +
                           " INNER JOIN \"OWHS\" AS \"OWHS\"" +
                           " ON \"OWTR\".\"ToWhsCode\" = \"OWHS\".\"WhsCode\"" +
                           " WHERE \"OWTR\".\"DocEntry\" = '" + baseDocEntry + "'";
            try
            {
                oRecordSet.DoQuery(queryEndAddrs);
                while (!oRecordSet.EoF)
                {
                    endAddrs = endAddrs + oRecordSet.Fields.Item("Street").Value.ToString() + '\n' +
                        oRecordSet.Fields.Item("ZipCode").Value.ToString() + " " +
                        oRecordSet.Fields.Item("City").Value.ToString() + '\n' +
                        oRecordSet.Fields.Item("County").Value.ToString();
                    oRecordSet.MoveNext();
                    break;
                }
            }

            catch (Exception ex)
            {
                errorText = ex.Message;
                return;
            }

            try
            {
                oRecordSet.DoQuery(query);

                while (!oRecordSet.EoF)
                {
                    CompanyService oCompanyService = null;
                    GeneralService oGeneralService = null;
                    GeneralData oGeneralData = null;
                    oCompanyService = oCompany.GetCompanyService();
                    oGeneralService = oCompanyService.GetGeneralService("UDO_F_BDO_WBLD_D");
                    oGeneralData = ((GeneralData)(oGeneralService.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralData)));

                    string empID;
                    string empName;
                    Users.getUserEmployee(out empID, out empName, out errorText);

                    string cntctCode = oRecordSet.Fields.Item("CntctCode").Value.ToString();

                    oGeneralData.SetProperty("U_endAddrs", endAddrs);
                    oGeneralData.SetProperty("U_delvDate", oRecordSet.Fields.Item("DocDate").Value.ToString());
                    oGeneralData.SetProperty("U_recpInfo", empID == "0" ? "" : empID); //ჩამბარებელი
                    oGeneralData.SetProperty("U_recpInfN", empName);
                    oGeneralData.SetProperty("U_recvInfo", cntctCode == "0" ? "" : cntctCode); //მიმღები
                    oGeneralData.SetProperty("U_recvInfN", oRecordSet.Fields.Item("Name").Value.ToString());
                    oGeneralData.SetProperty("U_cardCode", oRecordSet.Fields.Item("CardCode").Value.ToString());
                    oGeneralData.SetProperty("U_baseDoc", oRecordSet.Fields.Item("DocEntry").Value);
                    oGeneralData.SetProperty("U_baseDTxt", oRecordSet.Fields.Item("DocEntry").Value.ToString());
                    oGeneralData.SetProperty("U_baseDocT", oRecordSet.Fields.Item("ObjType").Value.ToString());
                    oGeneralData.SetProperty("U_strAddrs", strAddrs);
                    oGeneralData.SetProperty("U_begDate", DateTime.Today);  //DateTime.Today.ToString("yyyyMMdd"));
                    oGeneralData.SetProperty("U_beginTime", DateTime.Now);

                    if (vehicleCode != null)
                    {
                        if (trnsType == "4")
                        {
                            oGeneralData.SetProperty("U_vehicNum", vehicleCode);
                        }
                        else
                        {
                            //-->სატრანსპორტო საშუალება
                            UserTable oUserTable = null;
                            oUserTable = oCompany.UserTables.Item("BDO_VECL");
                            oUserTable.GetByKey(vehicleCode);
                            oGeneralData.SetProperty("U_vehicle", vehicleCode);
                            string vehicleNumber = oUserTable.UserFields.Fields.Item("U_number").Value;
                            string vehicleTrailNum = oUserTable.UserFields.Fields.Item("U_trailNum").Value;
                            oGeneralData.SetProperty("U_vehicNum", vehicleNumber);
                            oGeneralData.SetProperty("U_trailNum", vehicleTrailNum);
                            //სატრანსპორტო საშუალება<--

                            //-->მძღოლი                     
                            if (driverCode == null || driverCode == "")
                            {
                                driverCode = oUserTable.UserFields.Fields.Item("U_drvCode").Value;
                            }
                            oUserTable = oCompany.UserTables.Item("BDO_DRVS");
                            oUserTable.GetByKey(driverCode);
                            string driverTin = oUserTable.UserFields.Fields.Item("U_tin").Value;
                            string driverNotResident = oUserTable.UserFields.Fields.Item("U_notRsdnt").Value;
                            oGeneralData.SetProperty("U_drvCode", driverCode);
                            oGeneralData.SetProperty("U_notRsdnt", driverNotResident);
                            oGeneralData.SetProperty("U_drivTin", driverTin);
                            //მძღოლი<--
                        }
                    }

                    if (trnsprter != null)
                    {
                        Recordset oRecordSetBP = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                        string queryBP = "SELECT " + "\"OCRD\".\"CardName\" AS \"CardName\", \"OCRD\".\"LicTradNum\" AS \"LicTradNum\" " + " FROM \"OCRD\"" + " WHERE " + "\"CardCode\"" + " = N'" + trnsprter + "'";
                        oRecordSetBP.DoQuery(queryBP);

                        if (!oRecordSetBP.EoF)
                        {
                            oGeneralData.SetProperty("U_tporter", trnsprter);
                            oGeneralData.SetProperty("U_tporterN", oRecordSetBP.Fields.Item("CardName").Value);
                            oGeneralData.SetProperty("U_tporterT", oRecordSetBP.Fields.Item("LicTradNum").Value);
                        }
                    }

                    trnsType = trnsType == null ? "1" : trnsType;
                    oGeneralData.SetProperty("U_type", "0"); //ტრანსპორტირებით
                    oGeneralData.SetProperty("U_status", "-1");
                    oGeneralData.SetProperty("U_payForTr", "-1");
                    oGeneralData.SetProperty("U_trnsType", trnsType); //საავტომობილო

                    try
                    {
                        var response = oGeneralService.Add(oGeneralData);
                        var docEntry = response.GetProperty("DocEntry");
                        newDocEntry = Convert.ToInt32(docEntry);
                    }
                    catch (Exception ex)
                    {
                        int errCode;
                        string errMsg;

                        oCompany.GetLastError(out errCode, out errMsg);
                        errorText = ex.Message;
                        errorText = BDOSResources.getTranslate("ErrorOfDocumentAdd") + " " + BDOSResources.getTranslate("ErrorDescription") + " " + errMsg + "! " + BDOSResources.getTranslate("Code") + " : " + errCode + "! " + errorText;
                    }

                    oRecordSet.MoveNext();
                    break;
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
                return;
            }
            finally
            {
                GC.Collect();
            }
        }

        private static void createDocumentFixedAssetTransferType(int baseDocEntry, string vehicleCode, string driverCode, string trnsType, string trnsprter, out int newDocEntry, out string errorText)
        {
            errorText = null;
            newDocEntry = 0;

            Recordset oRecordSet = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            string query = "SELECT \"FASTRD\".\"DocEntry\", \"FASTRD\".\"Object\", \"FASTRD\".\"U_CardCode\", \"FASTRD\".\"U_DocDate\", \"FASTRD\".\"U_TEmplID\" , \"OCPR\".\"Name\"" +
                               " FROM \"@BDOSFASTRD\" AS \"FASTRD\"" +
                               " LEFT JOIN \"OCPR\" AS \"OCPR\"" +
                               " ON \"FASTRD\".\"U_TEmplID\" = \"OCPR\".\"CntctCode\"" +
                               " WHERE \"DocEntry\" = '" + baseDocEntry + "'";

            string strAddrs = "";
            string queryStrAddrs = "SELECT \"FASTRD\".\"U_FLocCode\", \"OLCT\".\"U_BDOSAddres\"" +
                           " FROM \"@BDOSFASTRD\" AS \"FASTRD\"" +
                           " INNER JOIN \"OLCT\" AS \"OLCT\"" +
                           " ON \"FASTRD\".\"U_FLocCode\" = \"OLCT\".\"Code\"" +
                           " WHERE \"FASTRD\".\"DocEntry\" = '" + baseDocEntry + "'";


            try
            {
                oRecordSet.DoQuery(queryStrAddrs);
                while (!oRecordSet.EoF)
                {
                    strAddrs = strAddrs + oRecordSet.Fields.Item("U_BDOSAddres").Value.ToString();
                    oRecordSet.MoveNext();
                    break;
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
                return;
            }

            string endAddrs = "";
            string queryEndAddrs = "SELECT \"FASTRD\".\"U_TLocCode\", \"OLCT\".\"U_BDOSAddres\"" +
                           " FROM \"@BDOSFASTRD\" AS \"FASTRD\"" +
                           " INNER JOIN \"OLCT\" AS \"OLCT\"" +
                           " ON \"FASTRD\".\"U_TLocCode\" = \"OLCT\".\"Code\"" +
                           " WHERE \"FASTRD\".\"DocEntry\" = '" + baseDocEntry + "'";
            try
            {
                oRecordSet.DoQuery(queryEndAddrs);
                while (!oRecordSet.EoF)
                {
                    endAddrs = endAddrs + oRecordSet.Fields.Item("U_BDOSAddres").Value.ToString();
                    oRecordSet.MoveNext();
                    break;
                }
            }

            catch (Exception ex)
            {
                errorText = ex.Message;
                return;
            }

            try
            {
                oRecordSet.DoQuery(query);

                while (!oRecordSet.EoF)
                {
                    CompanyService oCompanyService = null;
                    GeneralService oGeneralService = null;
                    GeneralData oGeneralData = null;
                    oCompanyService = oCompany.GetCompanyService();
                    oGeneralService = oCompanyService.GetGeneralService("UDO_F_BDO_WBLD_D");
                    oGeneralData = ((GeneralData)(oGeneralService.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralData)));

                    string empID;
                    string empName;
                    Users.getUserEmployee(out empID, out empName, out errorText);


                    string cntctCode = oRecordSet.Fields.Item("U_TEmplID").Value.ToString();

                    oGeneralData.SetProperty("U_endAddrs", endAddrs);
                    oGeneralData.SetProperty("U_delvDate", oRecordSet.Fields.Item("U_DocDate").Value.ToString());
                    oGeneralData.SetProperty("U_recpInfo", empID == "0" ? "" : empID); //ჩამბარებელი
                    oGeneralData.SetProperty("U_recpInfN", empName);
                    oGeneralData.SetProperty("U_recvInfo", cntctCode == "0" ? "" : cntctCode); //მიმღები
                    oGeneralData.SetProperty("U_recvInfN", oRecordSet.Fields.Item("Name").Value.ToString());
                    oGeneralData.SetProperty("U_cardCode", oRecordSet.Fields.Item("U_CardCode").Value.ToString());
                    oGeneralData.SetProperty("U_baseDoc", oRecordSet.Fields.Item("DocEntry").Value);
                    oGeneralData.SetProperty("U_baseDTxt", oRecordSet.Fields.Item("DocEntry").Value.ToString());
                    oGeneralData.SetProperty("U_baseDocT", oRecordSet.Fields.Item("Object").Value.ToString());
                    oGeneralData.SetProperty("U_strAddrs", strAddrs);
                    oGeneralData.SetProperty("U_begDate", DateTime.Today);  //DateTime.Today.ToString("yyyyMMdd"));
                    oGeneralData.SetProperty("U_beginTime", DateTime.Now);

                    if (vehicleCode != null)
                    {
                        if (trnsType == "4")
                        {
                            oGeneralData.SetProperty("U_vehicNum", vehicleCode);
                        }
                        else
                        {
                            //-->სატრანსპორტო საშუალება
                            UserTable oUserTable = null;
                            oUserTable = oCompany.UserTables.Item("BDO_VECL");
                            oUserTable.GetByKey(vehicleCode);
                            oGeneralData.SetProperty("U_vehicle", vehicleCode);
                            string vehicleNumber = oUserTable.UserFields.Fields.Item("U_number").Value;
                            string vehicleTrailNum = oUserTable.UserFields.Fields.Item("U_trailNum").Value;
                            oGeneralData.SetProperty("U_vehicNum", vehicleNumber);
                            oGeneralData.SetProperty("U_trailNum", vehicleTrailNum);
                            //სატრანსპორტო საშუალება<--

                            //-->მძღოლი                     
                            if (driverCode == null || driverCode == "")
                            {
                                driverCode = oUserTable.UserFields.Fields.Item("U_drvCode").Value;
                            }
                            oUserTable = oCompany.UserTables.Item("BDO_DRVS");
                            oUserTable.GetByKey(driverCode);
                            string driverTin = oUserTable.UserFields.Fields.Item("U_tin").Value;
                            string driverNotResident = oUserTable.UserFields.Fields.Item("U_notRsdnt").Value;
                            oGeneralData.SetProperty("U_drvCode", driverCode);
                            oGeneralData.SetProperty("U_notRsdnt", driverNotResident);
                            oGeneralData.SetProperty("U_drivTin", driverTin);
                            //მძღოლი<--
                        }
                    }

                    if (trnsprter != null)
                    {
                        Recordset oRecordSetBP = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                        string queryBP = "SELECT " + "\"OCRD\".\"CardName\" AS \"CardName\", \"OCRD\".\"LicTradNum\" AS \"LicTradNum\" " + " FROM \"OCRD\"" + " WHERE " + "\"CardCode\"" + " = N'" + trnsprter + "'";
                        oRecordSetBP.DoQuery(queryBP);

                        if (!oRecordSetBP.EoF)
                        {
                            oGeneralData.SetProperty("U_tporter", trnsprter);
                            oGeneralData.SetProperty("U_tporterN", oRecordSetBP.Fields.Item("CardName").Value);
                            oGeneralData.SetProperty("U_tporterT", oRecordSetBP.Fields.Item("LicTradNum").Value);
                        }
                    }

                    trnsType = trnsType == null ? "1" : trnsType;
                    oGeneralData.SetProperty("U_type", "0"); //ტრანსპორტირებით
                    oGeneralData.SetProperty("U_status", "-1");
                    oGeneralData.SetProperty("U_payForTr", "-1");
                    oGeneralData.SetProperty("U_trnsType", trnsType); //საავტომობილო

                    try
                    {
                        var response = oGeneralService.Add(oGeneralData);
                        var docEntry = response.GetProperty("DocEntry");
                        newDocEntry = Convert.ToInt32(docEntry);
                    }
                    catch (Exception ex)
                    {
                        int errCode;
                        string errMsg;

                        oCompany.GetLastError(out errCode, out errMsg);
                        errorText = ex.Message;
                        errorText = BDOSResources.getTranslate("ErrorOfDocumentAdd") + " " + BDOSResources.getTranslate("ErrorDescription") + " " + errMsg + "! " + BDOSResources.getTranslate("Code") + " : " + errCode + "! " + errorText;
                    }

                    oRecordSet.MoveNext();
                    break;
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
                return;
            }
            finally
            {
                GC.Collect();
            }
        }

        private static void createDocumentInvoiceCreditMemoType(int baseDocEntry, string vehicleCode, string driverCode, string trnsType, string trnsprter, out int newDocEntry, out string errorText)
        {
            errorText = null;
            newDocEntry = 0;

            Recordset oRecordSet = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            string query = "SELECT \"ORIN\".\"DocEntry\", \"ORIN\".\"ObjType\", \"ORIN\".\"CardCode\", \"ORIN\".\"Address2\", \"ORIN\".\"DocDate\", \"ORIN\".\"CntctCode\" , \"OCPR\".\"Name\"" +
                               " FROM \"ORIN\" AS \"ORIN\"" +
                               " LEFT JOIN \"OCPR\" AS \"OCPR\"" +
                               " ON \"ORIN\".\"CntctCode\" = \"OCPR\".\"CntctCode\"" +
                               " WHERE \"DocEntry\" = '" + baseDocEntry + "'";

            string strAddrs = "";
            string queryWhs = "SELECT DISTINCT \"RIN1\".\"WhsCode\", \"OWHS\".\"Street\", \"OWHS\".\"ZipCode\", \"OWHS\".\"City\", \"OWHS\".\"County\"" +
                           " FROM \"RIN1\" AS \"RIN1\"" +
                           " INNER JOIN \"OWHS\" AS \"OWHS\"" +
                           " ON \"RIN1\".\"WhsCode\" = \"OWHS\".\"WhsCode\"" +
                           " WHERE \"RIN1\".\"DocEntry\" = '" + baseDocEntry + "'";
            try
            {
                oRecordSet.DoQuery(queryWhs);
                if (oRecordSet.RecordCount == 1)
                {
                    while (!oRecordSet.EoF)
                    {
                        strAddrs = strAddrs + oRecordSet.Fields.Item("Street").Value.ToString() + '\n' +
                            oRecordSet.Fields.Item("ZipCode").Value.ToString() + " " +
                            oRecordSet.Fields.Item("City").Value.ToString() + '\n' +
                            oRecordSet.Fields.Item("County").Value.ToString();
                        oRecordSet.MoveNext();
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
                return;
            }

            try
            {
                oRecordSet.DoQuery(query);

                while (!oRecordSet.EoF)
                {
                    CompanyService oCompanyService = null;
                    GeneralService oGeneralService = null;
                    GeneralData oGeneralData = null;
                    oCompanyService = oCompany.GetCompanyService();
                    oGeneralService = oCompanyService.GetGeneralService("UDO_F_BDO_WBLD_D");
                    oGeneralData = ((GeneralData)(oGeneralService.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralData)));

                    string empID;
                    string empName;
                    Users.getUserEmployee(out empID, out empName, out errorText);

                    string cntctCode = oRecordSet.Fields.Item("CntctCode").Value.ToString();

                    oGeneralData.SetProperty("U_strAddrs", oRecordSet.Fields.Item("Address2").Value.ToString()); //U_endAddrs
                    oGeneralData.SetProperty("U_delvDate", oRecordSet.Fields.Item("DocDate").Value.ToString());
                    oGeneralData.SetProperty("U_recpInfo", empID == "0" ? "" : empID); //ჩამბარებელი
                    oGeneralData.SetProperty("U_recpInfN", empName);
                    oGeneralData.SetProperty("U_recvInfo", cntctCode == "0" ? "" : cntctCode); //მიმღები
                    oGeneralData.SetProperty("U_recvInfN", oRecordSet.Fields.Item("Name").Value.ToString());
                    oGeneralData.SetProperty("U_cardCode", oRecordSet.Fields.Item("CardCode").Value.ToString());
                    oGeneralData.SetProperty("U_baseDoc", oRecordSet.Fields.Item("DocEntry").Value);
                    oGeneralData.SetProperty("U_baseDTxt", oRecordSet.Fields.Item("DocEntry").Value.ToString());
                    oGeneralData.SetProperty("U_baseDocT", oRecordSet.Fields.Item("ObjType").Value.ToString());
                    oGeneralData.SetProperty("U_endAddrs", strAddrs); //U_strAddrs
                    oGeneralData.SetProperty("U_begDate", DateTime.Today);  //DateTime.Today.ToString("yyyyMMdd"));
                    oGeneralData.SetProperty("U_beginTime", DateTime.Now);

                    if (vehicleCode != null)
                    {
                        if (trnsType == "4")
                        {
                            oGeneralData.SetProperty("U_vehicNum", vehicleCode);
                        }
                        else
                        {
                            //-->სატრანსპორტო საშუალება
                            UserTable oUserTable = null;
                            oUserTable = oCompany.UserTables.Item("BDO_VECL");
                            oUserTable.GetByKey(vehicleCode);
                            oGeneralData.SetProperty("U_vehicle", vehicleCode);
                            string vehicleNumber = oUserTable.UserFields.Fields.Item("U_number").Value;
                            string vehicleTrailNum = oUserTable.UserFields.Fields.Item("U_trailNum").Value;
                            oGeneralData.SetProperty("U_vehicNum", vehicleNumber);
                            oGeneralData.SetProperty("U_trailNum", vehicleTrailNum);
                            //სატრანსპორტო საშუალება<--

                            //-->მძღოლი                     
                            if (driverCode == null || driverCode == "")
                            {
                                driverCode = oUserTable.UserFields.Fields.Item("U_drvCode").Value;
                            }
                            oUserTable = oCompany.UserTables.Item("BDO_DRVS");
                            oUserTable.GetByKey(driverCode);
                            string driverTin = oUserTable.UserFields.Fields.Item("U_tin").Value;
                            string driverNotResident = oUserTable.UserFields.Fields.Item("U_notRsdnt").Value;
                            oGeneralData.SetProperty("U_drvCode", driverCode);
                            oGeneralData.SetProperty("U_notRsdnt", driverNotResident);
                            oGeneralData.SetProperty("U_drivTin", driverTin);
                            //მძღოლი<--
                        }
                    }

                    if (trnsprter != null)
                    {
                        Recordset oRecordSetBP = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                        string queryBP = "SELECT " + "\"OCRD\".\"CardName\" AS \"CardName\", \"OCRD\".\"LicTradNum\" AS \"LicTradNum\" " + " FROM \"OCRD\"" + " WHERE " + "\"CardCode\"" + " = N'" + trnsprter + "'";
                        oRecordSetBP.DoQuery(queryBP);

                        if (!oRecordSetBP.EoF)
                        {
                            oGeneralData.SetProperty("U_tporter", trnsprter);
                            oGeneralData.SetProperty("U_tporterN", oRecordSetBP.Fields.Item("CardName").Value);
                            oGeneralData.SetProperty("U_tporterT", oRecordSetBP.Fields.Item("LicTradNum").Value);
                        }
                    }

                    trnsType = trnsType == null ? "1" : trnsType;
                    oGeneralData.SetProperty("U_type", "0"); //ტრანსპორტირებით
                    oGeneralData.SetProperty("U_status", "-1");
                    oGeneralData.SetProperty("U_payForTr", "-1");
                    oGeneralData.SetProperty("U_trnsType", trnsType); //საავტომობილო

                    try
                    {
                        var response = oGeneralService.Add(oGeneralData);
                        var docEntry = response.GetProperty("DocEntry");
                        newDocEntry = Convert.ToInt32(docEntry);
                    }
                    catch (Exception ex)
                    {
                        int errCode;
                        string errMsg;

                        oCompany.GetLastError(out errCode, out errMsg);
                        errorText = ex.Message;
                        errorText = BDOSResources.getTranslate("ErrorOfDocumentAdd") + " " + BDOSResources.getTranslate("ErrorDescription") + " " + errMsg + "! " + BDOSResources.getTranslate("Code") + " : " + errCode + "! " + errorText;
                    }

                    oRecordSet.MoveNext();
                    break;
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
                return;
            }
            finally
            {
                GC.Collect();
            }
        }

        private static void createDocumentGoodsIssueType(int baseDocEntry, string vehicleCode, string driverCode, string trnsType, string trnsprter, out int newDocEntry, out string errorText)
        {
            errorText = null;
            newDocEntry = 0;

            Recordset oRecordSet = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            string query = "SELECT \"OIGE\".\"DocEntry\", \"OIGE\".\"ObjType\", \"OIGE\".\"Address\", \"OIGE\".\"DocDate\"" +
                               " FROM \"OIGE\" AS \"OIGE\"" +
                                " WHERE \"DocEntry\" = '" + baseDocEntry + "'";

            string strAddrs = "";
            string queryWhs = "SELECT DISTINCT \"IGE1\".\"WhsCode\", \"OWHS\".\"Street\", \"OWHS\".\"ZipCode\", \"OWHS\".\"City\", \"OWHS\".\"County\"" +
                           " FROM \"IGE1\" AS \"IGE1\"" +
                           " INNER JOIN \"OWHS\" AS \"OWHS\"" +
                           " ON \"IGE1\".\"WhsCode\" = \"OWHS\".\"WhsCode\"" +
                           " WHERE \"IGE1\".\"DocEntry\" = '" + baseDocEntry + "'";
            try
            {
                oRecordSet.DoQuery(queryWhs);
                if (oRecordSet.RecordCount == 1)
                {
                    while (!oRecordSet.EoF)
                    {
                        strAddrs = strAddrs + oRecordSet.Fields.Item("Street").Value.ToString() + '\n' +
                            oRecordSet.Fields.Item("ZipCode").Value.ToString() + " " +
                            oRecordSet.Fields.Item("City").Value.ToString() + '\n' +
                            oRecordSet.Fields.Item("County").Value.ToString();
                        oRecordSet.MoveNext();
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
                return;
            }

            string endAddrs = "";
            string queryEndAddrs = "SELECT \"OIGE\".\"ToWhsCode\", \"OWHS\".\"Street\", \"OWHS\".\"ZipCode\", \"OWHS\".\"City\", \"OWHS\".\"County\"" +
                           " FROM \"OIGE\" AS \"OIGE\"" +
                           " INNER JOIN \"OWHS\" AS \"OWHS\"" +
                           " ON \"OIGE\".\"ToWhsCode\" = \"OWHS\".\"WhsCode\"" +
                           " WHERE \"OIGE\".\"DocEntry\" = '" + baseDocEntry + "'";
            try
            {
                oRecordSet.DoQuery(queryEndAddrs);
                while (!oRecordSet.EoF)
                {
                    endAddrs = endAddrs + oRecordSet.Fields.Item("Street").Value.ToString() + '\n' +
                        oRecordSet.Fields.Item("ZipCode").Value.ToString() + " " +
                        oRecordSet.Fields.Item("City").Value.ToString() + '\n' +
                        oRecordSet.Fields.Item("County").Value.ToString();
                    oRecordSet.MoveNext();
                    break;
                }
            }

            catch (Exception ex)
            {
                errorText = ex.Message;
                return;
            }

            try
            {
                oRecordSet.DoQuery(query);

                while (!oRecordSet.EoF)
                {
                    CompanyService oCompanyService = null;
                    GeneralService oGeneralService = null;
                    GeneralData oGeneralData = null;
                    oCompanyService = oCompany.GetCompanyService();
                    oGeneralService = oCompanyService.GetGeneralService("UDO_F_BDO_WBLD_D");
                    oGeneralData = ((GeneralData)(oGeneralService.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralData)));

                    string empID;
                    string empName;
                    Users.getUserEmployee(out empID, out empName, out errorText);

                    oGeneralData.SetProperty("U_endAddrs", endAddrs);
                    oGeneralData.SetProperty("U_delvDate", oRecordSet.Fields.Item("DocDate").Value.ToString());
                    oGeneralData.SetProperty("U_recpInfo", empID == "0" ? "" : empID); //ჩამბარებელი
                    oGeneralData.SetProperty("U_recpInfN", empName);
                    oGeneralData.SetProperty("U_recvInfo", ""); //მიმღები
                    oGeneralData.SetProperty("U_recvInfN", "");
                    oGeneralData.SetProperty("U_cardCode", "");
                    oGeneralData.SetProperty("U_baseDoc", oRecordSet.Fields.Item("DocEntry").Value);
                    oGeneralData.SetProperty("U_baseDTxt", oRecordSet.Fields.Item("DocEntry").Value.ToString());
                    oGeneralData.SetProperty("U_baseDocT", oRecordSet.Fields.Item("ObjType").Value.ToString());
                    oGeneralData.SetProperty("U_strAddrs", strAddrs);
                    oGeneralData.SetProperty("U_begDate", DateTime.Today);  //DateTime.Today.ToString("yyyyMMdd"));
                    oGeneralData.SetProperty("U_beginTime", DateTime.Now);

                    if (vehicleCode != null)
                    {
                        if (trnsType == "4")
                        {
                            oGeneralData.SetProperty("U_vehicNum", vehicleCode);
                        }
                        else
                        {
                            //-->სატრანსპორტო საშუალება
                            UserTable oUserTable = null;
                            oUserTable = oCompany.UserTables.Item("BDO_VECL");
                            oUserTable.GetByKey(vehicleCode);
                            oGeneralData.SetProperty("U_vehicle", vehicleCode);
                            string vehicleNumber = oUserTable.UserFields.Fields.Item("U_number").Value;
                            string vehicleTrailNum = oUserTable.UserFields.Fields.Item("U_trailNum").Value;
                            oGeneralData.SetProperty("U_vehicNum", vehicleNumber);
                            oGeneralData.SetProperty("U_trailNum", vehicleTrailNum);
                            //სატრანსპორტო საშუალება<--

                            //-->მძღოლი                     
                            if (driverCode == null || driverCode == "")
                            {
                                driverCode = oUserTable.UserFields.Fields.Item("U_drvCode").Value;
                            }
                            oUserTable = oCompany.UserTables.Item("BDO_DRVS");
                            oUserTable.GetByKey(driverCode);
                            string driverTin = oUserTable.UserFields.Fields.Item("U_tin").Value;
                            string driverNotResident = oUserTable.UserFields.Fields.Item("U_notRsdnt").Value;
                            oGeneralData.SetProperty("U_drvCode", driverCode);
                            oGeneralData.SetProperty("U_notRsdnt", driverNotResident);
                            oGeneralData.SetProperty("U_drivTin", driverTin);
                            //მძღოლი<--
                        }
                    }

                    if (trnsprter != null)
                    {
                        Recordset oRecordSetBP = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                        string queryBP = "SELECT " + "\"OCRD\".\"CardName\" AS \"CardName\", \"OCRD\".\"LicTradNum\" AS \"LicTradNum\" " + " FROM \"OCRD\"" + " WHERE " + "\"CardCode\"" + " = N'" + trnsprter + "'";
                        oRecordSetBP.DoQuery(queryBP);

                        if (!oRecordSetBP.EoF)
                        {
                            oGeneralData.SetProperty("U_tporter", trnsprter);
                            oGeneralData.SetProperty("U_tporterN", oRecordSetBP.Fields.Item("CardName").Value);
                            oGeneralData.SetProperty("U_tporterT", oRecordSetBP.Fields.Item("LicTradNum").Value);
                        }
                    }

                    trnsType = trnsType == null ? "1" : trnsType;
                    oGeneralData.SetProperty("U_type", "0"); //ტრანსპორტირებით
                    oGeneralData.SetProperty("U_status", "-1");
                    oGeneralData.SetProperty("U_payForTr", "-1");
                    oGeneralData.SetProperty("U_trnsType", trnsType); //საავტომობილო

                    try
                    {
                        var response = oGeneralService.Add(oGeneralData);
                        var docEntry = response.GetProperty("DocEntry");
                        newDocEntry = Convert.ToInt32(docEntry);
                    }
                    catch (Exception ex)
                    {
                        int errCode;
                        string errMsg;

                        oCompany.GetLastError(out errCode, out errMsg);
                        errorText = ex.Message;
                        errorText = BDOSResources.getTranslate("ErrorOfDocumentAdd") + " " + BDOSResources.getTranslate("ErrorDescription") + " " + errMsg + "! " + BDOSResources.getTranslate("Code") + " : " + errCode + "! " + errorText;
                    }

                    oRecordSet.MoveNext();
                    break;
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
                return;
            }
            finally
            {
                GC.Collect();
            }
        }

        private static void CreateDocumentArCorrectionInvoiceType(int baseDocEntry, string vehicleCode, string driverCode, string trnsType, string trnsprter, out int newDocEntry, out string errorText)
        {
            errorText = null;
            newDocEntry = 0;

            var oRecordSet = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            var query =
                "SELECT \"OCSI\".\"DocEntry\", \"OCSI\".\"ObjType\", \"OCSI\".\"CardCode\", \"OCSI\".\"Address2\", \"OCSI\".\"DocDate\", \"OCSI\".\"CntctCode\" , \"OCPR\".\"Name\"" +
                " FROM \"OCSI\" AS \"OCSI\"" +
                " LEFT JOIN \"OCPR\" AS \"OCPR\"" +
                " ON \"OCSI\".\"CntctCode\" = \"OCPR\".\"CntctCode\"" +
                " WHERE \"DocEntry\" = '" + baseDocEntry + "'";

            var strAddrs = "";
            var queryWhs =
                "SELECT DISTINCT \"CSI1\".\"WhsCode\", \"OWHS\".\"Street\", \"OWHS\".\"ZipCode\", \"OWHS\".\"City\", \"OWHS\".\"County\"" +
                " FROM \"CSI1\" AS \"CSI1\"" +
                " INNER JOIN \"OWHS\" AS \"OWHS\"" +
                " ON \"CSI1\".\"WhsCode\" = \"OWHS\".\"WhsCode\"" +
                " WHERE \"CSI1\".\"DocEntry\" = '" + baseDocEntry + "'";
            try
            {
                oRecordSet.DoQuery(queryWhs);
                if (oRecordSet.RecordCount == 1)
                    while (!oRecordSet.EoF)
                    {
                        strAddrs = strAddrs + oRecordSet.Fields.Item("Street").Value.ToString() + '\n' +
                                   oRecordSet.Fields.Item("ZipCode").Value.ToString() + " " +
                                   oRecordSet.Fields.Item("City").Value.ToString() + '\n' +
                                   oRecordSet.Fields.Item("County").Value.ToString();
                        oRecordSet.MoveNext();
                        break;
                    }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
                return;
            }

            try
            {
                oRecordSet.DoQuery(query);

                while (!oRecordSet.EoF)
                {
                    CompanyService oCompanyService = null;
                    GeneralService oGeneralService = null;
                    GeneralData oGeneralData = null;
                    oCompanyService = oCompany.GetCompanyService();
                    oGeneralService = oCompanyService.GetGeneralService("UDO_F_BDO_WBLD_D");
                    oGeneralData =
                        (GeneralData)oGeneralService.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralData);

                    Users.getUserEmployee(out var empId, out var empName, out errorText);

                    string cntctCode = oRecordSet.Fields.Item("CntctCode").Value.ToString();

                    oGeneralData.SetProperty("U_strAddrs",
                        oRecordSet.Fields.Item("Address2").Value.ToString()); //U_endAddrs
                    oGeneralData.SetProperty("U_delvDate", oRecordSet.Fields.Item("DocDate").Value.ToString());
                    oGeneralData.SetProperty("U_recpInfo", empId == "0" ? "" : empId); //ჩამბარებელი
                    oGeneralData.SetProperty("U_recpInfN", empName);
                    oGeneralData.SetProperty("U_recvInfo", cntctCode == "0" ? "" : cntctCode); //მიმღები
                    oGeneralData.SetProperty("U_recvInfN", oRecordSet.Fields.Item("Name").Value.ToString());
                    oGeneralData.SetProperty("U_cardCode", oRecordSet.Fields.Item("CardCode").Value.ToString());
                    oGeneralData.SetProperty("U_baseDoc", oRecordSet.Fields.Item("DocEntry").Value);
                    oGeneralData.SetProperty("U_baseDTxt", oRecordSet.Fields.Item("DocEntry").Value.ToString());
                    oGeneralData.SetProperty("U_baseDocT", oRecordSet.Fields.Item("ObjType").Value.ToString());
                    oGeneralData.SetProperty("U_endAddrs", strAddrs); //U_strAddrs
                    oGeneralData.SetProperty("U_begDate", DateTime.Today); //DateTime.Today.ToString("yyyyMMdd"));
                    oGeneralData.SetProperty("U_beginTime", DateTime.Now);

                    if (vehicleCode != null)
                    {
                        if (trnsType == "4")
                        {
                            oGeneralData.SetProperty("U_vehicNum", vehicleCode);
                        }
                        else
                        {
                            //-->სატრანსპორტო საშუალება
                            var oUserTable = oCompany.UserTables.Item("BDO_VECL");
                            oUserTable.GetByKey(vehicleCode);
                            oGeneralData.SetProperty("U_vehicle", vehicleCode);
                            string vehicleNumber = oUserTable.UserFields.Fields.Item("U_number").Value;
                            string vehicleTrailNum = oUserTable.UserFields.Fields.Item("U_trailNum").Value;
                            oGeneralData.SetProperty("U_vehicNum", vehicleNumber);
                            oGeneralData.SetProperty("U_trailNum", vehicleTrailNum);
                            //სატრანსპორტო საშუალება<--

                            //-->მძღოლი                     
                            if (string.IsNullOrEmpty(driverCode))
                                driverCode = oUserTable.UserFields.Fields.Item("U_drvCode").Value;
                            oUserTable = oCompany.UserTables.Item("BDO_DRVS");
                            oUserTable.GetByKey(driverCode);
                            string driverTin = oUserTable.UserFields.Fields.Item("U_tin").Value;
                            string driverNotResident = oUserTable.UserFields.Fields.Item("U_notRsdnt").Value;
                            oGeneralData.SetProperty("U_drvCode", driverCode);
                            oGeneralData.SetProperty("U_notRsdnt", driverNotResident);
                            oGeneralData.SetProperty("U_drivTin", driverTin);
                            //მძღოლი<--
                        }
                    }

                    if (trnsprter != null)
                    {
                        var oRecordSetBP = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                        var queryBP =
                            "SELECT " +
                            "\"OCRD\".\"CardName\" AS \"CardName\", \"OCRD\".\"LicTradNum\" AS \"LicTradNum\" " +
                            " FROM \"OCRD\"" + " WHERE " + "\"CardCode\"" + " = N'" + trnsprter + "'";
                        oRecordSetBP.DoQuery(queryBP);

                        if (!oRecordSetBP.EoF)
                        {
                            oGeneralData.SetProperty("U_tporter", trnsprter);
                            oGeneralData.SetProperty("U_tporterN", oRecordSetBP.Fields.Item("CardName").Value);
                            oGeneralData.SetProperty("U_tporterT", oRecordSetBP.Fields.Item("LicTradNum").Value);
                        }
                    }

                    trnsType = trnsType ?? "1";
                    oGeneralData.SetProperty("U_type", "0"); //ტრანსპორტირებით
                    oGeneralData.SetProperty("U_status", "-1");
                    oGeneralData.SetProperty("U_payForTr", "-1");
                    oGeneralData.SetProperty("U_trnsType", trnsType); //საავტომობილო

                    try
                    {
                        var response = oGeneralService.Add(oGeneralData);
                        var docEntry = response.GetProperty("DocEntry");
                        newDocEntry = Convert.ToInt32(docEntry);
                    }
                    catch (Exception ex)
                    {
                        oCompany.GetLastError(out var errCode, out var errMsg);
                        errorText = ex.Message;
                        errorText = BDOSResources.getTranslate("ErrorOfDocumentAdd") + " " +
                                    BDOSResources.getTranslate("ErrorDescription") + " " + errMsg + "! " +
                                    BDOSResources.getTranslate("Code") + " : " + errCode + "! " + errorText;
                    }

                    oRecordSet.MoveNext();
                    break;
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
        public static void setSizeForm(SAPbouiCOM.Form oForm)
        {
            oForm.Freeze(true);
            try
            {
                oForm.ClientHeight = uiApp.Desktop.Height / 2;
                //oForm.Height = Program.uiApp.Desktop.Width / 4;
                oForm.Left = (uiApp.Desktop.Width - oForm.Width) / 2;
                oForm.Top = (uiApp.Desktop.Height - oForm.Height) / 3;
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
            try
            {
                oForm.Freeze(true);

                SAPbouiCOM.Item oItem = null;
                int height = 15;

                int top = 6;

                oItem = oForm.Items.Item("0_U_E");
                oItem.Left = 160;
                oItem.Width = 140;

                top = top + height + 1;

                oItem = oForm.Items.Item("1_U_S");
                oItem.Top = top;
                oItem = oForm.Items.Item("1_U_E");
                oItem.Top = top;
                top = top + height + 1;

                oItem = oForm.Items.Item("beginTimeS");
                oItem.Top = top;
                oItem = oForm.Items.Item("beginTimeE");
                oItem.Top = top;
                top = top + height + 1;

                oItem = oForm.Items.Item("2_U_S");
                oItem.Top = top;
                oItem = oForm.Items.Item("2_U_E");
                oItem.Top = top;
                top = top + height + 1;

                oItem = oForm.Items.Item("3_U_S");
                oItem.Top = top;
                oItem = oForm.Items.Item("3_U_E");
                oItem.Top = top;
                top = top + height + 1;

                oItem = oForm.Items.Item("4_U_S");
                oItem.Top = top;
                oItem = oForm.Items.Item("4_U_E");
                oItem.Top = top;
                top = top + height + 1;

                oItem = oForm.Items.Item("5_U_S");
                oItem.Top = top;
                oItem = oForm.Items.Item("5_U_C");
                oItem.Top = top;
                top = top + height + 1;

                oItem = oForm.Items.Item("6_U_S");
                oItem.Top = top;
                oItem = oForm.Items.Item("6_U_C");
                oItem.Top = top;

                top = 6;

                oItem = oForm.Items.Item("7_U_S");
                oItem.Top = top;
                oItem = oForm.Items.Item("8_U_C");
                oItem.Top = top;
                oItem = oForm.Items.Item("9_U_E");
                oItem.Top = top;
                top = top + height + 1;

                oItem = oForm.Items.Item("10_U_S");
                oItem.Top = top;
                oItem = oForm.Items.Item("10_U_C");
                oItem.Top = top;
                top = top + height + 1;

                oItem = oForm.Items.Item("11_U_S");
                oItem.Top = top;
                oItem = oForm.Items.Item("11_U_E");
                oItem.Top = top;
                top = top + height + 1;

                oItem = oForm.Items.Item("12_U_S");
                oItem.Top = top;
                oItem = oForm.Items.Item("12_U_E");
                oItem.Top = top;
                top = top + height + 1;

                //oItem = oForm.Items.Item("13_U_E");
                //oItem.Top = top;
                oItem = oForm.Items.Item("13_U_S");
                oItem.Top = top;
                oItem = oForm.Items.Item("14_U_E");
                oItem.Top = top;
                oItem = oForm.Items.Item("14_U_LB");
                oItem.Top = top;

                //Address section ---->
                top = 150;

                oItem = oForm.Items.Item("15_U_S");
                oItem.Top = top;
                top = top + 25;

                oItem = oForm.Items.Item("16_U_S");
                oItem.Top = top;
                oItem = oForm.Items.Item("16_U_E");
                oItem.Top = top;
                top = top + height + 1;

                oItem = oForm.Items.Item("34_U_E");
                oItem.Top = top;
                oItem = oForm.Items.Item("17_U_S");
                oItem.Top = top;
                oItem = oForm.Items.Item("17_U_E");
                oItem.Top = top;
                oItem = oForm.Items.Item("17_U_E1");
                oItem.Top = top;
                oItem = oForm.Items.Item("17_U_LB");
                oItem.Top = top;
                top = top + height + 1;

                top = 150;
                top = top + 25;

                oItem = oForm.Items.Item("19_U_S");
                oItem.Top = top;
                oItem = oForm.Items.Item("19_U_E");
                oItem.Top = top;
                top = top + height + 1;

                oItem = oForm.Items.Item("20_U_S");
                oItem.Top = top;
                oItem = oForm.Items.Item("20_U_E");
                oItem.Top = top;
                oItem = oForm.Items.Item("20_U_E1");
                oItem.Top = top;
                top = top + height + 1;

                oItem = oForm.Items.Item("21_U_S");
                oItem.Top = top;
                oItem = oForm.Items.Item("21_U_E");
                oItem.Top = top;

                //Transport section ---->
                top = 255;

                oItem = oForm.Items.Item("22_U_S");
                oItem.Top = top;
                top = top + 25;

                oItem = oForm.Items.Item("23_U_S");
                oItem.Top = top;
                oItem = oForm.Items.Item("23_U_C");
                oItem.Top = top;
                top = top + height + 1;

                oItem = oForm.Items.Item("24_U_S");
                oItem.Top = top;
                oItem = oForm.Items.Item("24_U_E");
                oItem.Top = top;
                oItem = oForm.Items.Item("24_U_LB");
                oItem.Top = top;
                oItem = oForm.Items.Item("24_U_E1");
                oItem.Top = top;
                top = top + height + 1;

                oItem = oForm.Items.Item("25_U_S");
                oItem.Top = top;
                oItem = oForm.Items.Item("25_U_E");
                oItem.Top = top;
                top = 255 + 25;

                oItem = oForm.Items.Item("26_U_S");
                oItem.Top = top;
                oItem = oForm.Items.Item("26_U_E");
                oItem.Top = top;
                oItem = oForm.Items.Item("26_U_LB");
                oItem.Top = top;
                oItem = oForm.Items.Item("26_U_B");
                oItem.Top = top - 2;
                top = top + height + 1;

                oItem = oForm.Items.Item("27_U_S");
                oItem.Top = top;
                oItem = oForm.Items.Item("27_U_E");
                oItem.Top = top;
                top = top + height + 1;

                oItem = oForm.Items.Item("28_U_CH");
                oItem.Top = top;
                top = top + height + 1;

                oItem = oForm.Items.Item("29_U_S");
                oItem.Top = top;
                oItem = oForm.Items.Item("29_U_E");
                oItem.Top = top;
                oItem = oForm.Items.Item("29_U_C");
                oItem.Top = top;
                top = top + height + 1;

                oItem = oForm.Items.Item("30_U_S");
                oItem.Top = top;
                oItem = oForm.Items.Item("30_U_E");
                oItem.Top = top;
                oItem = oForm.Items.Item("30_U_E1");
                oItem.Top = top;
                oItem = oForm.Items.Item("30_U_LB");
                oItem.Top = top;
                top = top + height + 1;

                oItem = oForm.Items.Item("31_U_S");
                oItem.Top = top;
                oItem = oForm.Items.Item("31_U_E");
                oItem.Top = top;

                //სარდაფი
                top = top + 25;

                oItem = oForm.Items.Item("32_U_S");
                oItem.Top = top;
                oItem = oForm.Items.Item("32_U_E");
                oItem.Top = top;
                top = top + height + 1;

                oItem = oForm.Items.Item("33_U_S");
                oItem.Top = top;
                oItem = oForm.Items.Item("33_U_E");
                oItem.Top = top;
                top = top + 3 * height + 1;

                oItem = oForm.Items.Item("18_U_S");
                oItem.Top = top;
                oItem = oForm.Items.Item("18_U_E");
                oItem.Top = top;

                //ღილაკები
                oItem = oForm.Items.Item("1");
                oItem.Top = oForm.ClientHeight - 25;

                oItem = oForm.Items.Item("2");
                oItem.Top = oForm.ClientHeight - 25;

                oItem = oForm.Items.Item("33_U_BC");
                oItem.Left = oForm.ClientWidth - 6 - oItem.Width;
                oItem.Top = oForm.ClientHeight - 25;
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

        public static void chooseFromList(SAPbouiCOM.Form oForm, SAPbouiCOM.IChooseFromListEvent oCFLEvento, string itemUID, bool beforeAction)
        {
            try
            {
                string sCFL_ID = oCFLEvento.ChooseFromListUID;
                SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(sCFL_ID);

                if (beforeAction)
                {
                    if (sCFL_ID == "Contact_CFL")
                    {
                        SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();

                        SAPbouiCOM.Condition oCon = oCons.Add();
                        oCon.Alias = "CardCode";
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = oForm.DataSources.DBDataSources.Item("@BDO_WBLD").GetValue("U_cardCode", 0).Trim();
                        oCFL.SetConditions(oCons);
                    }
                    else if (sCFL_ID == "BaseDoc_CFL" + oForm.DataSources.DBDataSources.Item("@BDO_WBLD").GetValue("U_baseDocT", 0).Trim())
                    {
                        SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();

                        string cardCode = oForm.DataSources.DBDataSources.Item("@BDO_WBLD").GetValue("U_cardCode", 0).Trim();
                        string tableName = "OINV";
                        switch (oCFL.ObjectType)
                        {
                            case "13":
                                tableName = "OINV";
                                break;
                            case "15":
                                tableName = "ODLN";
                                break;
                            case "165":
                                tableName = "OCSI";
                                break;
                            case "67":
                                tableName = "OWTR";
                                break;
                            case "UDO_F_BDOSFASTRD_D":
                                tableName = "@BDOSFASTRD";
                                break;
                            case "14":
                                tableName = "ORIN";
                                break;
                            case "60":
                                tableName = "OIGE";
                                break;
                        }

                        Recordset oRecordSet = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                        string query = "SELECT \"DOC\".\"DocEntry\" " +
                        "FROM \"" + tableName + "\" AS \"DOC\" " +
                        "WHERE \"DOC\".\"CardCode\" = N'" + cardCode + "' AND \"DOC\".\"DocStatus\" = 'O' AND \"DOC\".\"DocEntry\" " +
                        "NOT IN (SELECT \"BDO_WBLD\".\"U_baseDoc\" " +
                        "FROM \"@BDO_WBLD\" AS \"BDO_WBLD\" " +
                        "WHERE \"BDO_WBLD\".\"U_baseDocT\" = '" + oCFL.ObjectType + "' AND \"BDO_WBLD\".\"U_baseDoc\" <> '') " +
                        "GROUP BY \"DOC\".\"DocEntry\"";

                        try
                        {
                            oRecordSet.DoQuery(query);
                            int recordCount = oRecordSet.RecordCount;
                            int i = 1;

                            while (!oRecordSet.EoF)
                            {
                                SAPbouiCOM.Condition oCon = oCons.Add();
                                oCon.Alias = "DocEntry";
                                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                                oCon.CondVal = oRecordSet.Fields.Item("DocEntry").Value.ToString();
                                oCon.Relationship = (i == recordCount) ? SAPbouiCOM.BoConditionRelationship.cr_NONE : SAPbouiCOM.BoConditionRelationship.cr_OR;

                                i = i + 1;
                                oRecordSet.MoveNext();
                            }
                            oCFL.SetConditions(oCons);
                        }
                        catch (Exception ex)
                        {
                            throw new Exception(ex.Message);
                        }
                    }
                }
                else
                {
                    SAPbouiCOM.DataTable oDataTable = null;
                    oDataTable = oCFLEvento.SelectedObjects;

                    if (oDataTable != null)
                    {
                        if (sCFL_ID == "VehicleCode_CFL")
                        {
                            string trnsType = oForm.DataSources.DBDataSources.Item("@BDO_WBLD").GetValue("U_trnsType", 0).Trim();

                            string vehicleCode = Convert.ToString(oDataTable.GetValue("Code", 0));
                            string vehicleNumber = Convert.ToString(oDataTable.GetValue("U_number", 0));
                            string vehicleTrailNum = Convert.ToString(oDataTable.GetValue("U_trailNum", 0));

                            UserTable oUserTable = null;
                            oUserTable = oCompany.UserTables.Item("BDO_VECL");
                            oUserTable.GetByKey(vehicleCode);
                            string driverCode = oUserTable.UserFields.Fields.Item("U_drvCode").Value;

                            oUserTable = oCompany.UserTables.Item("BDO_DRVS");
                            oUserTable.GetByKey(driverCode);
                            string driverTin = oUserTable.UserFields.Fields.Item("U_tin").Value;
                            string driverNotResident = oUserTable.UserFields.Fields.Item("U_notRsdnt").Value;

                            oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_vehicle", 0, vehicleCode);
                            oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_vehicNum", 0, vehicleNumber);

                            if (trnsType != "4") //სხვა
                            {
                                oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_trailNum", 0, vehicleTrailNum);

                                oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_drvCode", 0, driverCode);
                                oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_drivTin", 0, driverTin);
                                oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_notRsdnt", 0, driverNotResident);
                            }

                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            {
                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                            }
                        }
                        else if (sCFL_ID == "DriverCode_CFL")
                        {
                            string driverCode = Convert.ToString(oDataTable.GetValue("Code", 0));
                            string driverTin = Convert.ToString(oDataTable.GetValue("U_tin", 0));
                            string driverNotResident = Convert.ToString(oDataTable.GetValue("U_notRsdnt", 0));

                            oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_drvCode", 0, driverCode);
                            oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_drivTin", 0, driverTin);
                            oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_notRsdnt", 0, driverNotResident);

                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            {
                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                            }
                        }
                        else if (sCFL_ID == "BusinessPartner_CFL")
                        {
                            string businessPartnerCode = Convert.ToString(oDataTable.GetValue("CardCode", 0));
                            string businessPartnerName = Convert.ToString(oDataTable.GetValue("CardName", 0));
                            string businessPartnerLicTradNum = Convert.ToString(oDataTable.GetValue("LicTradNum", 0));
                            oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_tporter", 0, businessPartnerCode);
                            oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_tporterN", 0, businessPartnerName);
                            oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_tporterT", 0, businessPartnerLicTradNum);

                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            {
                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                            }
                        }
                        else if (sCFL_ID == "Employee_CFL")
                        {
                            string employeeID = Convert.ToString(oDataTable.GetValue("empID", 0));
                            string employeeFirstName = Convert.ToString(oDataTable.GetValue("firstName", 0));
                            string employeeLastName = Convert.ToString(oDataTable.GetValue("lastName", 0));
                            if (itemUID == "20_U_E")
                            {
                                oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_recvInfo", 0, employeeID);
                                oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_recvInfN", 0, employeeFirstName + " " + employeeLastName);
                            }
                            else
                            {
                                oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_recpInfo", 0, employeeID);
                                oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_recpInfN", 0, employeeFirstName + " " + employeeLastName);
                            }
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            {
                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                            }
                        }
                        else if (sCFL_ID == "Contact_CFL")
                        {
                            string contactID = Convert.ToString(oDataTable.GetValue("CntctCode", 0));
                            string contactName = Convert.ToString(oDataTable.GetValue("Name", 0));
                            oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_recvInfo", 0, contactID);
                            oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_recvInfN", 0, contactName);

                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            {
                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                            }
                        }
                        else if (sCFL_ID == "BaseDoc_CFL" + oForm.DataSources.DBDataSources.Item("@BDO_WBLD").GetValue("U_baseDocT", 0).Trim())
                        {
                            int docCode = oDataTable.GetValue("DocEntry", 0);
                            oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_baseDoc", 0, docCode.ToString());
                            oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_baseDTxt", 0, docCode.ToString());
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            {
                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                            }
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
                GC.Collect();
            }
        }

        public static void setVisibleFormItems(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.Item oItem = null;

            oForm.Freeze(true);
            try
            {
                string baseDocSt = oForm.DataSources.DBDataSources.Item("@BDO_WBLD").GetValue("U_baseDoc", 0);
                int baseDoc = string.IsNullOrEmpty(baseDocSt) ? 0 : Convert.ToInt32(baseDocSt);
                string baseDocType = oForm.DataSources.DBDataSources.Item("@BDO_WBLD").GetValue("U_baseDocT", 0).Trim();

                oItem = oForm.Items.Item("14_U_E");
                oItem.Enabled = baseDoc == 0 ? true : false;

                if (!string.IsNullOrEmpty(baseDocType))
                    oForm.Items.Item("6_U_C").Enabled = baseDocType != "13" && baseDocType != "15" ? false : true; // != A/R Invoice

                oItem = oForm.Items.Item("32_U_E");
                oItem.Enabled = false;

                string waybillType = oForm.DataSources.DBDataSources.Item("@BDO_WBLD").GetValue("U_type", 0).Trim(); //--->
                oItem = oForm.Items.Item("22_U_S");
                oItem.Visible = (waybillType == "1") ? false : true;
                oItem = oForm.Items.Item("23_U_S");
                oItem.Visible = (waybillType == "1") ? false : true;
                oItem = oForm.Items.Item("23_U_C");
                oItem.Visible = (waybillType == "1") ? false : true;
                oItem = oForm.Items.Item("24_U_S");
                oItem.Visible = (waybillType == "1") ? false : true;
                oItem = oForm.Items.Item("24_U_E");
                oItem.Visible = (waybillType == "1") ? false : true;
                oItem = oForm.Items.Item("24_U_LB");
                oItem.Visible = (waybillType == "1") ? false : true;
                oItem = oForm.Items.Item("24_U_E1");
                oItem.Visible = (waybillType == "1") ? false : true;
                oItem = oForm.Items.Item("25_U_S");
                oItem.Visible = (waybillType == "1") ? false : true;
                oItem = oForm.Items.Item("25_U_E");
                oItem.Visible = (waybillType == "1") ? false : true;
                oItem = oForm.Items.Item("26_U_S");
                oItem.Visible = (waybillType == "1") ? false : true;
                oItem = oForm.Items.Item("26_U_E");
                oItem.Visible = (waybillType == "1") ? false : true;
                oItem = oForm.Items.Item("26_U_LB");
                oItem.Visible = (waybillType == "1") ? false : true;
                oItem = oForm.Items.Item("26_U_B");
                oItem.Visible = (waybillType == "1") ? false : true;
                oItem = oForm.Items.Item("27_U_S");
                oItem.Visible = (waybillType == "1") ? false : true;
                oItem = oForm.Items.Item("27_U_E");
                oItem.Visible = (waybillType == "1") ? false : true;
                oItem = oForm.Items.Item("28_U_CH");
                oItem.Visible = (waybillType == "1") ? false : true;
                oItem = oForm.Items.Item("29_U_S");
                oItem.Visible = (waybillType == "1") ? false : true;
                oItem = oForm.Items.Item("29_U_E");
                oItem.Visible = (waybillType == "1") ? false : true;
                oItem = oForm.Items.Item("29_U_C");
                oItem.Visible = (waybillType == "1") ? false : true;
                oItem = oForm.Items.Item("30_U_S");
                oItem.Visible = (waybillType == "1") ? false : true;
                oItem = oForm.Items.Item("30_U_E");
                oItem.Visible = (waybillType == "1") ? false : true;
                oItem = oForm.Items.Item("30_U_E1");
                oItem.Visible = (waybillType == "1") ? false : true;
                oItem = oForm.Items.Item("30_U_LB");
                oItem.Visible = (waybillType == "1") ? false : true;
                oItem = oForm.Items.Item("31_U_S");
                oItem.Visible = (waybillType == "1") ? false : true;
                oItem = oForm.Items.Item("31_U_E");
                oItem.Visible = (waybillType == "1") ? false : true;
                //<---

                string trnsType = oForm.DataSources.DBDataSources.Item("@BDO_WBLD").GetValue("U_trnsType", 0).Trim(); //--->
                if (waybillType == "0")
                {
                    if (trnsType == "6") //გადამზიდავი საავტომობილო
                    {
                        oItem = oForm.Items.Item("22_U_S");
                        oItem.Visible = true;

                        oItem = oForm.Items.Item("23_U_S");
                        oItem.Visible = true;
                        oItem = oForm.Items.Item("23_U_C");
                        oItem.Visible = true;

                        oItem = oForm.Items.Item("24_U_S");
                        oItem.Visible = true;
                        oItem = oForm.Items.Item("24_U_E");
                        oItem.Enabled = false;
                        oItem = oForm.Items.Item("24_U_LB");
                        oItem.Visible = true;
                        oItem = oForm.Items.Item("24_U_E1");
                        oItem.Enabled = false;

                        oItem = oForm.Items.Item("25_U_S");
                        oItem.Visible = true;
                        oItem = oForm.Items.Item("25_U_E");
                        oItem.Enabled = false;

                        oItem = oForm.Items.Item("26_U_S");
                        oItem.Visible = true;
                        oItem = oForm.Items.Item("26_U_E");
                        oItem.Enabled = false;
                        oItem = oForm.Items.Item("26_U_LB");
                        oItem.Visible = true;
                        oItem = oForm.Items.Item("26_U_B");
                        oItem.Enabled = false;

                        oItem = oForm.Items.Item("27_U_S");
                        oItem.Visible = true;
                        oItem = oForm.Items.Item("27_U_E");
                        oItem.Visible = true;
                        oItem.Enabled = false;

                        oItem = oForm.Items.Item("28_U_CH");
                        oItem.Visible = true;
                        oItem.Enabled = false;

                        //oItem = oForm.Items.Item("29_U_S");
                        //oItem.Enabled = true;
                        //oItem = oForm.Items.Item("29_U_E");
                        //oItem.Enabled = true;
                        //oItem = oForm.Items.Item("29_U_C");
                        //oItem.Enabled = true;

                        oItem = oForm.Items.Item("30_U_S");
                        oItem.Visible = true;
                        oItem = oForm.Items.Item("30_U_E");
                        oItem.Visible = true;
                        oItem = oForm.Items.Item("30_U_E1");
                        oItem.Visible = true;
                        oItem = oForm.Items.Item("30_U_LB");
                        oItem.Visible = true;

                        oItem = oForm.Items.Item("31_U_S");
                        oItem.Visible = true;
                        oItem = oForm.Items.Item("31_U_E");
                        oItem.Visible = true;
                    }
                    else if (trnsType == "1" || trnsType == "5") //საავტომობილო || საავტომობილო უცხო ქვეყნის
                    {
                        oItem = oForm.Items.Item("22_U_S");
                        oItem.Visible = true;

                        oItem = oForm.Items.Item("23_U_S");
                        oItem.Visible = true;
                        oItem = oForm.Items.Item("23_U_C");
                        oItem.Visible = true;

                        oItem = oForm.Items.Item("24_U_S");
                        oItem.Visible = true;
                        oItem = oForm.Items.Item("24_U_E");
                        oItem.Visible = true;
                        oItem.Enabled = true;
                        oItem = oForm.Items.Item("24_U_LB");
                        oItem.Visible = true;
                        oItem = oForm.Items.Item("24_U_E1");
                        oItem.Visible = true;
                        oItem.Enabled = true;

                        oItem = oForm.Items.Item("25_U_S");
                        oItem.Visible = true;
                        oItem = oForm.Items.Item("25_U_E");
                        oItem.Visible = true;
                        oItem.Enabled = true;

                        oItem = oForm.Items.Item("26_U_S");
                        oItem.Visible = true;
                        oItem = oForm.Items.Item("26_U_E");
                        oItem.Visible = true;
                        oItem.Enabled = true;
                        oItem = oForm.Items.Item("26_U_LB");
                        oItem.Visible = true;
                        oItem = oForm.Items.Item("26_U_B");
                        oItem.Visible = true;
                        oItem.Enabled = true;

                        oItem = oForm.Items.Item("27_U_S");
                        oItem.Visible = true;
                        oItem = oForm.Items.Item("27_U_E");
                        oItem.Visible = true;
                        oItem.Enabled = true;

                        oItem = oForm.Items.Item("28_U_CH");
                        oItem.Visible = true;
                        oItem.Enabled = true;

                        //oItem = oForm.Items.Item("29_U_S");
                        //oItem.Enabled = false;
                        //oItem = oForm.Items.Item("29_U_E");
                        //oItem.Enabled = false;
                        //oItem = oForm.Items.Item("29_U_C");
                        //oItem.Enabled = false;

                        oItem = oForm.Items.Item("30_U_S");
                        oItem.Visible = false;
                        oItem = oForm.Items.Item("30_U_E");
                        oItem.Visible = false;
                        oItem = oForm.Items.Item("30_U_E1");
                        oItem.Visible = false;
                        oItem = oForm.Items.Item("30_U_LB");
                        oItem.Visible = false;

                        oItem = oForm.Items.Item("31_U_S");
                        oItem.Visible = false;
                        oItem = oForm.Items.Item("31_U_E");
                        oItem.Visible = false;
                    }
                    else if (trnsType == "2" || trnsType == "3") //სარკინიგზო || საავიაციო
                    {
                        oItem = oForm.Items.Item("22_U_S");
                        oItem.Visible = true;

                        oItem = oForm.Items.Item("23_U_S");
                        oItem.Visible = true;
                        oItem = oForm.Items.Item("23_U_C");
                        oItem.Visible = true;

                        oItem = oForm.Items.Item("24_U_S");
                        oItem.Visible = false;
                        oItem = oForm.Items.Item("24_U_E");
                        oItem.Visible = false;
                        oItem = oForm.Items.Item("24_U_LB");
                        oItem.Visible = false;
                        oItem = oForm.Items.Item("24_U_E1");
                        oItem.Visible = false;

                        oItem = oForm.Items.Item("25_U_S");
                        oItem.Visible = false;
                        oItem = oForm.Items.Item("25_U_E");
                        oItem.Visible = false;

                        oItem = oForm.Items.Item("26_U_S");
                        oItem.Visible = false;
                        oItem = oForm.Items.Item("26_U_E");
                        oItem.Visible = false;
                        oItem = oForm.Items.Item("26_U_LB");
                        oItem.Visible = false;
                        oItem = oForm.Items.Item("26_U_B");
                        oItem.Visible = false;

                        oItem = oForm.Items.Item("27_U_S");
                        oItem.Visible = false;
                        oItem = oForm.Items.Item("27_U_E");
                        oItem.Visible = false;

                        oItem = oForm.Items.Item("28_U_CH");
                        oItem.Visible = false;

                        oItem = oForm.Items.Item("29_U_S");
                        oItem.Visible = false;
                        oItem = oForm.Items.Item("29_U_E");
                        oItem.Visible = false;
                        oItem = oForm.Items.Item("29_U_C");
                        oItem.Visible = false;

                        oItem = oForm.Items.Item("30_U_S");
                        oItem.Visible = false;
                        oItem = oForm.Items.Item("30_U_E");
                        oItem.Visible = false;
                        oItem = oForm.Items.Item("30_U_E1");
                        oItem.Visible = false;
                        oItem = oForm.Items.Item("30_U_LB");
                        oItem.Visible = false;

                        oItem = oForm.Items.Item("31_U_S");
                        oItem.Visible = false;
                        oItem = oForm.Items.Item("31_U_E");
                        oItem.Visible = false;
                    }
                    else if (trnsType == "4") //სხვა
                    {
                        oItem = oForm.Items.Item("22_U_S");
                        oItem.Visible = true;

                        oItem = oForm.Items.Item("23_U_S");
                        oItem.Visible = true;
                        oItem = oForm.Items.Item("23_U_C");
                        oItem.Visible = true;

                        oItem = oForm.Items.Item("24_U_S");
                        oItem.Visible = true;
                        oItem = oForm.Items.Item("24_U_E");
                        oItem.Visible = true;
                        oItem.Enabled = true;
                        oItem = oForm.Items.Item("24_U_LB");
                        oItem.Visible = true;
                        oItem = oForm.Items.Item("24_U_E1");
                        oItem.Visible = true;
                        oItem.Enabled = true;

                        oItem = oForm.Items.Item("25_U_S");
                        oItem.Visible = false;
                        oItem = oForm.Items.Item("25_U_E");
                        oItem.Visible = false;

                        oItem = oForm.Items.Item("26_U_S");
                        oItem.Visible = false;
                        oItem = oForm.Items.Item("26_U_E");
                        oItem.Visible = false;
                        oItem = oForm.Items.Item("26_U_LB");
                        oItem.Visible = false;
                        oItem = oForm.Items.Item("26_U_B");
                        oItem.Visible = false;

                        oItem = oForm.Items.Item("27_U_S");
                        oItem.Visible = false;
                        oItem = oForm.Items.Item("27_U_E");
                        oItem.Visible = false;

                        oItem = oForm.Items.Item("28_U_CH");
                        oItem.Visible = false;

                        oItem = oForm.Items.Item("29_U_S");
                        oItem.Visible = false;
                        oItem = oForm.Items.Item("29_U_E");
                        oItem.Visible = false;
                        oItem = oForm.Items.Item("29_U_C");
                        oItem.Visible = false;

                        oItem = oForm.Items.Item("30_U_S");
                        oItem.Visible = false;
                        oItem = oForm.Items.Item("30_U_E");
                        oItem.Visible = false;
                        oItem = oForm.Items.Item("30_U_E1");
                        oItem.Visible = false;
                        oItem = oForm.Items.Item("30_U_LB");
                        oItem.Visible = false;

                        oItem = oForm.Items.Item("31_U_S");
                        oItem.Visible = false;
                        oItem = oForm.Items.Item("31_U_E");
                        oItem.Visible = false;
                    }
                }
                //<---

                List<string> listValidValues = new List<string>(); //--->
                string waybillStatus = oForm.DataSources.DBDataSources.Item("@BDO_WBLD").GetValue("U_status", 0).Trim();
                //string waybillID = oForm.DataSources.DBDataSources.Item("@BDO_WBLD").GetValue("U_wblID", 0).Trim();

                if (waybillStatus == "-1" || waybillStatus == "")
                {
                    listValidValues.Add(BDOSResources.getTranslate("RSCreate"));
                    if (trnsType == "6") //გადამზიდავი საავტომობილო
                    {
                        listValidValues.Add(BDOSResources.getTranslate("RSSendToTransporter"));
                    }
                    else
                    {
                        listValidValues.Add(BDOSResources.getTranslate("RSActivation"));
                    }
                    listValidValues.Add(BDOSResources.getTranslate("RSUpdateStatus"));
                }
                else if (waybillStatus == "1") //"შენახული"
                {
                    if (trnsType == "6") //გადამზიდავი საავტომობილო
                    {
                        listValidValues.Add(BDOSResources.getTranslate("RSSendToTransporter"));
                    }
                    else
                    {
                        listValidValues.Add(BDOSResources.getTranslate("RSActivation"));
                    }
                    listValidValues.Add(BDOSResources.getTranslate("RSUpdateStatus"));
                }
                else if (waybillStatus == "2") //"აქტიური"
                {
                    listValidValues.Add(BDOSResources.getTranslate("RSCorrection"));
                    listValidValues.Add(BDOSResources.getTranslate("RSFinish"));
                    listValidValues.Add(BDOSResources.getTranslate("RSCancel"));
                    listValidValues.Add(BDOSResources.getTranslate("RSUpdateStatus"));
                }
                else if (waybillStatus == "3") //"დასრულებული"
                {
                    listValidValues.Add(BDOSResources.getTranslate("RSCorrection"));
                    listValidValues.Add(BDOSResources.getTranslate("RSCancel"));
                    listValidValues.Add(BDOSResources.getTranslate("RSUpdateStatus"));
                }
                else if (waybillStatus == "4" || waybillStatus == "5") //"წაშლილი" || "გაუქმებული"
                {
                    listValidValues.Add(BDOSResources.getTranslate("RSCreate"));
                    if (trnsType == "6") //გადამზიდავი საავტომობილო
                    {
                        listValidValues.Add(BDOSResources.getTranslate("RSSendToTransporter"));
                    }
                    else
                    {
                        listValidValues.Add(BDOSResources.getTranslate("RSActivation"));
                    }
                    listValidValues.Add(BDOSResources.getTranslate("RSUpdateStatus"));
                }
                else if (waybillStatus == "6") //"გადამზიდავთან გადაგზავნილი"
                {
                    listValidValues.Add(BDOSResources.getTranslate("RSFinish"));
                    listValidValues.Add(BDOSResources.getTranslate("RSUpdateStatus"));
                }

                SAPbouiCOM.ButtonCombo oButtonCombo = ((SAPbouiCOM.ButtonCombo)(oForm.Items.Item("33_U_BC").Specific));
                int count = oButtonCombo.ValidValues.Count;

                for (int i = 0; i < count; i++)
                {
                    oButtonCombo.ValidValues.Remove(i.ToString(), SAPbouiCOM.BoSearchKey.psk_ByValue);
                }
                for (int i = 0; i < listValidValues.Count(); i++)
                {
                    oButtonCombo.ValidValues.Add(i == 0 && listValidValues[i] == "" ? "-1" : i.ToString(), listValidValues[i]);
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
            
            FormsB1.WB_TAX_AuthorizationsItems(oForm);

        }

        public static void formDataLoad(SAPbouiCOM.Form oForm)
        {
            string caption = "";

            try
            {
                string baseDocType = oForm.DataSources.DBDataSources.Item("@BDO_WBLD").GetValue("U_baseDocT", 0).Trim();

                switch (baseDocType)
                {
                    case "13":
                        caption = BDOSResources.getTranslate("ARInvoice"); //"A/R Invoice";
                        break;
                    case "15":
                        caption = BDOSResources.getTranslate("Delivery"); //"Delivery";
                        break;
                    case "165":
                        caption = "AR Correction Invoice"; //AR Correction Invoice
                        break;
                    case "67":
                        caption = BDOSResources.getTranslate("StockTransfer"); //"Inventory Transfer";
                        break;
                    case "UDO_F_BDOSFASTRD_D":
                        caption = BDOSResources.getTranslate("FixedAssetTransferDocument"); //"Fixed AssetTransfer";
                        break;
                    case "14":
                        caption = BDOSResources.getTranslate("ARCreditMemo"); //"A/R Credit Memo";
                        break;
                    case "60":
                        caption = BDOSResources.getTranslate("GoodsIssue"); //"Goods Issue";
                        break;
                }

                SAPbouiCOM.StaticText oStaticText = (SAPbouiCOM.StaticText)oForm.Items.Item("13_U_S").Specific;
                oStaticText.Caption = caption;

                SAPbouiCOM.EditText oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("14_U_E").Specific;
                oEditText.ChooseFromListUID = "BaseDoc_CFL" + baseDocType;
                oEditText.ChooseFromListAlias = "DocEntry";

                SAPbouiCOM.LinkedButton oLinkedButton = (SAPbouiCOM.LinkedButton)oForm.Items.Item("14_U_LB").Specific;
                oLinkedButton.LinkedObjectType = baseDocType;
                
                oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("20_U_E").Specific;
                if (baseDocType == "13" || baseDocType == "14" || baseDocType == "165" || oForm.DataSources.DBDataSources.Item("@BDO_WBLD").GetValue("U_cardCode", 0).Trim() != "")
                {
                    oEditText.ChooseFromListUID = "Contact_CFL"; //საკონტაქტო პირი
                    oEditText.ChooseFromListAlias = "Name";
                }
                else
                {
                    oEditText.ChooseFromListUID = "Employee_CFL"; //თანამშრომელი
                    oEditText.ChooseFromListAlias = "empID";
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                GC.Collect();
            }
        }

        public static void comboSelect(SAPbouiCOM.Form oForm, string itemUID, bool before_Action)
        {
            string errorText;
            try
            {
                SAPbouiCOM.ButtonCombo oButtonCombo = ((SAPbouiCOM.ButtonCombo)(oForm.Items.Item("33_U_BC").Specific));

                if (!before_Action)
                {
                    string waybillType = oForm.DataSources.DBDataSources.Item("@BDO_WBLD").GetValue("U_type", 0).Trim();
                    string baseDocType = oForm.DataSources.DBDataSources.Item("@BDO_WBLD").GetValue("U_baseDocT", 0).Trim();

                    if (itemUID == "6_U_C")
                    {
                        if (waybillType == "1") //თუ ტრანსპორტირების გარეშეა უნდა გასუფთავდეს ყველა ველი რაც ტრანსპორტირებას ეხება
                        {
                            oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_trnsType", 0, "-1");
                            oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_vehicle", 0, "");
                            oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_vehicNum", 0, "");
                            oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_trailNum", 0, "");
                            oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_drvCode", 0, "");
                            oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_drivTin", 0, "");
                            oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_notRsdnt", 0, "N");
                            oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_trnsExpn", 0, "0");
                            oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_payForTr", 0, "-1");
                            oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_tporter", 0, "");
                            oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_tporterN", 0, "");
                            oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_tporterT", 0, "");
                        }
                        else if (waybillType == "0")
                        {
                            oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_trnsType", 0, "1");
                        }
                    }

                    if (itemUID == "23_U_C")
                    {
                        if (waybillType == "0") //ტრანსპორტირებით 
                        {
                            string trnsType = oForm.DataSources.DBDataSources.Item("@BDO_WBLD").GetValue("U_trnsType", 0).Trim(); //--->

                            if (trnsType == "6") //გადამზიდავი საავტომობილო
                            {
                                oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_vehicle", 0, "");
                                oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_vehicNum", 0, "");
                                oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_trailNum", 0, "");
                                oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_drvCode", 0, "");
                                oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_drivTin", 0, "");
                                oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_notRsdnt", 0, "N");
                            }
                            else if (trnsType == "1" || trnsType == "5") //საავტომობილო || საავტომობილო უცხო ქვეყნის
                            {
                                oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_trnsExpn", 0, "0");
                                oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_payForTr", 0, "-1");
                                oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_tporter", 0, "");
                                oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_tporterN", 0, "");
                                oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_tporterT", 0, "");
                            }
                            else if (trnsType == "2" || trnsType == "3") //სარკინიგზო || საავიაციო
                            {
                                oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_vehicle", 0, "");
                                oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_vehicNum", 0, "");
                                oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_trailNum", 0, "");
                                oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_drvCode", 0, "");
                                oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_drivTin", 0, "");
                                oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_notRsdnt", 0, "N");
                                oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_trnsExpn", 0, "0");
                                oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_payForTr", 0, "-1");
                                oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_tporter", 0, "");
                                oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_tporterN", 0, "");
                                oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_tporterT", 0, "");
                            }
                            else if (trnsType == "4") //სხვა
                            {
                                oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_trailNum", 0, "");
                                oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_drvCode", 0, "");
                                oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_drivTin", 0, "");
                                oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_notRsdnt", 0, "N");
                                oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_trnsExpn", 0, "0");
                                oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_payForTr", 0, "-1");
                                oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_tporter", 0, "");
                                oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_tporterN", 0, "");
                                oForm.DataSources.DBDataSources.Item("@BDO_WBLD").SetValue("U_tporterT", 0, "");
                            }
                        }
                    }

                    if (itemUID == "33_U_BC")
                    {
                        string operationRS = null;
                        if (oButtonCombo.Selected != null)
                        {
                            operationRS = oButtonCombo.Selected.Description;
                        }

                        oForm.Freeze(false);
                        oButtonCombo.Caption = BDOSResources.getTranslate("Operations");

                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            uiApp.MessageBox(BDOSResources.getTranslate("ToCompleteOperationWriteDocument"));
                            return;
                        }

                        //printUDO( oForm.DataSources.DBDataSources.Item("@BDO_WBLD").GetValue("DocEntry", 0));
                        //return;
                        int answer = 0;
                        answer = Program.uiApp.MessageBox(BDOSResources.getTranslate("DoYouWantToContinue") + " " + BDOSResources.getTranslate(operationRS) + "?", 1, BDOSResources.getTranslate("Yes"), BDOSResources.getTranslate("No"), "");
                        if (answer == 1)
                        {
                            if (operationRS == BDOSResources.getTranslate("RSCreate"))
                            {
                                int docEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("@BDO_WBLD").GetValue("DocEntry", 0));
                                int baseDocEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("@BDO_WBLD").GetValue("U_baseDoc", 0));
                                //string baseDocType = oForm.DataSources.DBDataSources.Item("@BDO_WBLD").GetValue("U_baseDocT", 0).Trim();
                                saveWaybill(docEntry, baseDocEntry, operationRS, out errorText);
                                if (errorText != null)
                                {
                                    uiApp.MessageBox(errorText);
                                }
                                else
                                {
                                    uiApp.MessageBox(BDOSResources.getTranslate("WaybillCreatedSuccesfully"));
                                }

                            }
                            else if (operationRS == BDOSResources.getTranslate("RSActivation"))
                            {
                                int docEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("@BDO_WBLD").GetValue("DocEntry", 0));
                                int baseDocEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("@BDO_WBLD").GetValue("U_baseDoc", 0));
                                //string baseDocType = oForm.DataSources.DBDataSources.Item("@BDO_WBLD").GetValue("U_baseDocT", 0).Trim();
                                saveWaybill(docEntry, baseDocEntry, operationRS, out errorText);
                                if (errorText != null)
                                {
                                    uiApp.MessageBox(errorText);
                                }
                                else
                                {
                                    uiApp.MessageBox(BDOSResources.getTranslate("WaybillActivateSuccesfully"));
                                }
                            }
                            else if (operationRS == BDOSResources.getTranslate("RSSendToTransporter"))
                            {
                                int docEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("@BDO_WBLD").GetValue("DocEntry", 0));
                                int baseDocEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("@BDO_WBLD").GetValue("U_baseDoc", 0));
                                //string baseDocType = oForm.DataSources.DBDataSources.Item("@BDO_WBLD").GetValue("U_baseDocT", 0).Trim();
                                saveWaybill(docEntry, baseDocEntry, operationRS, out errorText);
                                if (errorText != null)
                                {
                                    uiApp.MessageBox(errorText);
                                }
                                else
                                {
                                    uiApp.MessageBox(BDOSResources.getTranslate("WaybillSentToTransporterSuccesfully"));
                                }
                            }
                            else if (operationRS == BDOSResources.getTranslate("RSCorrection"))
                            {
                                int docEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("@BDO_WBLD").GetValue("DocEntry", 0));
                                int baseDocEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("@BDO_WBLD").GetValue("U_baseDoc", 0));
                                //string baseDocType = oForm.DataSources.DBDataSources.Item("@BDO_WBLD").GetValue("U_baseDocT", 0).Trim();
                                saveWaybill(docEntry, baseDocEntry, operationRS, out errorText);
                                if (errorText != null)
                                {
                                    uiApp.MessageBox(errorText);
                                }
                                else
                                {
                                    uiApp.MessageBox(BDOSResources.getTranslate("WaybillCorrectedSuccesfully"));
                                }
                            }
                            else if (operationRS == BDOSResources.getTranslate("RSFinish"))
                            {
                                int docEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("@BDO_WBLD").GetValue("DocEntry", 0));
                                int baseDocEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("@BDO_WBLD").GetValue("U_baseDoc", 0));
                                closeWaybill(docEntry, baseDocEntry, out errorText);
                                if (errorText != null)
                                {
                                    uiApp.MessageBox(errorText);
                                }
                                else
                                {
                                    uiApp.MessageBox(BDOSResources.getTranslate("WaybillFinishedSuccesfully"));
                                }
                            }
                            else if (operationRS == BDOSResources.getTranslate("RSCancel"))
                            {
                                int docEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("@BDO_WBLD").GetValue("DocEntry", 0));
                                int baseDocEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("@BDO_WBLD").GetValue("U_baseDoc", 0));
                                refWaybill(docEntry, baseDocEntry, out errorText);
                                if (errorText != null)
                                {
                                    uiApp.MessageBox(errorText);
                                }
                                else
                                {
                                    uiApp.MessageBox(BDOSResources.getTranslate("WaybillCanceledSuccesfully"));
                                }
                            }
                            else if (operationRS == BDOSResources.getTranslate("RSUpdateStatus"))
                            {
                                int docEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("@BDO_WBLD").GetValue("DocEntry", 0));
                                int baseDocEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("@BDO_WBLD").GetValue("U_baseDoc", 0));
                                getWaybill(docEntry, baseDocEntry, out errorText);
                                if (errorText != null)
                                {
                                    uiApp.MessageBox(errorText);
                                }
                                else
                                {
                                    uiApp.MessageBox(BDOSResources.getTranslate("WaybillUpdatedStatusSuccesfully"));
                                }
                            }
                            if (operationRS != null)
                            {
                                FormsB1.SimulateRefresh();
                            }
                        }
                    }
                }
                else
                {
                    if (itemUID == "33_U_BC")
                    {
                        oForm.Freeze(true);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                GC.Collect();
            }
        }

        public static Dictionary<string, string> getWaybillDocumentInfo(int docEntry, string baseDocType, out string errorText)
        {
            errorText = null;

            Dictionary<string, string> wblDocInfo = new Dictionary<string, string>();
            wblDocInfo.Add("DocEntry", "0");
            wblDocInfo.Add("wblID", "");
            wblDocInfo.Add("number", "");
            wblDocInfo.Add("status", "");
            wblDocInfo.Add("statusN", "-1");
            wblDocInfo.Add("actDate", new DateTime().ToString());
            wblDocInfo.Add("CreateDate", new DateTime().ToString());

            Recordset oRecordSet = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            string query = "SELECT " +
                "\"DocEntry\", \"U_wblID\", \"U_status\", \"U_number\", \"U_actDate\", \"CreateDate\"" +
                "FROM \"@BDO_WBLD\" " +
                "WHERE  \"Canceled\"='N' AND \"U_baseDoc\" = '" + docEntry + "' AND \"U_baseDocT\" = '" + baseDocType + "'";

            try
            {
                oRecordSet.DoQuery(query);

                while (!oRecordSet.EoF)
                {
                    wblDocInfo["DocEntry"] = oRecordSet.Fields.Item("DocEntry").Value.ToString();
                    wblDocInfo["wblID"] = oRecordSet.Fields.Item("U_wblID").Value.ToString();
                    wblDocInfo["number"] = oRecordSet.Fields.Item("U_number").Value.ToString();
                    wblDocInfo["status"] = statusAsString(oRecordSet.Fields.Item("U_status").Value.ToString());
                    wblDocInfo["statusN"] = oRecordSet.Fields.Item("U_status").Value.ToString();
                    wblDocInfo["actDate"] = oRecordSet.Fields.Item("U_actDate").Value.ToString();
                    wblDocInfo["CreateDate"] = oRecordSet.Fields.Item("CreateDate").Value.ToString();

                    oRecordSet.MoveNext();
                    break;
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                GC.Collect();
            }

            return wblDocInfo;
        }

        public static void cancellation(int docEntry, string operation, out string errorText)
        {
            errorText = null;

            CompanyService oCompanyService = null;
            GeneralService oGeneralService = null;
            GeneralData oGeneralData = null;
            GeneralDataParams oGeneralParams = null;
            oCompanyService = oCompany.GetCompanyService();
            oGeneralService = oCompanyService.GetGeneralService("UDO_F_BDO_WBLD_D");
            oGeneralData = ((GeneralData)(oGeneralService.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralData)));

            try
            {
                //Get UDO record
                oGeneralParams = oGeneralService.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralDataParams);
                oGeneralParams.SetProperty("DocEntry", docEntry);
                oGeneralData = oGeneralService.GetByParams(oGeneralParams);

                if (operation == "Update")
                {
                    oGeneralData.SetProperty("U_baseDoc", 0);
                    oGeneralData.SetProperty("U_baseDTxt", "");
                    oGeneralService.Update(oGeneralData);
                }
                else if (operation == "Cancel")
                {
                    oGeneralService.Cancel(oGeneralParams);
                }
                else if (operation == "Close")
                {
                    oGeneralService.Close(oGeneralParams);
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oGeneralService);
                GC.Collect();
            }
        }

        public static string statusAsString(string WbStatus)
        {
            string WbstatusString = "";

            if (WbStatus == "-1")
            {
                WbstatusString = "";
            }
            if (WbStatus == "1")
            {
                WbstatusString = BDOSResources.getTranslate("Saved");
            }
            if (WbStatus == "2")
            {
                WbstatusString = BDOSResources.getTranslate("Active");
            }
            if (WbStatus == "3")
            {
                WbstatusString = BDOSResources.getTranslate("finished");
            }
            if (WbStatus == "4")
            {
                WbstatusString = BDOSResources.getTranslate("deleted");
            }
            if (WbStatus == "5")
            {
                WbstatusString = BDOSResources.getTranslate("Canceled");
            }
            if (WbStatus == "6")
            {
                WbstatusString = BDOSResources.getTranslate("SentToTransporter");
            }

            return WbstatusString;
        }

        public static string trnsTypeAsString(string trnsType)
        {
            string trnsTypeString = "";

            if (trnsType == "-1")
            {
                trnsTypeString = "";
            }
            if (trnsType == "1")
            {
                trnsTypeString = BDOSResources.getTranslate("Auto");
            }
            if (trnsType == "2")
            {
                trnsTypeString = BDOSResources.getTranslate("Railway");
            }
            if (trnsType == "3")
            {
                trnsTypeString = BDOSResources.getTranslate("Aviation");
            }
            if (trnsType == "4")
            {
                trnsTypeString = BDOSResources.getTranslate("other");
            }
            if (trnsType == "5")
            {
                trnsTypeString = BDOSResources.getTranslate("AutoOtherCountry");
            }
            if (trnsType == "6")
            {
                trnsTypeString = BDOSResources.getTranslate("AutoTransporter");
            }

            return trnsTypeString;
        }

        public static bool printUDO(string strDocNum)
        {
            // get menu UID of report
            Recordset oRS = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            oRS.DoQuery("SELECT \"MenuUID\" FROM \"OCMN\" WHERE \"Name\" = 'INV_WBLD - NEW' AND \"Type\" = 'C'");

            if (oRS.RecordCount == 0)
            {
                uiApp.MessageBox("Report layout 'ReportName' not found.", 0, "OK", null, null);
                return false;
            }

            // execute menu and enter document number
            uiApp.ActivateMenuItem(oRS.Fields.Item(0).Value.ToString()); //21481e483b1f42f8a9999f67652e8fa1
            //Program.uiApp.ActivateMenuItem("21481e483b1f42f8a9999f67652e8fa1");
            SAPbouiCOM.Form form = uiApp.Forms.ActiveForm;
            ((SAPbouiCOM.EditText)form.Items.Item("1000003").Specific).String = strDocNum;
            form.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular); // abrir reporte
            form.Close();
            return true;
        }

        public static void uiApp_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            SAPbouiCOM.Form oForm = uiApp.Forms.GetForm(BusinessObjectInfo.FormTypeEx, currentFormCount);


           if (BusinessObjectInfo.BeforeAction)
            {
                string errorText = null;
                FormsB1.WB_TAX_AuthorizationsOperations(BusinessObjectInfo.FormTypeEx, BusinessObjectInfo.EventType ,  out errorText);
                if (errorText!=null)
                {
                    uiApp.SetStatusBarMessage(errorText);
                    uiApp.MessageBox(errorText);
                    BubbleEvent = false;
                }
            }

            if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD && !BusinessObjectInfo.BeforeAction)
            {
                formDataLoad(oForm);
                setVisibleFormItems(oForm);
            }
            else if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD && BusinessObjectInfo.BeforeAction)
            {
                if (oForm.DataSources.DBDataSources.Item("@BDO_WBLD").GetValue("U_baseDocT", 0).Trim() == "")
                {
                    uiApp.MessageBox(BDOSResources.getTranslate("CreateWaybillAllowedBasedOnlyOtherDocument"));
                    BubbleEvent = false;
                }
            }
            //}
        }

        public static void uiApp_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD && pVal.BeforeAction)
                {
                    createFormItems(oForm, out errorText);
                    FORM_LOAD_FOR_VISIBLE = true;
                    FORM_LOAD_FOR_ACTIVATE = true;
                }

                else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE && !pVal.BeforeAction)
                {
                    if (FORM_LOAD_FOR_VISIBLE)
                    {
                        setSizeForm(oForm);
                        oForm.Freeze(true);
                        oForm.Title = BDOSResources.getTranslate("Waybill");
                        oForm.Freeze(false);
                        FORM_LOAD_FOR_VISIBLE = false;
                        setVisibleFormItems(oForm);
                    }
                }

                else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_RESIZE && !pVal.BeforeAction)
                {
                    resizeForm(oForm);
                }

                else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_COMBO_SELECT)
                {
                    oForm.Freeze(true);
                    comboSelect(oForm, pVal.ItemUID, pVal.BeforeAction);
                    oForm.Freeze(false);
                    if (!pVal.BeforeAction)
                    {
                        setVisibleFormItems(oForm);
                        oForm.Update();
                    }
                }

                else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE && !pVal.BeforeAction)
                {
                    if (FORM_LOAD_FOR_ACTIVATE)
                    {
                        oForm.Freeze(true);
                        try
                        {
                            SAPbouiCOM.StaticText staticText = oForm.Items.Item("0_U_S").Specific;
                            staticText.Caption = BDOSResources.getTranslate("DocEntry");

                            FORM_LOAD_FOR_ACTIVATE = false;
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

                //else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE && !pVal.BeforeAction)
                //{
                //    if (Program.FORM_LOAD_FOR_ACTIVATE)
                //    {
                //        //fillNewDocument( oForm, out errorText);
                //        oForm.Freeze(true);
                //        setVisibleFormItems(oForm);
                //        formDataLoad(oForm, out errorText);
                //        oForm.Freeze(false);
                //        oForm.Update();
                //        Program.FORM_LOAD_FOR_ACTIVATE = false;
                //    }
                //}

                else if ((pVal.ItemUID == "14_U_E" || pVal.ItemUID == "17_U_E" || pVal.ItemUID == "20_U_E" || pVal.ItemUID == "24_U_E" || pVal.ItemUID == "26_U_B" || pVal.ItemUID == "30_U_E") && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                {
                    SAPbouiCOM.IChooseFromListEvent oCFLEvento = null;
                    oCFLEvento = (SAPbouiCOM.IChooseFromListEvent)pVal;

                    chooseFromList(oForm, oCFLEvento, pVal.ItemUID, pVal.BeforeAction);
                }
                else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK & pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "33_U_BC")
                    {
                        FormsB1.WB_TAX_AuthorizationsOperations("UDO_FT_UDO_F_BDO_WBLD_D", SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE, out errorText);
                        if (errorText != null)
                        {
                            Program.uiApp.SetStatusBarMessage(errorText);
                            Program.uiApp.MessageBox(errorText);
                            return;
                        }
                    }

                }
            }
        }

        #region RS.GE
        public static void saveWaybill(int docEntry, int baseDocEntry, string operationRS, out string errorText)
        {
            errorText = null;

            Dictionary<string, string> rsSettings = CompanyDetails.getRSSettings(out errorText);
            if (errorText != null)
            {
                return;
            }

            string su = rsSettings["SU"];
            string sp = rsSettings["SP"];

            WayBill oWayBill = new WayBill(rsSettings["ProtocolType"]);

            bool chek_service_user = oWayBill.chek_service_user(su, sp, out errorText);
            if (chek_service_user == false)
            {
                errorText = BDOSResources.getTranslate("ServiceUserPasswordNotCorrect");
                return;
            }

            Recordset oRecordSet = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            string query = "SELECT " +
            "\"BDO_WBLD\".\"DocEntry\", " +
            "\"BDO_WBLD\".\"U_number\", " +
            "\"BDO_WBLD\".\"U_type\", " +
            "\"BDO_WBLD\".\"U_status\", " +
            "\"BDO_WBLD\".\"U_actDate\", " +
            "\"BDO_WBLD\".\"U_begDate\", " +
            "\"BDO_WBLD\".\"U_beginTime\", " +
            "\"BDO_WBLD\".\"U_strAddrs\", " +
            "\"BDO_WBLD\".\"U_endAddrs\", " +
            "\"BDO_WBLD\".\"U_comment\", " +
            "\"BDO_WBLD\".\"U_delvDate\", " +
            "\"BDO_WBLD\".\"U_trnsType\", " +
            "\"BDO_WBLD\".\"U_vehicle\", " +
            "\"BDO_WBLD\".\"U_vehicNum\", " +
            "\"BDO_WBLD\".\"U_trailNum\", " +
            "\"BDO_WBLD\".\"U_drvCode\", " +
            "\"BDO_WBLD\".\"U_drivTin\", " +
            "\"BDO_WBLD\".\"U_notRsdnt\", " +
            "\"BDO_WBLD\".\"U_tporter\", " +
            "\"BDO_WBLD\".\"U_tporterT\", " +
            "\"BDO_WBLD\".\"U_payForTr\", " +
            "\"BDO_WBLD\".\"U_trnsExpn\", " +
            "\"BDO_WBLD\".\"U_recpInfo\", " +
            "\"BDO_WBLD\".\"U_recvInfo\", " +
            "\"BDO_WBLD\".\"U_wblID\", " +
            "\"BDO_WBLD\".\"U_recpInfN\", " +
            "\"BDO_WBLD\".\"U_recvInfN\", " +
            "\"BDO_WBLD\".\"U_tporterN\", " +
            "\"BDO_WBLD\".\"U_baseDoc\", " +
            "\"BDO_WBLD\".\"U_baseDocT\", " +
            "\"BDO_WBLD\".\"U_cardCode\", " +
            "\"OCRD\".\"CardName\", " +
            "\"OCRD\".\"LicTradNum\", " +
            "\"OCRD\".\"U_BDO_TaxTyp\", " +
            "\"OCRD\".\"Country\" " +
            "FROM \"@BDO_WBLD\" AS \"BDO_WBLD\" " +
            "LEFT JOIN \"OCRD\" AS \"OCRD\" " +
            "ON \"BDO_WBLD\".\"U_cardCode\" = \"OCRD\".\"CardCode\" " +
            "WHERE \"BDO_WBLD\".\"DocEntry\" = '" + docEntry + "'";

            string[] array_HEADER = null;
            string U_baseDocT = null;

            try
            {
                oRecordSet.DoQuery(query);
                string ID = "";
                string STATUS = null;
                DateTime BEGIN_DATE = DateTime.MinValue;

                while (!oRecordSet.EoF)
                {
                    U_baseDocT = oRecordSet.Fields.Item("U_baseDocT").Value.ToString();

                    ID = oRecordSet.Fields.Item("U_wblID").Value.ToString(); //ID - გადაეცემა 0 თუ იქმნება ახალი
                    ID = ID == "" ? "0" : ID;
                    string TYPE = oRecordSet.Fields.Item("U_type").Value.ToString(); //TYPE - ზედნადების ტიპი
                    switch (U_baseDocT)
                    {
                        case "13":
                            TYPE = TYPE == "0" ? "2" : "3"; //ტრანსპორტირებით = 2 //ტრანსპორტირების გარეშე = 3 //A/R Invoice
                            break;
                        case "15":
                            TYPE = TYPE == "0" ? "2" : "3"; //ტრანსპორტირებით = 2 //ტრანსპორტირების გარეშე = 3 //Delivery
                            break;
                        case "67":
                            TYPE = "1"; //შიდა გადაზიდვა  //Inventory Transfer
                            break;
                        case "UDO_F_BDOSFASTRD_D":
                            TYPE = "1"; //შიდა გადაზიდვა  //Fixed Asset Transfer
                            break;
                        case "14":
                        case "165":
                            TYPE = "5"; //უკან დაბრუნება  //A/R Credit Memo //A/R Correction Invoice
                            break;
                        case "60":
                            TYPE = "1"; //შიდა გადაზიდვა  //Goods Issue
                            break;
                    }

                    string IS_NON_RESIDENT = oRecordSet.Fields.Item("U_BDO_TaxTyp").Value.ToString(); //IS_NON_RESIDENT - თუ უცხოელია 10
                    string BUYER_TIN = (U_baseDocT == "67" || U_baseDocT == "60" || U_baseDocT == "UDO_F_BDOSFASTRD_D") ? "" : oRecordSet.Fields.Item("LicTradNum").Value.ToString(); //BUYER_TIN - მყიდველის პირადი ან საიდენტიფიკაციო ნომერი
                    string CHEK_BUYER_TIN = (U_baseDocT == "67" || U_baseDocT == "60" || U_baseDocT == "UDO_F_BDOSFASTRD_D") ? "GE" : oRecordSet.Fields.Item("Country").Value.ToString(); //CHEK_BUYER_TIN – თუ უცხოელია 0 თუ საქართველოს მოქალაქე 1
                    CHEK_BUYER_TIN = CHEK_BUYER_TIN == "GE" ? "1" : "0";
                    if (IS_NON_RESIDENT == "10") CHEK_BUYER_TIN = "0";
                    string BUYER_NAME = (U_baseDocT == "67" || U_baseDocT == "60" || U_baseDocT == "UDO_F_BDOSFASTRD_D") ? "" : oRecordSet.Fields.Item("CardName").Value.ToString(); //BUYER_NAME - მყიდველის სახელი
                    string START_ADDRESS = oRecordSet.Fields.Item("U_strAddrs").Value.ToString(); //START_ADDRESS - ტრანსპორტირების დაწყების ადგილი
                    string END_ADDRESS = oRecordSet.Fields.Item("U_endAddrs").Value.ToString(); //END_ADDRESS - ტრანსპორტირების დასრულების ადგილი
                    string DRIVER_TIN = oRecordSet.Fields.Item("U_drivTin").Value.ToString(); //DRIVER_TIN - მძღოლის პირადი ნომერი
                    string CHEK_DRIVER_TIN = oRecordSet.Fields.Item("U_notRsdnt").Value.ToString(); //CHEK_DRIVER_TIN – თუ უცხოელია 0 თუ საქართველოს მოქალაქე 1
                    CHEK_DRIVER_TIN = (CHEK_DRIVER_TIN == "N" || string.IsNullOrEmpty(CHEK_DRIVER_TIN)) ? "1" : "0";
                    string DRIVER_NAME = oRecordSet.Fields.Item("U_drvCode").Value.ToString(); //DRIVER_NAME -მძღოლის სახელი
                    string TRANSPORT_COAST = oRecordSet.Fields.Item("U_trnsExpn").Value.ToString(); //TRANSPORT_COAST -ტრანსპორტირების ღირებულება
                    string RECEPTION_INFO = oRecordSet.Fields.Item("U_recpInfN").Value.ToString(); //RECEPTION_INFO - მიმწოდებლის ინფორმაცია
                    string RECEIVER_INFO = oRecordSet.Fields.Item("U_recvInfN").Value.ToString(); //RECEIVER_INFO - მიმღების ინფორმაცია               
                    DateTime DELIVERY_DATE = DateTime.MinValue; //DELIVERY_DATE - მიწოდების თარიღი გადასცემთ უკვე აქტიურს დახურვის წინ               
                    string STATUS_DOC = oRecordSet.Fields.Item("U_status").Value.ToString(); //STATUS - ზედნადების სტატუსი: 0-შენახული 1-აქტივირებული 2 დახურული

                    //switch (operationRS)
                    //{
                    if (operationRS == BDOSResources.getTranslate("RSCreate"))
                    {
                        STATUS = "0";

                    }
                    else if (operationRS == BDOSResources.getTranslate("RSActivation"))
                    {
                        STATUS = "1";

                    }
                    else if (operationRS == BDOSResources.getTranslate("RSSendToTransporter"))
                    {
                        STATUS = "8";
                    }

                    else if (operationRS == BDOSResources.getTranslate("RSCorrection"))
                    {
                        if (STATUS_DOC == "2") { STATUS = "1"; } //"აქტიური"
                        if (STATUS_DOC == "6") { STATUS = "5"; } //"გადამზიდავთან გადაგზავნილი"
                        if (STATUS_DOC == "3") { STATUS = "2"; } //"დასრულებული"                       

                    }
                    //}
                    string SELER_UN_ID = oWayBill.un_id.ToString(); //SELER_UN_ID - გამყიდველის უნიკალური ნომერი. გიბრუნებთ chek_service_user
                    string PAR_ID = "0"; //PAR_ID - მშობელი ზედნადების ID ქვე ზედნადების დროს TYPE = 6
                    string FULL_AMOUNT = "";
                    string CAR_NUMBER = oRecordSet.Fields.Item("U_vehicNum").Value.ToString(); //CAR_NUMBER - მანქანის ნომერი
                    string WAYBILL_NUMBER = oRecordSet.Fields.Item("U_number").Value.ToString();//WAYBILL_NUMBER
                    string S_USER_ID = oWayBill.un_user_id.ToString(); //S_USER_ID

                    BEGIN_DATE = DateTime.TryParse(oRecordSet.Fields.Item("U_begDate").Value.ToString("yyyyMMdd") == "18991230" ? "" : oRecordSet.Fields.Item("U_begDate").Value.ToString(), out BEGIN_DATE) == false ? DateTime.MinValue : BEGIN_DATE;  //BEGIN_DATE - ტრანსპორტირების დაწყების თარიღი
                    BEGIN_DATE = BEGIN_DATE == DateTime.MinValue || BEGIN_DATE < DateTime.Today ? DateTime.Now : BEGIN_DATE;

                    //ტრანსორტირების დაწყების საათები
                    decimal U_beginTime = Convert.ToDecimal(oRecordSet.Fields.Item("U_beginTime").Value);
                    int Hour = Convert.ToInt32(Math.Floor(U_beginTime / 100));
                    int Min = Convert.ToInt32(U_beginTime - Hour * 100);
                    BEGIN_DATE = new DateTime(BEGIN_DATE.Year, BEGIN_DATE.Month, BEGIN_DATE.Day, Hour, Min, 0);

                    string TRAN_COST_PAYER = oRecordSet.Fields.Item("U_payForTr").Value.ToString(); //TRAN_COST_PAYER- ტრანსპორტირების ღირებულებას თუ იხდის მყიდველი - 1; გამყიდველი - 2;
                    TRAN_COST_PAYER = TRAN_COST_PAYER == "-1" ? "0" : TRAN_COST_PAYER;
                    string TRANS_ID = oRecordSet.Fields.Item("U_trnsType").Value.ToString(); //TRANS_ID - ტრანსპორტის ტიპის id

                    if (TRANS_ID == "6" && STATUS == "1") //გადამზიდი საავტომობილო
                    {
                        errorText = BDOSResources.getTranslate("CannotActivateWaybillWithAutoTransporter");
                        return;
                    }
                    else if (TRANS_ID != "6" && STATUS == "5")
                    {
                        errorText = BDOSResources.getTranslate("WaybillTranportTypeIsNotAutoTransporter");
                        return;
                    }

                    string TRANS_TXT = ""; //TRANS_TXT - ტრანსპორტირების ტიპი, თუ არჩეულია „სხვა“ TRANS_ID = 4

                    if (TYPE == "2" || TYPE == "1" || TYPE == "5") //ტრანსპორტირებით || შიდა გადაზიდვა || უკან დაბრუნება
                    {
                        switch (TRANS_ID)
                        {
                            case "1":
                                TRANS_ID = "1"; TRANS_TXT = oRecordSet.Fields.Item("U_trailNum").Value.ToString();  //საავტომობილო
                                break;
                            case "2":
                                TRANS_ID = "2"; //სარკინიგზო
                                break;
                            case "3":
                                TRANS_ID = "3"; //საავიაციო
                                break;
                            case "4":
                                TRANS_ID = "4"; TRANS_TXT = oRecordSet.Fields.Item("U_vehicNum").Value.ToString(); //სხვა
                                break;
                            case "5":
                                TRANS_ID = "6"; TRANS_TXT = oRecordSet.Fields.Item("U_trailNum").Value.ToString(); //საავტომობილო - უცხო ქვეყნის
                                break;
                            case "6":
                                TRANS_ID = "7"; TRANS_TXT = oRecordSet.Fields.Item("U_trailNum").Value.ToString(); //გადამზიდი საავტომობილო
                                break;
                        }
                    }
                    else if (TYPE == "3") //ტრანსპორტირების გარეშე 
                    {
                        TRANS_ID = "0";
                        END_ADDRESS = START_ADDRESS;
                    }

                    string COMMENT = oRecordSet.Fields.Item("U_comment").Value.ToString(); //COMMENT - კომენტარი, შენიშვნა
                    string TRANSPORTER_TIN = oRecordSet.Fields.Item("U_tporterT").Value.ToString(); //TRANSPORTER_TIN - გადამზიდველი

                    array_HEADER = new string[28];

                    array_HEADER[0] = ID; //ID - გადაეცემა 0 თუ იქმნება ახალი
                    array_HEADER[1] = TYPE; //TYPE - ზედნადების ტიპი
                    array_HEADER[2] = BUYER_TIN; //BUYER_TIN - მყიდველის პირადი ან საიდენტიფიკაციო ნომერი
                    array_HEADER[3] = CHEK_BUYER_TIN; //CHEK_BUYER_TIN – თუ უცხოელია 0 თუ საქართველოს მოქალაქე 1
                    array_HEADER[4] = BUYER_NAME; //BUYER_NAME - მყიდველის სახელი
                    array_HEADER[5] = START_ADDRESS; //START_ADDRESS - ტრანსპორტირების დაწყების ადგილი
                    array_HEADER[6] = END_ADDRESS; //END_ADDRESS - ტრანსპორტირების დასრულების ადგილი
                    array_HEADER[7] = DRIVER_TIN; //DRIVER_TIN - მძღოლის პირადი ნომერი
                    array_HEADER[8] = CHEK_DRIVER_TIN; //CHEK_DRIVER_TIN – თუ უცხოელია 0 თუ საქართველოს მოქალაქე 1
                    array_HEADER[9] = DRIVER_NAME; //DRIVER_NAME -მძღოლის სახელი
                    array_HEADER[10] = TRANSPORT_COAST; //TRANSPORT_COAST -ტრანსპორტირების ღირებულება
                    array_HEADER[11] = RECEPTION_INFO; //RECEPTION_INFO - მიმწოდებლის ინფორმაცია
                    array_HEADER[12] = RECEIVER_INFO; //RECEIVER_INFO - მიმღების ინფორმაცია
                    array_HEADER[13] = DELIVERY_DATE == DateTime.MinValue ? "" : String.Format("{0:s}", DELIVERY_DATE); //DELIVERY_DATE - მიწოდების თარიღი გადასცემთ უკვე აქტიურს დახურვის წინ
                    array_HEADER[14] = STATUS; //STATUS - ზედნადების სტატუსი: 0-შენახული 1-აქტივირებული 2 დახურული
                    array_HEADER[15] = SELER_UN_ID; //SELER_UN_ID - გამყიდველის უნიკალური ნომერი. გიბრუნებთ chek_service_user
                    array_HEADER[16] = PAR_ID; //PAR_ID - მშობელი ზედნადების ID ქვე ზედნადების დროს TYPE=6
                    array_HEADER[17] = FULL_AMOUNT; //FULL_AMOUNT
                    array_HEADER[18] = CAR_NUMBER; //CAR_NUMBER - მანქანის ნომერი
                    array_HEADER[19] = WAYBILL_NUMBER; //WAYBILL_NUMBER
                    array_HEADER[20] = S_USER_ID; //S_USER_ID
                    array_HEADER[21] = String.Format("{0:s}", BEGIN_DATE); //"2016-11-02T10:15:21"; //BEGIN_DATE - ტრანსპორტირების დაწყების თარიღი //
                    array_HEADER[22] = TRAN_COST_PAYER; //TRAN_COST_PAYER- ტრანსპორტირების ღირებულებას თუ იხდის მყიდველი - 1; გამყიდველი - 2;
                    array_HEADER[23] = TRANS_ID; //TRANS_ID - ტრანსპორტის ტიპის id
                    array_HEADER[24] = TRANS_TXT; //TRANS_TXT - ტრანსპორტირების ტიპი, თუ არჩეულია „სხვა“ TRANS_ID=4
                    array_HEADER[25] = COMMENT; //COMMENT - კომენტარი, შენიშვნა
                    array_HEADER[26] = TRANSPORTER_TIN; //TRANSPORTER_TIN - გადამზიდველი

                    oRecordSet.MoveNext();
                    break;
                }

                string[][] array_GOODS = null;

                double QUANTITYRS = 0;
                double AMOUNTRS = 0;

                if (U_baseDocT == "13") //A/R Invoice
                {
                    getArrayGoodsARInvoiceType(oWayBill, baseDocEntry, ID, out array_GOODS, out QUANTITYRS, out AMOUNTRS, out errorText);
                }

                else if (U_baseDocT == "15") //Delivery
                {
                    getArrayGoodsDeliveryType(oWayBill, baseDocEntry, ID, out array_GOODS, out QUANTITYRS, out AMOUNTRS, out errorText);
                }
                else if (U_baseDocT == "67") //Inventory Transfer
                {
                    getArrayGoodsInventoryTransferType(oWayBill, baseDocEntry, ID, out array_GOODS, out QUANTITYRS, out errorText);
                }

                else if (U_baseDocT == "UDO_F_BDOSFASTRD_D") //Fixed Asset Transfer
                {
                    getArrayGoodsFixedAssetTransferType(oWayBill, baseDocEntry, ID, out array_GOODS, out QUANTITYRS, out errorText);
                }

                else if (U_baseDocT == "14") //A/R Credit Memo
                {
                    getArrayGoodsInvoiceCreditMemoType(oWayBill, baseDocEntry, ID, out array_GOODS, out QUANTITYRS, out AMOUNTRS, out errorText);
                }
                else if (U_baseDocT == "165") //A/R Correction Invoice
                {
                    getArrayGoodsInvoiceCorrectionType(oWayBill, baseDocEntry, ID, out array_GOODS, out QUANTITYRS, out AMOUNTRS, out errorText);
                }
                else if (U_baseDocT == "60") //Goods Issue
                {
                    getArrayGoodsGoodsIssueType(oWayBill, baseDocEntry, ID, out array_GOODS, out QUANTITYRS, out errorText);
                }
                if (errorText != null)
                {
                    return;
                }

                if (array_GOODS != null)
                {
                    int newGoods = 0;
                    for (int i = 0; i < array_GOODS.Count(); i++)
                    {
                        if (array_GOODS[i][6] != "-1")
                        {
                            newGoods = +1;
                        }
                    }
                    if (newGoods == 0)
                    {
                        if (operationRS == BDOSResources.getTranslate("RSCorrection"))
                        {
                            errorText = BDOSResources.getTranslate("WaybillTranportTypeIsNotAutoTransporter") + " " + BDOSResources.getTranslate("ReasonIs") + ":" + " " + BDOSResources.getTranslate("ItemsAreCorrected");
                        }
                        else
                        {
                            errorText = BDOSResources.getTranslate("WaybillTableIsEmpty");
                        }
                        return;
                    }
                }

                int save_waybill_result_int = oWayBill.save_waybill(array_HEADER, array_GOODS, out errorText);

                if (save_waybill_result_int != 1)
                {
                    errorText = array_HEADER[27];
                    return;
                }

                string ID_RS = array_HEADER[0];
                string WAYBILL_NUMBER_RS = array_HEADER[19];

                switch (STATUS)
                {
                    case "0":
                        STATUS = "1"; //"შენახული"
                        break;
                    case "1":
                        STATUS = "2";  //"აქტიური"
                        break;
                    case "2":
                        STATUS = "3";  //"დასრულებული"
                        break;
                    case "-1":
                        STATUS = "4";  //"წაშლილი"
                        break;
                    case "-2":
                        STATUS = "5";  //"გაუქმებული"
                        break;
                    case "8":
                        STATUS = "6";  //"გადამზიდავთან გადაგზავნილი"
                        break;
                }

                CompanyService oCompanyService = null;
                GeneralService oGeneralService = null;
                GeneralData oGeneralData = null;
                GeneralDataParams oGeneralParams = null;
                oCompanyService = oCompany.GetCompanyService();
                oGeneralService = oCompanyService.GetGeneralService("UDO_F_BDO_WBLD_D");
                oGeneralData = ((GeneralData)(oGeneralService.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralData)));

                //Get UDO record
                oGeneralParams = oGeneralService.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralDataParams);
                oGeneralParams.SetProperty("DocEntry", docEntry);
                oGeneralData = oGeneralService.GetByParams(oGeneralParams);

                //Update UDO record
                oGeneralData.SetProperty("U_wblID", ID_RS.ToString());
                oGeneralData.SetProperty("U_status", STATUS);
                oGeneralData.SetProperty("U_number", WAYBILL_NUMBER_RS);
                oGeneralData.SetProperty("U_begDate", BEGIN_DATE);
                if (STATUS == "2")
                {
                    oGeneralData.SetProperty("U_actDate", DateTime.Today);
                }

                oGeneralService.Update(oGeneralData);
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                GC.Collect();
            }
        }

        private static void getArrayGoodsARInvoiceType(WayBill oWayBill, int baseDocEntry, string ID, out string[][] array_GOODS, out double QUANTITYRS, out double AMOUNTRS, out string errorText)
        {
            errorText = null;
            array_GOODS = null;
            QUANTITYRS = 0;
            AMOUNTRS = 0;

            Dictionary<string, string> rsSettings = CompanyDetails.getRSSettings(out errorText);
            if (errorText != null)
            {
                return;
            }
            string itemCode = rsSettings["ItemCode"];

            Recordset oRecordSet = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            string query = "SELECT " +
            "'" + ID + "'" + " AS \"ID\", " +
            "\"MNTB\".\"LineNum\" AS \"LineNum\", " +
            "\"MNTB\".\"DocEntry\" AS \"DocEntry\", " +
            "\"MNTB\".\"ItemCode\" AS \"ItemCode\", " +
            "\"OITM\".\"CodeBars\" AS \"CodeBars\", " +
            "\"OITM\".\"SWW\" AS \"AdditionalIdentifier\", " +
            "\"MNTB\".\"Dscription\" AS \"W_NAME\", " +
            "CASE WHEN \"BDO_RSUOM\".\"U_RSCode\" is null THEN '99' ELSE \"BDO_RSUOM\".\"U_RSCode\" END AS \"UNIT_ID\", " +
            "CASE WHEN \"MNTB\".\"unitMsr\"='' THEN 'სხვა' ELSE \"MNTB\".\"unitMsr\" END  AS \"UNIT_TXT\", " +
            "\"MNTB\".\"VatPrcnt\" AS \"VAT_TYPE\", " +
            "\"MNTB\".\"VatGroup\"AS \"VatGroup\", " +
            "'0' AS \"A_ID\", " +
            "SUM(\"MNTB\".\"Quantity\") AS \"QUANTITY\", " +
            "SUM(\"MNTB\".\"GTotal\") AS \"AMOUNT\", " +
            "CASE WHEN SUM(\"MNTB\".\"Quantity\") = 0 THEN 0 ELSE SUM(\"MNTB\".\"GTotal\")/SUM(\"MNTB\".\"Quantity\") END AS \"PRICE\", " +
            "SUM(\"MNTB\".\"LineVat\") AS \"LineVat\" " +

            "FROM " +

            "(SELECT " +
            "\"INV1\".\"DocEntry\", " +
            "\"INV1\".\"LineNum\", " +
            "\"INV1\".\"ItemCode\", " +
            "\"INV1\".\"Dscription\", " +
            "\"INV1\".\"unitMsr\", " +
            "\"INV1\".\"Quantity\" * \"INV1\".\"NumPerMsr\" AS \"Quantity\", " +
            "\"INV1\".\"GTotal\", " +
            "\"INV1\".\"VatPrcnt\", " +
            "\"INV1\".\"VatGroup\", " +
            "\"INV1\".\"LineVat\" " +

            "FROM \"INV1\" " +

            "LEFT JOIN \"OITM\" AS \"OITM\" " +
            "ON \"INV1\".\"ItemCode\" = \"OITM\".\"ItemCode\" " +

            "WHERE \"INV1\".\"DocEntry\" = '" + baseDocEntry + "' AND (\"OITM\".\"ItemType\" = 'I' AND \"OITM\".\"InvntItem\" = 'Y' OR \"OITM\".\"ItemType\" = 'F' )  " +

            "UNION ALL " +

            "SELECT " +
            "\"RIN1\".\"BaseEntry\", " +
            "\"RIN1\".\"BaseLine\", " +
            "\"RIN1\".\"ItemCode\", " +
            "\"RIN1\".\"Dscription\", " +
            "\"RIN1\".\"unitMsr\", " +
            "\"RIN1\".\"Quantity\" * (-1) * (CASE WHEN \"RIN1\".\"NoInvtryMv\" = 'Y' THEN 0 ELSE 1 END) * \"RIN1\".\"NumPerMsr\", " +
            "\"RIN1\".\"GTotal\" * (-1), " +
            "\"RIN1\".\"VatPrcnt\", " +
            "\"RIN1\".\"VatGroup\", " +
            "\"RIN1\".\"LineVat\" * (-1) " +

            "FROM \"RIN1\" " +

            "INNER JOIN \"ORIN\" " +
            "ON \"ORIN\".\"DocEntry\" = \"RIN1\".\"DocEntry\" " +

            "LEFT JOIN \"OITM\" AS \"OITM\" " +
            "ON \"RIN1\".\"ItemCode\" = \"OITM\".\"ItemCode\" " +

            "WHERE \"RIN1\".\"BaseEntry\" = '" + baseDocEntry + "' AND \"RIN1\".\"TargetType\" < 0  AND \"ORIN\".\"U_BDO_CNTp\" <> 1 AND ((\"OITM\".\"ItemType\" = 'I' AND \"OITM\".\"InvntItem\" = 'Y') OR \"OITM\".\"ItemType\" = 'F' ) " +
            "UNION ALL " +

            "SELECT " +
                "\"CSI1\".\"BaseEntry\", " +
                "\"CSI1\".\"BaseLine\", " +
                "\"CSI1\".\"ItemCode\", " +
                "\"CSI1\".\"Dscription\", " +
                "\"CSI1\".\"unitMsr\", " +
                "\"CSI1\".\"Quantity\" * (CASE WHEN \"CSI1\".\"NoInvtryMv\" = 'Y' THEN 0 ELSE 1 END) * \"CSI1\".\"NumPerMsr\", " +
                "\"CSI1\".\"GTotal\" , " +
                "\"CSI1\".\"VatPrcnt\", " +
                "\"CSI1\".\"VatGroup\", " +
                "\"CSI1\".\"LineVat\"  " +

                "FROM \"CSI1\" " +

                "INNER JOIN \"OCSI\" " +
                "ON \"OCSI\".\"DocEntry\" = \"CSI1\".\"DocEntry\" " +

                "LEFT JOIN \"OITM\" AS \"OITM\" " +
                "ON \"CSI1\".\"ItemCode\" = \"OITM\".\"ItemCode\" " +

                "WHERE \"CSI1\".\"BaseEntry\" = '" + baseDocEntry + "' AND \"CSI1\".\"TargetType\" < 0  AND \"OCSI\".\"U_BDOSCITp\" <> 1 AND ((\"OITM\".\"ItemType\" = 'I' AND \"OITM\".\"InvntItem\" = 'Y') OR \"OITM\".\"ItemType\" = 'F' ) ) AS \"MNTB\" " +

            "LEFT JOIN \"OITM\" AS \"OITM\" " +
            "ON \"MNTB\".\"ItemCode\" = \"OITM\".\"ItemCode\" " +

            "LEFT JOIN \"OUOM\" AS \"OUOM\" " +
            "ON \"MNTB\".\"unitMsr\" = \"OUOM\".\"UomName\" " +

            "LEFT JOIN \"@BDO_RSUOM\" AS \"BDO_RSUOM\" " +
            "ON \"OUOM\".\"UomEntry\" = \"BDO_RSUOM\".\"U_UomEntry\" " +

            "GROUP BY " +
            "\"MNTB\".\"DocEntry\", " +
            "\"MNTB\".\"LineNum\", " +
            "\"MNTB\".\"ItemCode\", " +
            "\"MNTB\".\"Dscription\", " +
            "\"OITM\".\"CodeBars\", " +
            "\"OITM\".\"SWW\", " +
            "\"BDO_RSUOM\".\"U_RSCode\", " +
            "\"MNTB\".\"unitMsr\", " +
            "\"MNTB\".\"VatPrcnt\", " +
            "\"MNTB\".\"VatGroup\" " +
            "HAVING SUM(\"MNTB\".\"Quantity\") > 0 ";

            try
            {
                oRecordSet.DoQuery(query);
                int recordCount = oRecordSet.RecordCount;

                NumberFormatInfo Nfi = new NumberFormatInfo() { NumberDecimalSeparator = "." };
                int i = 0;

                //წასაშლელი Goods --->      
                string[] array_HEADER = null;
                string[][] array_GOODS_RS = null;
                string[][] arry_SUB_WAYBILLS = null;

                if ((ID == "0" || ID == "") == false)
                {
                    int get_waybill_result_int = oWayBill.get_waybill(Convert.ToInt32(ID), out array_HEADER, out array_GOODS_RS, out arry_SUB_WAYBILLS, out errorText);
                    if (get_waybill_result_int != 1)
                    {
                        return;
                    }
                    if (array_HEADER != null)
                    {
                        QUANTITYRS = Convert.ToDouble(array_HEADER[44], CultureInfo.InvariantCulture);
                        AMOUNTRS = Convert.ToDouble(array_HEADER[45], CultureInfo.InvariantCulture);
                    }
                }

                int j = 0;
                int countRS = array_GOODS_RS == null ? 0 : array_GOODS_RS.Count();
                array_GOODS = new string[recordCount + countRS][];
                for (j = 0; j < countRS; j++)
                {
                    array_GOODS[j] = new string[13];
                    array_GOODS[j][0] = array_GOODS_RS[j][0]; //ID
                    array_GOODS[j][1] = array_GOODS_RS[j][1]; //W_NAME
                    array_GOODS[j][2] = array_GOODS_RS[j][2]; //UNIT_ID 
                    array_GOODS[j][3] = ""; //ერთეულის სახელი UNIT_TXT
                    array_GOODS[j][4] = array_GOODS_RS[j][3]; //QUANTITY
                    array_GOODS[j][5] = array_GOODS_RS[j][4]; //PRICE
                    array_GOODS[j][6] = "-1"; //STATUS 1 ან -1 თუ გადაეცით -1 შესაბამისი საქონელი წაიშლება
                    array_GOODS[j][7] = array_GOODS_RS[j][5]; //AMOUNT
                    array_GOODS[j][8] = array_GOODS_RS[j][6]; //პროგრამის კოდი
                    array_GOODS[j][9] = array_GOODS_RS[j][7]; //A_ID
                    array_GOODS[j][10] = array_GOODS_RS[j][8]; //VAT_TYPE 0 - ჩეულებრივი; 1 - ნულოვალი; 2 - დაუბეგრავი
                    array_GOODS[j][11] = array_GOODS_RS[j][9]; //QUANTITY_EXT
                }
                //<--- წასაშლელი Goods

                i = j;
                while (!oRecordSet.EoF)
                {
                    array_GOODS[i] = new string[13];
                    array_GOODS[i][0] = "0"; //ID ზედნადებში საქონლის ჩანაწერის ID გადაეცემა 0 თუ ახალი იქმნება
                    array_GOODS[i][1] = oRecordSet.Fields.Item("W_NAME").Value.ToString(); //W_NAME
                    array_GOODS[i][2] = oRecordSet.Fields.Item("UNIT_ID").Value.ToString(); //UNIT_ID 1
                    //array_GOODS[i][2] = array_GOODS[i][2] == null ? "99" : array_GOODS[i][2];
                    string UNIT_TXT = oRecordSet.Fields.Item("UNIT_TXT").Value.ToString();
                    array_GOODS[i][3] = array_GOODS[i][2] == "99" ? (UNIT_TXT == "" ? "სხვა" : UNIT_TXT) : "";//ერთეულის სახელი აუცილებელია როდესაც UNIT_ID=99 („სხვა“)UNIT_TXT
                    array_GOODS[i][4] = oRecordSet.Fields.Item("QUANTITY").Value.ToString(Nfi); //QUANTITY
                    array_GOODS[i][5] = oRecordSet.Fields.Item("PRICE").Value.ToString(Nfi); //PRICE
                    array_GOODS[i][6] = "1"; //STATUS 1 ან -1 თუ გადაეცით -1 შესაბამისი საქონელი წაიშლება
                    array_GOODS[i][7] = oRecordSet.Fields.Item("AMOUNT").Value.ToString(Nfi); //AMOUNT
                    switch (itemCode)
                    {
                        case "0":
                            array_GOODS[i][8] = oRecordSet.Fields.Item("ItemCode").Value.ToString(); //პროგრამის კოდი //BAR_CODE
                            break;
                        case "1":
                            array_GOODS[i][8] = oRecordSet.Fields.Item("AdditionalIdentifier").Value.ToString(); //არტიკული       //BAR_CODE  
                            break;
                        case "2":
                            array_GOODS[i][8] = oRecordSet.Fields.Item("CodeBars").Value.ToString(); //ძირითადი შტრიხკოდი  //BAR_CODE
                            break;
                    }
                    array_GOODS[i][9] = oRecordSet.Fields.Item("A_ID").Value.ToString(); //A_ID თუ აქციზური არ არის გადაეცით 0.
                    string VAT_TYPE = oRecordSet.Fields.Item("VAT_TYPE").Value.ToString(); //VAT_TYPE 0 - ჩეულებრივი; 1 - ნულოვალი; 2 - დაუბეგრავი
                    string VatGroup = oRecordSet.Fields.Item("VatGroup").Value.ToString();
                    VAT_TYPE = VatGroup == "X0" ? "" : VAT_TYPE;
                    switch (VAT_TYPE)
                    {
                        //case "18": VAT_TYPE = "0"; //ჩეულებრივი 18%
                        //    break;
                        case "0":
                            VAT_TYPE = "1"; //ნულოვანი 0%
                            break;
                        case "":
                            VAT_TYPE = "2"; //დაუბეგრავი
                            break;
                        default:
                            VAT_TYPE = "0"; //ჩეულებრივი 18%
                            break;
                    }
                    array_GOODS[i][10] = VAT_TYPE; //VAT_TYPE 0 - ჩეულებრივი; 1 - ნულოვალი; 2 - დაუბეგრავი
                    array_GOODS[i][11] = ""; //QUANTITY_EXT          

                    i = i + 1;
                    oRecordSet.MoveNext();
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                GC.Collect();
            }
        }

        private static void getArrayGoodsDeliveryType(WayBill oWayBill, int baseDocEntry, string ID, out string[][] array_GOODS, out double QUANTITYRS, out double AMOUNTRS, out string errorText)
        {
            errorText = null;
            array_GOODS = null;
            QUANTITYRS = 0;
            AMOUNTRS = 0;

            Dictionary<string, string> rsSettings = CompanyDetails.getRSSettings(out errorText);
            if (errorText != null)
            {
                return;
            }
            string itemCode = rsSettings["ItemCode"];

            Recordset oRecordSet = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            string query = "SELECT " +
            "'" + ID + "'" + " AS \"ID\", " +
            "\"MNTB\".\"LineNum\" AS \"LineNum\", " +
            "\"MNTB\".\"DocEntry\" AS \"DocEntry\", " +
            "\"MNTB\".\"ItemCode\" AS \"ItemCode\", " +
            "\"OITM\".\"CodeBars\" AS \"CodeBars\", " +
            "\"OITM\".\"SWW\" AS \"AdditionalIdentifier\", " +
            "\"MNTB\".\"Dscription\" AS \"W_NAME\", " +
            "CASE WHEN \"BDO_RSUOM\".\"U_RSCode\" is null THEN '99' ELSE \"BDO_RSUOM\".\"U_RSCode\" END AS \"UNIT_ID\", " +
            "CASE WHEN \"MNTB\".\"unitMsr\"='' THEN 'სხვა' ELSE \"MNTB\".\"unitMsr\" END  AS \"UNIT_TXT\", " +
            "\"MNTB\".\"VatPrcnt\" AS \"VAT_TYPE\", " +
            "\"MNTB\".\"VatGroup\"AS \"VatGroup\", " +
            "'0' AS \"A_ID\", " +
            "SUM(\"MNTB\".\"Quantity\") AS \"QUANTITY\", " +
            "SUM(\"MNTB\".\"GTotal\") AS \"AMOUNT\", " +
            "CASE WHEN SUM(\"MNTB\".\"Quantity\") = 0 THEN 0 ELSE SUM(\"MNTB\".\"GTotal\")/SUM(\"MNTB\".\"Quantity\") END AS \"PRICE\", " +
            "SUM(\"MNTB\".\"LineVat\") AS \"LineVat\" " +

            "FROM " +

            "(SELECT " +
            "\"DLN1\".\"DocEntry\", " +
            "\"DLN1\".\"LineNum\", " +
            "\"DLN1\".\"ItemCode\", " +
            "\"DLN1\".\"Dscription\", " +
            "\"DLN1\".\"unitMsr\", " +
            "\"DLN1\".\"Quantity\" * \"DLN1\".\"NumPerMsr\" AS \"Quantity\", " +
            "\"DLN1\".\"GTotal\", " +
            "\"DLN1\".\"VatPrcnt\", " +
            "\"DLN1\".\"VatGroup\", " +
            "\"DLN1\".\"LineVat\" " +

            "FROM \"DLN1\" " +

            "LEFT JOIN \"OITM\" AS \"OITM\" " +
            "ON \"DLN1\".\"ItemCode\" = \"OITM\".\"ItemCode\" " +

            "WHERE \"DLN1\".\"DocEntry\" = '" + baseDocEntry + "' AND (\"OITM\".\"ItemType\" = 'I' AND \"OITM\".\"InvntItem\" = 'Y' OR \"OITM\".\"ItemType\" = 'F' )  " +

            "UNION ALL " +

            "SELECT " +
            "\"RIN1\".\"ActBaseEnt\", " +
            "\"RIN1\".\"ActBaseLn\", " +
            "\"RIN1\".\"ItemCode\", " +
            "\"RIN1\".\"Dscription\", " +
            "\"RIN1\".\"unitMsr\", " +
            "\"RIN1\".\"Quantity\" * (-1) * (CASE WHEN \"RIN1\".\"NoInvtryMv\" = 'Y' THEN 0 ELSE 1 END) * \"RIN1\".\"NumPerMsr\", " +
            "\"RIN1\".\"GTotal\" * (-1), " +
            "\"RIN1\".\"VatPrcnt\", " +
            "\"RIN1\".\"VatGroup\", " +
            "\"RIN1\".\"LineVat\" * (-1) " +

            "FROM \"RIN1\" " +

            "INNER JOIN \"ORIN\" " +
            "ON \"ORIN\".\"DocEntry\" = \"RIN1\".\"DocEntry\" " +

            "LEFT JOIN \"OITM\" AS \"OITM\" " +
            "ON \"RIN1\".\"ItemCode\" = \"OITM\".\"ItemCode\" " +

            "WHERE \"RIN1\".\"ActBaseEnt\" = '" + baseDocEntry + "' AND \"RIN1\".\"TargetType\" < 0  AND \"ORIN\".\"U_BDO_CNTp\" <> 1 AND ((\"OITM\".\"ItemType\" = 'I' AND \"OITM\".\"InvntItem\" = 'Y') OR \"OITM\".\"ItemType\" = 'F' ) ) AS \"MNTB\" " +

            "LEFT JOIN \"OITM\" AS \"OITM\" " +
            "ON \"MNTB\".\"ItemCode\" = \"OITM\".\"ItemCode\" " +

            "LEFT JOIN \"OUOM\" AS \"OUOM\" " +
            "ON \"MNTB\".\"unitMsr\" = \"OUOM\".\"UomName\" " +

            "LEFT JOIN \"@BDO_RSUOM\" AS \"BDO_RSUOM\" " +
            "ON \"OUOM\".\"UomEntry\" = \"BDO_RSUOM\".\"U_UomEntry\" " +

            "GROUP BY " +
            "\"MNTB\".\"DocEntry\", " +
            "\"MNTB\".\"LineNum\", " +
            "\"MNTB\".\"ItemCode\", " +
            "\"MNTB\".\"Dscription\", " +
            "\"OITM\".\"CodeBars\", " +
            "\"OITM\".\"SWW\", " +
            "\"BDO_RSUOM\".\"U_RSCode\", " +
            "\"MNTB\".\"unitMsr\", " +
            "\"MNTB\".\"VatPrcnt\", " +
            "\"MNTB\".\"VatGroup\" " +
            "HAVING SUM(\"MNTB\".\"Quantity\") > 0 ";

            try
            {
                oRecordSet.DoQuery(query);
                int recordCount = oRecordSet.RecordCount;

                NumberFormatInfo Nfi = new NumberFormatInfo() { NumberDecimalSeparator = "." };
                int i = 0;

                //წასაშლელი Goods --->      
                string[] array_HEADER = null;
                string[][] array_GOODS_RS = null;
                string[][] arry_SUB_WAYBILLS = null;

                if ((ID == "0" || ID == "") == false)
                {
                    int get_waybill_result_int = oWayBill.get_waybill(Convert.ToInt32(ID), out array_HEADER, out array_GOODS_RS, out arry_SUB_WAYBILLS, out errorText);
                    if (get_waybill_result_int != 1)
                    {
                        return;
                    }
                    if (array_HEADER != null)
                    {
                        QUANTITYRS = Convert.ToDouble(array_HEADER[44], CultureInfo.InvariantCulture);
                        AMOUNTRS = Convert.ToDouble(array_HEADER[45], CultureInfo.InvariantCulture);
                    }
                }

                int j = 0;
                int countRS = array_GOODS_RS == null ? 0 : array_GOODS_RS.Count();
                array_GOODS = new string[recordCount + countRS][];
                for (j = 0; j < countRS; j++)
                {
                    array_GOODS[j] = new string[13];
                    array_GOODS[j][0] = array_GOODS_RS[j][0]; //ID
                    array_GOODS[j][1] = array_GOODS_RS[j][1]; //W_NAME
                    array_GOODS[j][2] = array_GOODS_RS[j][2]; //UNIT_ID 
                    array_GOODS[j][3] = ""; //ერთეულის სახელი UNIT_TXT
                    array_GOODS[j][4] = array_GOODS_RS[j][3]; //QUANTITY
                    array_GOODS[j][5] = array_GOODS_RS[j][4]; //PRICE
                    array_GOODS[j][6] = "-1"; //STATUS 1 ან -1 თუ გადაეცით -1 შესაბამისი საქონელი წაიშლება
                    array_GOODS[j][7] = array_GOODS_RS[j][5]; //AMOUNT
                    array_GOODS[j][8] = array_GOODS_RS[j][6]; //პროგრამის კოდი
                    array_GOODS[j][9] = array_GOODS_RS[j][7]; //A_ID
                    array_GOODS[j][10] = array_GOODS_RS[j][8]; //VAT_TYPE 0 - ჩეულებრივი; 1 - ნულოვალი; 2 - დაუბეგრავი
                    array_GOODS[j][11] = array_GOODS_RS[j][9]; //QUANTITY_EXT
                }
                //<--- წასაშლელი Goods

                i = j;
                while (!oRecordSet.EoF)
                {
                    array_GOODS[i] = new string[13];
                    array_GOODS[i][0] = "0"; //ID ზედნადებში საქონლის ჩანაწერის ID გადაეცემა 0 თუ ახალი იქმნება
                    array_GOODS[i][1] = oRecordSet.Fields.Item("W_NAME").Value.ToString(); //W_NAME
                    array_GOODS[i][2] = oRecordSet.Fields.Item("UNIT_ID").Value.ToString(); //UNIT_ID 1
                    //array_GOODS[i][2] = array_GOODS[i][2] == null ? "99" : array_GOODS[i][2];
                    string UNIT_TXT = oRecordSet.Fields.Item("UNIT_TXT").Value.ToString();
                    array_GOODS[i][3] = array_GOODS[i][2] == "99" ? (UNIT_TXT == "" ? "სხვა" : UNIT_TXT) : "";//ერთეულის სახელი აუცილებელია როდესაც UNIT_ID=99 („სხვა“)UNIT_TXT
                    array_GOODS[i][4] = oRecordSet.Fields.Item("QUANTITY").Value.ToString(Nfi); //QUANTITY
                    array_GOODS[i][5] = oRecordSet.Fields.Item("PRICE").Value.ToString(Nfi); //PRICE
                    array_GOODS[i][6] = "1"; //STATUS 1 ან -1 თუ გადაეცით -1 შესაბამისი საქონელი წაიშლება
                    array_GOODS[i][7] = oRecordSet.Fields.Item("AMOUNT").Value.ToString(Nfi); //AMOUNT
                    switch (itemCode)
                    {
                        case "0":
                            array_GOODS[i][8] = oRecordSet.Fields.Item("ItemCode").Value.ToString(); //პროგრამის კოდი //BAR_CODE
                            break;
                        case "1":
                            array_GOODS[i][8] = oRecordSet.Fields.Item("AdditionalIdentifier").Value.ToString(); //არტიკული       //BAR_CODE  
                            break;
                        case "2":
                            array_GOODS[i][8] = oRecordSet.Fields.Item("CodeBars").Value.ToString(); //ძირითადი შტრიხკოდი  //BAR_CODE
                            break;
                    }
                    array_GOODS[i][9] = oRecordSet.Fields.Item("A_ID").Value.ToString(); //A_ID თუ აქციზური არ არის გადაეცით 0.
                    string VAT_TYPE = oRecordSet.Fields.Item("VAT_TYPE").Value.ToString(); //VAT_TYPE 0 - ჩეულებრივი; 1 - ნულოვალი; 2 - დაუბეგრავი
                    string VatGroup = oRecordSet.Fields.Item("VatGroup").Value.ToString();
                    VAT_TYPE = VatGroup == "X0" ? "" : VAT_TYPE;
                    switch (VAT_TYPE)
                    {
                        //case "18": VAT_TYPE = "0"; //ჩეულებრივი 18%
                        //    break;
                        case "0":
                            VAT_TYPE = "1"; //ნულოვანი 0%
                            break;
                        case "":
                            VAT_TYPE = "2"; //დაუბეგრავი
                            break;
                        default:
                            VAT_TYPE = "0"; //ჩეულებრივი 18%
                            break;
                    }
                    array_GOODS[i][10] = VAT_TYPE; //VAT_TYPE 0 - ჩეულებრივი; 1 - ნულოვალი; 2 - დაუბეგრავი
                    array_GOODS[i][11] = ""; //QUANTITY_EXT          

                    i = i + 1;
                    oRecordSet.MoveNext();
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                GC.Collect();
            }
        }

        private static void getArrayGoodsInvoiceCreditMemoType(WayBill oWayBill, int baseDocEntry, string ID, out string[][] array_GOODS, out double QUANTITYRS, out double AMOUNTRS, out string errorText)
        {
            errorText = null;
            array_GOODS = null;
            QUANTITYRS = 0;
            AMOUNTRS = 0;

            Dictionary<string, string> rsSettings = CompanyDetails.getRSSettings(out errorText);
            if (errorText != null)
            {
                return;
            }
            string itemCode = rsSettings["ItemCode"];

            Recordset oRecordSet = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            string query = "SELECT " +
            "'" + ID + "'" + " AS \"ID\", " +
            "\"MNTB\".\"LineNum\" AS \"LineNum\", " +
            "\"MNTB\".\"DocEntry\" AS \"DocEntry\", " +
            "\"MNTB\".\"ItemCode\" AS \"ItemCode\", " +
            "\"OITM\".\"CodeBars\" AS \"CodeBars\", " +
            "\"OITM\".\"SWW\" AS \"AdditionalIdentifier\", " +
            "\"MNTB\".\"Dscription\" AS \"W_NAME\", " +
            "CASE WHEN \"BDO_RSUOM\".\"U_RSCode\" is null THEN '99' ELSE \"BDO_RSUOM\".\"U_RSCode\" END AS \"UNIT_ID\", " +
            "CASE WHEN \"MNTB\".\"unitMsr\"='' THEN 'სხვა' ELSE \"MNTB\".\"unitMsr\" END  AS \"UNIT_TXT\", " +
            "\"MNTB\".\"VatPrcnt\" AS \"VAT_TYPE\", " +
            "\"MNTB\".\"VatGroup\"AS \"VatGroup\", " +
            "'0' AS \"A_ID\", " +
            "SUM(\"MNTB\".\"Quantity\") AS \"QUANTITY\", " +
            "SUM(\"MNTB\".\"GTotal\") AS \"AMOUNT\", " +
            "CASE WHEN SUM(\"MNTB\".\"Quantity\") = 0 THEN 0 ELSE SUM(\"MNTB\".\"GTotal\")/SUM(\"MNTB\".\"Quantity\") END AS \"PRICE\", " +
            "SUM(\"MNTB\".\"LineVat\") AS \"LineVat\" " +

            "FROM " +

            "(SELECT " +
            "\"RIN1\".\"DocEntry\", " +
            "\"RIN1\".\"LineNum\", " +
            "\"RIN1\".\"ItemCode\", " +
            "\"RIN1\".\"Dscription\", " +
            "\"RIN1\".\"unitMsr\", " +
            "\"RIN1\".\"Quantity\" * (CASE WHEN \"RIN1\".\"NoInvtryMv\" = 'Y' THEN 0 ELSE 1 END) * \"RIN1\".\"NumPerMsr\" AS \"Quantity\", " +
            "\"RIN1\".\"GTotal\" , " +
            "\"RIN1\".\"VatPrcnt\", " +
            "\"RIN1\".\"VatGroup\", " +
            "\"RIN1\".\"LineVat\" " +

            "FROM \"RIN1\" " +

            "INNER JOIN \"ORIN\" " +
            "ON \"ORIN\".\"DocEntry\" = \"RIN1\".\"DocEntry\" " +

            "LEFT JOIN \"OITM\" AS \"OITM\" " +
            "ON \"RIN1\".\"ItemCode\" = \"OITM\".\"ItemCode\" " +

            "WHERE \"RIN1\".\"DocEntry\" = '" + baseDocEntry + "' AND \"RIN1\".\"TargetType\" < 0  AND \"ORIN\".\"U_BDO_CNTp\" = 1 AND ((\"OITM\".\"ItemType\" = 'I' AND \"OITM\".\"InvntItem\" = 'Y') OR \"OITM\".\"ItemType\" = 'F' ) ) AS \"MNTB\" " +

            "LEFT JOIN \"OITM\" AS \"OITM\" " +
            "ON \"MNTB\".\"ItemCode\" = \"OITM\".\"ItemCode\" " +

            "LEFT JOIN \"OUOM\" AS \"OUOM\" " +
            "ON \"MNTB\".\"unitMsr\" = \"OUOM\".\"UomName\" " +

            "LEFT JOIN \"@BDO_RSUOM\" AS \"BDO_RSUOM\" " +
            "ON \"OUOM\".\"UomEntry\" = \"BDO_RSUOM\".\"U_UomEntry\" " +

            "GROUP BY " +
            "\"MNTB\".\"DocEntry\", " +
            "\"MNTB\".\"LineNum\", " +
            "\"MNTB\".\"ItemCode\", " +
            "\"MNTB\".\"Dscription\", " +
            "\"OITM\".\"CodeBars\", " +
            "\"OITM\".\"SWW\", " +
            "\"BDO_RSUOM\".\"U_RSCode\", " +
            "\"MNTB\".\"unitMsr\", " +
            "\"MNTB\".\"VatPrcnt\", " +
            "\"MNTB\".\"VatGroup\" " +
            "HAVING SUM(\"MNTB\".\"Quantity\") > 0 ";

            try
            {
                oRecordSet.DoQuery(query);
                int recordCount = oRecordSet.RecordCount;

                NumberFormatInfo Nfi = new NumberFormatInfo() { NumberDecimalSeparator = "." };
                int i = 0;

                //წასაშლელი Goods --->      
                string[] array_HEADER = null;
                string[][] array_GOODS_RS = null;
                string[][] arry_SUB_WAYBILLS = null;

                if ((ID == "0" || ID == "") == false)
                {
                    int get_waybill_result_int = oWayBill.get_waybill(Convert.ToInt32(ID), out array_HEADER, out array_GOODS_RS, out arry_SUB_WAYBILLS, out errorText);
                    if (get_waybill_result_int != 1)
                    {
                        return;
                    }
                    if (array_HEADER != null)
                    {
                        QUANTITYRS = Convert.ToDouble(array_HEADER[44], CultureInfo.InvariantCulture);
                        AMOUNTRS = Convert.ToDouble(array_HEADER[45], CultureInfo.InvariantCulture);
                    }
                }

                int j = 0;
                int countRS = array_GOODS_RS == null ? 0 : array_GOODS_RS.Count();
                array_GOODS = new string[recordCount + countRS][];
                for (j = 0; j < countRS; j++)
                {
                    array_GOODS[j] = new string[13];
                    array_GOODS[j][0] = array_GOODS_RS[j][0]; //ID
                    array_GOODS[j][1] = array_GOODS_RS[j][1]; //W_NAME
                    array_GOODS[j][2] = array_GOODS_RS[j][2]; //UNIT_ID 
                    array_GOODS[j][3] = ""; //ერთეულის სახელი UNIT_TXT
                    array_GOODS[j][4] = array_GOODS_RS[j][3]; //QUANTITY
                    array_GOODS[j][5] = array_GOODS_RS[j][4]; //PRICE
                    array_GOODS[j][6] = "-1"; //STATUS 1 ან -1 თუ გადაეცით -1 შესაბამისი საქონელი წაიშლება
                    array_GOODS[j][7] = array_GOODS_RS[j][5]; //AMOUNT
                    array_GOODS[j][8] = array_GOODS_RS[j][6]; //პროგრამის კოდი
                    array_GOODS[j][9] = array_GOODS_RS[j][7]; //A_ID
                    array_GOODS[j][10] = array_GOODS_RS[j][8]; //VAT_TYPE 0 - ჩეულებრივი; 1 - ნულოვალი; 2 - დაუბეგრავი
                    array_GOODS[j][11] = array_GOODS_RS[j][9]; //QUANTITY_EXT
                }
                //<--- წასაშლელი Goods

                i = j;
                while (!oRecordSet.EoF)
                {
                    array_GOODS[i] = new string[13];
                    array_GOODS[i][0] = "0"; //ID ზედნადებში საქონლის ჩანაწერის ID გადაეცემა 0 თუ ახალი იქმნება
                    array_GOODS[i][1] = oRecordSet.Fields.Item("W_NAME").Value.ToString(); //W_NAME
                    array_GOODS[i][2] = oRecordSet.Fields.Item("UNIT_ID").Value.ToString(); //UNIT_ID 1
                    string UNIT_TXT = oRecordSet.Fields.Item("UNIT_TXT").Value.ToString();
                    array_GOODS[i][3] = array_GOODS[i][2] == "99" ? (UNIT_TXT == "" ? "სხვა" : UNIT_TXT) : "";//ერთეულის სახელი აუცილებელია როდესაც UNIT_ID=99 („სხვა“)UNIT_TXT
                    array_GOODS[i][4] = oRecordSet.Fields.Item("QUANTITY").Value.ToString(Nfi); //QUANTITY
                    array_GOODS[i][5] = oRecordSet.Fields.Item("PRICE").Value.ToString(Nfi); //PRICE
                    array_GOODS[i][6] = "1"; //STATUS 1 ან -1 თუ გადაეცით -1 შესაბამისი საქონელი წაიშლება
                    array_GOODS[i][7] = oRecordSet.Fields.Item("AMOUNT").Value.ToString(Nfi); //AMOUNT
                    switch (itemCode)
                    {
                        case "0":
                            array_GOODS[i][8] = oRecordSet.Fields.Item("ItemCode").Value.ToString(); //პროგრამის კოდი //BAR_CODE
                            break;
                        case "1":
                            array_GOODS[i][8] = oRecordSet.Fields.Item("AdditionalIdentifier").Value.ToString(); //არტიკული       //BAR_CODE  
                            break;
                        case "2":
                            array_GOODS[i][8] = oRecordSet.Fields.Item("CodeBars").Value.ToString(); //ძირითადი შტრიხკოდი  //BAR_CODE
                            break;
                    }
                    array_GOODS[i][9] = oRecordSet.Fields.Item("A_ID").Value.ToString(); //A_ID თუ აქციზური არ არის გადაეცით 0.
                    string VAT_TYPE = oRecordSet.Fields.Item("VAT_TYPE").Value.ToString(); //VAT_TYPE 0 - ჩეულებრივი; 1 - ნულოვალი; 2 - დაუბეგრავი
                    string VatGroup = oRecordSet.Fields.Item("VatGroup").Value.ToString();
                    VAT_TYPE = VatGroup == "X0" ? "" : VAT_TYPE;
                    switch (VAT_TYPE)
                    {
                        //case "18": VAT_TYPE = "0"; //ჩვეულებრივი 18%
                        //    break;
                        case "0":
                            VAT_TYPE = "1"; //ნულოვანი 0%
                            break;
                        case "":
                            VAT_TYPE = "2"; //დაუბეგრავი
                            break;
                        default:
                            VAT_TYPE = "0"; //ჩეულებრივი 18%
                            break;
                    }
                    array_GOODS[i][10] = VAT_TYPE; //VAT_TYPE 0 - ჩეულებრივი; 1 - ნულოვალი; 2 - დაუბეგრავი
                    array_GOODS[i][11] = ""; //QUANTITY_EXT          

                    i = i + 1;
                    oRecordSet.MoveNext();
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                GC.Collect();
            }
        }

        private static void getArrayGoodsInvoiceCorrectionType(WayBill oWayBill, int baseDocEntry, string ID, out string[][] array_GOODS, out double QUANTITYRS, out double AMOUNTRS, out string errorText)
        {
            array_GOODS = null;
            QUANTITYRS = 0;
            AMOUNTRS = 0;

            Dictionary<string, string> rsSettings = CompanyDetails.getRSSettings(out errorText);
            if (errorText != null)
            {
                return;
            }
            string itemCode = rsSettings["ItemCode"];

            Recordset oRecordSet = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            string query = "SELECT " +
            "'" + ID + "'" + " AS \"ID\", " +
            "\"MNTB\".\"DocEntry\" AS \"DocEntry\", " +
            "\"MNTB\".\"ItemCode\" AS \"ItemCode\", " +
            "\"OITM\".\"CodeBars\" AS \"CodeBars\", " +
            "\"OITM\".\"SWW\" AS \"AdditionalIdentifier\", " +
            "\"MNTB\".\"Dscription\" AS \"W_NAME\", " +
            "CASE WHEN \"BDO_RSUOM\".\"U_RSCode\" is null THEN '99' ELSE \"BDO_RSUOM\".\"U_RSCode\" END AS \"UNIT_ID\", " +
            "CASE WHEN \"MNTB\".\"unitMsr\"='' THEN 'სხვა' ELSE \"MNTB\".\"unitMsr\" END  AS \"UNIT_TXT\", " +
            "\"MNTB\".\"VatPrcnt\" AS \"VAT_TYPE\", " +
            "\"MNTB\".\"VatGroup\"AS \"VatGroup\", " +
            "'0' AS \"A_ID\", " +
            "SUM(\"MNTB\".\"Quantity\") AS \"QUANTITY\", " +
            "SUM(\"MNTB\".\"GTotal\") AS \"AMOUNT\", " +
            "CASE WHEN SUM(\"MNTB\".\"Quantity\") = 0 THEN 0 ELSE SUM(\"MNTB\".\"GTotal\")/SUM(\"MNTB\".\"Quantity\") END AS \"PRICE\", " +
            "SUM(\"MNTB\".\"LineVat\") AS \"LineVat\" " +

            "FROM " +

            "(SELECT " +
            "\"CSI1\".\"DocEntry\", " +
            "\"CSI1\".\"ItemCode\", " +
            "\"CSI1\".\"Dscription\", " +
            "\"CSI1\".\"unitMsr\", " +
            "ABS(SUM(\"CSI1\".\"Quantity\" * (CASE WHEN \"CSI1\".\"NoInvtryMv\" = 'Y' THEN 0 ELSE 1 END) * \"CSI1\".\"NumPerMsr\")) AS \"Quantity\", " +
            "ABS(SUM(\"CSI1\".\"GTotal\")) AS \"GTotal\", " +
            "\"CSI1\".\"VatPrcnt\", " +
            "\"CSI1\".\"VatGroup\", " +
            "ABS(SUM(\"CSI1\".\"LineVat\")) AS \"LineVat\"" +

            "FROM \"CSI1\" " +

            "INNER JOIN \"OCSI\" " +
            "ON \"OCSI\".\"DocEntry\" = \"CSI1\".\"DocEntry\" " +

            "LEFT JOIN \"OITM\" AS \"OITM\" " +
            "ON \"CSI1\".\"ItemCode\" = \"OITM\".\"ItemCode\" " +

            "WHERE \"CSI1\".\"DocEntry\" = '" + baseDocEntry + "' AND \"CSI1\".\"TargetType\" < 0  AND \"OCSI\".\"U_BDOSCITp\" = 1 AND ((\"OITM\".\"ItemType\" = 'I' AND \"OITM\".\"InvntItem\" = 'Y') OR \"OITM\".\"ItemType\" = 'F' ) group by \"CSI1\".\"DocEntry\",\"CSI1\".\"Dscription\",\"CSI1\".\"unitMsr\",\"CSI1\".\"VatPrcnt\", \"CSI1\".\"VatGroup\",\"CSI1\".\"ItemCode\") AS \"MNTB\" " +

            "LEFT JOIN \"OITM\" AS \"OITM\" " +
            "ON \"MNTB\".\"ItemCode\" = \"OITM\".\"ItemCode\" " +

            "LEFT JOIN \"OUOM\" AS \"OUOM\" " +
            "ON \"MNTB\".\"unitMsr\" = \"OUOM\".\"UomName\" " +

            "LEFT JOIN \"@BDO_RSUOM\" AS \"BDO_RSUOM\" " +
            "ON \"OUOM\".\"UomEntry\" = \"BDO_RSUOM\".\"U_UomEntry\" " +

            "GROUP BY " +
            "\"MNTB\".\"DocEntry\", " +
            "\"MNTB\".\"ItemCode\", " +
            "\"MNTB\".\"Dscription\", " +
            "\"OITM\".\"CodeBars\", " +
            "\"OITM\".\"SWW\", " +
            "\"BDO_RSUOM\".\"U_RSCode\", " +
            "\"MNTB\".\"unitMsr\", " +
            "\"MNTB\".\"VatPrcnt\", " +
            "\"MNTB\".\"VatGroup\" " +
            "HAVING SUM(\"MNTB\".\"Quantity\") > 0 ";

            try
            {
                oRecordSet.DoQuery(query);
                int recordCount = oRecordSet.RecordCount;

                NumberFormatInfo Nfi = new NumberFormatInfo() { NumberDecimalSeparator = "." };
                int i = 0;

                //წასაშლელი Goods --->      
                string[] array_HEADER = null;
                string[][] array_GOODS_RS = null;
                string[][] arry_SUB_WAYBILLS = null;

                if ((ID == "0" || ID == "") == false)
                {
                    int get_waybill_result_int = oWayBill.get_waybill(Convert.ToInt32(ID), out array_HEADER, out array_GOODS_RS, out arry_SUB_WAYBILLS, out errorText);
                    if (get_waybill_result_int != 1)
                    {
                        return;
                    }
                    if (array_HEADER != null)
                    {
                        QUANTITYRS = Convert.ToDouble(array_HEADER[44], CultureInfo.InvariantCulture);
                        AMOUNTRS = Convert.ToDouble(array_HEADER[45], CultureInfo.InvariantCulture);
                    }
                }

                int j = 0;
                int countRS = array_GOODS_RS == null ? 0 : array_GOODS_RS.Count();
                array_GOODS = new string[recordCount + countRS][];
                for (j = 0; j < countRS; j++)
                {
                    array_GOODS[j] = new string[13];
                    array_GOODS[j][0] = array_GOODS_RS[j][0]; //ID
                    array_GOODS[j][1] = array_GOODS_RS[j][1]; //W_NAME
                    array_GOODS[j][2] = array_GOODS_RS[j][2]; //UNIT_ID 
                    array_GOODS[j][3] = ""; //ერთეულის სახელი UNIT_TXT
                    array_GOODS[j][4] = array_GOODS_RS[j][3]; //QUANTITY
                    array_GOODS[j][5] = array_GOODS_RS[j][4]; //PRICE
                    array_GOODS[j][6] = "-1"; //STATUS 1 ან -1 თუ გადაეცით -1 შესაბამისი საქონელი წაიშლება
                    array_GOODS[j][7] = array_GOODS_RS[j][5]; //AMOUNT
                    array_GOODS[j][8] = array_GOODS_RS[j][6]; //პროგრამის კოდი
                    array_GOODS[j][9] = array_GOODS_RS[j][7]; //A_ID
                    array_GOODS[j][10] = array_GOODS_RS[j][8]; //VAT_TYPE 0 - ჩეულებრივი; 1 - ნულოვალი; 2 - დაუბეგრავი
                    array_GOODS[j][11] = array_GOODS_RS[j][9]; //QUANTITY_EXT
                }
                //<--- წასაშლელი Goods

                i = j;
                while (!oRecordSet.EoF)
                {
                    array_GOODS[i] = new string[13];
                    array_GOODS[i][0] = "0"; //ID ზედნადებში საქონლის ჩანაწერის ID გადაეცემა 0 თუ ახალი იქმნება
                    array_GOODS[i][1] = oRecordSet.Fields.Item("W_NAME").Value.ToString(); //W_NAME
                    array_GOODS[i][2] = oRecordSet.Fields.Item("UNIT_ID").Value.ToString(); //UNIT_ID 1
                    string UNIT_TXT = oRecordSet.Fields.Item("UNIT_TXT").Value.ToString();
                    array_GOODS[i][3] = array_GOODS[i][2] == "99" ? (UNIT_TXT == "" ? "სხვა" : UNIT_TXT) : "";//ერთეულის სახელი აუცილებელია როდესაც UNIT_ID=99 („სხვა“)UNIT_TXT
                    array_GOODS[i][4] = oRecordSet.Fields.Item("QUANTITY").Value.ToString(Nfi); //QUANTITY
                    array_GOODS[i][5] = oRecordSet.Fields.Item("PRICE").Value.ToString(Nfi); //PRICE
                    array_GOODS[i][6] = "1"; //STATUS 1 ან -1 თუ გადაეცით -1 შესაბამისი საქონელი წაიშლება
                    array_GOODS[i][7] = oRecordSet.Fields.Item("AMOUNT").Value.ToString(Nfi); //AMOUNT
                    switch (itemCode)
                    {
                        case "0":
                            array_GOODS[i][8] = oRecordSet.Fields.Item("ItemCode").Value.ToString(); //პროგრამის კოდი //BAR_CODE
                            break;
                        case "1":
                            array_GOODS[i][8] = oRecordSet.Fields.Item("AdditionalIdentifier").Value.ToString(); //არტიკული       //BAR_CODE  
                            break;
                        case "2":
                            array_GOODS[i][8] = oRecordSet.Fields.Item("CodeBars").Value.ToString(); //ძირითადი შტრიხკოდი  //BAR_CODE
                            break;
                    }
                    array_GOODS[i][9] = oRecordSet.Fields.Item("A_ID").Value.ToString(); //A_ID თუ აქციზური არ არის გადაეცით 0.
                    string VAT_TYPE = oRecordSet.Fields.Item("VAT_TYPE").Value.ToString(); //VAT_TYPE 0 - ჩეულებრივი; 1 - ნულოვალი; 2 - დაუბეგრავი
                    string VatGroup = oRecordSet.Fields.Item("VatGroup").Value.ToString();
                    VAT_TYPE = VatGroup == "X0" ? "" : VAT_TYPE;
                    switch (VAT_TYPE)
                    {
                        //case "18": VAT_TYPE = "0"; //ჩვეულებრივი 18%
                        //    break;
                        case "0":
                            VAT_TYPE = "1"; //ნულოვანი 0%
                            break;
                        case "":
                            VAT_TYPE = "2"; //დაუბეგრავი
                            break;
                        default:
                            VAT_TYPE = "0"; //ჩეულებრივი 18%
                            break;
                    }
                    array_GOODS[i][10] = VAT_TYPE; //VAT_TYPE 0 - ჩეულებრივი; 1 - ნულოვალი; 2 - დაუბეგრავი
                    array_GOODS[i][11] = ""; //QUANTITY_EXT          

                    i = i + 1;
                    oRecordSet.MoveNext();
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                GC.Collect();
            }
        }

        private static void getArrayGoodsInventoryTransferType(WayBill oWayBill, int baseDocEntry, string ID, out string[][] array_GOODS, out double QUANTITYRS, out string errorText)
        {
            errorText = null;
            array_GOODS = null;
            QUANTITYRS = 0;

            Dictionary<string, string> rsSettings = CompanyDetails.getRSSettings(out errorText);
            if (errorText != null)
            {
                return;
            }
            string itemCode = rsSettings["ItemCode"];

            Recordset oRecordSet = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            string query = "SELECT " +
            "\"WTR1\".\"LineNum\", " +
            "'" + ID + "'" + " AS \"ID\", " +
            "\"WTR1\".\"DocEntry\" AS \"DocEntry\", " +
            "\"WTR1\".\"ItemCode\" AS \"ItemCode\", " +
            "\"OITM\".\"CodeBars\" AS \"CodeBars\", " +
            "\"OITM\".\"SWW\" AS \"AdditionalIdentifier\", " +
            "\"WTR1\".\"Dscription\" AS \"W_NAME\", " +
            "CASE WHEN \"BDO_RSUOM\".\"U_RSCode\" is null THEN '99' ELSE \"BDO_RSUOM\".\"U_RSCode\" END AS \"UNIT_ID\", " +
            "CASE WHEN \"WTR1\".\"unitMsr\"='' THEN 'სხვა' ELSE \"WTR1\".\"unitMsr\" END  AS \"UNIT_TXT\", " +
            "\"WTR1\".\"VatPrcnt\" AS \"VAT_TYPE\", " +
            "\"WTR1\".\"VatGroup\"AS \"VatGroup\", " +
            "'0' AS \"A_ID\", " +
            "\"WTR1\".\"Quantity\" * \"WTR1\".\"NumPerMsr\" AS \"Quantity\", " +
            "'0' AS \"AMOUNT\"," +
            "'0' AS \"PRICE\" " +

            "FROM \"WTR1\" AS \"WTR1\" " +

            "LEFT JOIN \"OITM\" AS \"OITM\" " +
            "ON \"WTR1\".\"ItemCode\" = \"OITM\".\"ItemCode\" " +

            "LEFT JOIN \"OUOM\" AS \"OUOM\" " +
            "ON \"WTR1\".\"unitMsr\" = \"OUOM\".\"UomName\" " +

            "LEFT JOIN \"@BDO_RSUOM\" AS \"BDO_RSUOM\" " +
            "ON \"OUOM\".\"UomEntry\" = \"BDO_RSUOM\".\"U_UomEntry\" " +

            "WHERE \"WTR1\".\"DocEntry\" = '" + baseDocEntry + "' AND ((\"OITM\".\"ItemType\" = 'I' AND \"OITM\".\"InvntItem\" = 'Y') OR \"OITM\".\"ItemType\" = 'F' ) ";

            try
            {
                oRecordSet.DoQuery(query);
                int recordCount = oRecordSet.RecordCount;

                NumberFormatInfo Nfi = new NumberFormatInfo() { NumberDecimalSeparator = "." };
                int i = 0;

                //წასაშლელი Goods --->      
                string[] array_HEADER = null;
                string[][] array_GOODS_RS = null;
                string[][] arry_SUB_WAYBILLS = null;

                if ((ID == "0" || ID == "") == false)
                {
                    int get_waybill_result_int = oWayBill.get_waybill(Convert.ToInt32(ID), out array_HEADER, out array_GOODS_RS, out arry_SUB_WAYBILLS, out errorText);
                    if (get_waybill_result_int != 1)
                    {
                        return;
                    }
                    if (array_HEADER != null)
                    {
                        QUANTITYRS = Convert.ToDouble(array_HEADER[44], CultureInfo.InvariantCulture);
                    }
                }

                int j = 0;
                int countRS = array_GOODS_RS == null ? 0 : array_GOODS_RS.Count();
                array_GOODS = new string[recordCount + countRS][];
                for (j = 0; j < countRS; j++)
                {
                    array_GOODS[j] = new string[13];
                    array_GOODS[j][0] = array_GOODS_RS[j][0]; //ID
                    array_GOODS[j][1] = array_GOODS_RS[j][1]; //W_NAME
                    array_GOODS[j][2] = array_GOODS_RS[j][2]; //UNIT_ID 
                    array_GOODS[j][3] = ""; //ერთეულის სახელი UNIT_TXT
                    array_GOODS[j][4] = array_GOODS_RS[j][3]; //QUANTITY
                    array_GOODS[j][5] = array_GOODS_RS[j][4]; //PRICE
                    array_GOODS[j][6] = "-1"; //STATUS 1 ან -1 თუ გადაეცით -1 შესაბამისი საქონელი წაიშლება
                    array_GOODS[j][7] = array_GOODS_RS[j][5]; //AMOUNT
                    array_GOODS[j][8] = array_GOODS_RS[j][6]; //პროგრამის კოდი
                    array_GOODS[j][9] = array_GOODS_RS[j][7]; //A_ID
                    array_GOODS[j][10] = array_GOODS_RS[j][8]; //VAT_TYPE 0 - ჩეულებრივი; 1 - ნულოვალი; 2 - დაუბეგრავი
                    array_GOODS[j][11] = array_GOODS_RS[j][9]; //QUANTITY_EXT
                }
                //<--- წასაშლელი Goods

                i = j;
                while (!oRecordSet.EoF)
                {
                    array_GOODS[i] = new string[13];
                    array_GOODS[i][0] = "0"; //ID ზედნადებში საქონლის ჩანაწერის ID გადაეცემა 0 თუ ახალი იქმნება
                    array_GOODS[i][1] = oRecordSet.Fields.Item("W_NAME").Value.ToString(); //W_NAME
                    array_GOODS[i][2] = oRecordSet.Fields.Item("UNIT_ID").Value.ToString(); //UNIT_ID 1
                    string UNIT_TXT = oRecordSet.Fields.Item("UNIT_TXT").Value.ToString();
                    array_GOODS[i][3] = array_GOODS[i][2] == "99" ? (UNIT_TXT == "" ? "სხვა" : UNIT_TXT) : "";//ერთეულის სახელი აუცილებელია როდესაც UNIT_ID=99 („სხვა“)UNIT_TXT
                    array_GOODS[i][4] = oRecordSet.Fields.Item("QUANTITY").Value.ToString(Nfi); //QUANTITY
                    array_GOODS[i][5] = oRecordSet.Fields.Item("PRICE").Value.ToString(Nfi); //PRICE
                    array_GOODS[i][6] = "1"; //STATUS 1 ან -1 თუ გადაეცით -1 შესაბამისი საქონელი წაიშლება
                    array_GOODS[i][7] = oRecordSet.Fields.Item("AMOUNT").Value.ToString(Nfi); //AMOUNT
                    switch (itemCode)
                    {
                        case "0":
                            array_GOODS[i][8] = oRecordSet.Fields.Item("ItemCode").Value.ToString(); //პროგრამის კოდი //BAR_CODE
                            break;
                        case "1":
                            array_GOODS[i][8] = oRecordSet.Fields.Item("AdditionalIdentifier").Value.ToString(); //არტიკული       //BAR_CODE  
                            break;
                        case "2":
                            array_GOODS[i][8] = oRecordSet.Fields.Item("CodeBars").Value.ToString(); //ძირითადი შტრიხკოდი  //BAR_CODE
                            break;
                    }
                    array_GOODS[i][9] = oRecordSet.Fields.Item("A_ID").Value.ToString(); //A_ID თუ აქციზური არ არის გადაეცით 0.
                    string VAT_TYPE = oRecordSet.Fields.Item("VAT_TYPE").Value.ToString(); //VAT_TYPE 0 - ჩეულებრივი; 1 - ნულოვალი; 2 - დაუბეგრავი
                    string VatGroup = oRecordSet.Fields.Item("VatGroup").Value.ToString();
                    VAT_TYPE = VatGroup == "X0" ? "" : VAT_TYPE;
                    switch (VAT_TYPE)
                    {
                        // case "18": VAT_TYPE = "0"; //ჩეულებრივი 18%
                        //     break;
                        // case "0":
                        //     VAT_TYPE = "1"; //ნულოვანი 0%
                        //     break;
                        case "":
                            VAT_TYPE = "2"; //დაუბეგრავი
                            break;
                        default:
                            VAT_TYPE = "0"; //ჩეულებრივი 18%
                            break;
                    }
                    array_GOODS[i][10] = VAT_TYPE; //VAT_TYPE 0 - ჩეულებრივი; 1 - ნულოვალი; 2 - დაუბეგრავი
                    array_GOODS[i][11] = ""; //QUANTITY_EXT          

                    i = i + 1;
                    oRecordSet.MoveNext();
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                GC.Collect();
            }
        }

        private static void getArrayGoodsFixedAssetTransferType(WayBill oWayBill, int baseDocEntry, string ID, out string[][] array_GOODS, out double QUANTITYRS, out string errorText)
        {
            errorText = null;
            array_GOODS = null;
            QUANTITYRS = 0;

            Dictionary<string, string> rsSettings = CompanyDetails.getRSSettings(out errorText);
            if (errorText != null)
            {
                return;
            }
            string itemCode = rsSettings["ItemCode"];

            Recordset oRecordSet = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            string query = "SELECT " +
            "\"@BDOSFASTR1\".\"LineId\", " +
            "'" + ID + "'" + " AS \"ID\", " +
            "\"@BDOSFASTR1\".\"DocEntry\" AS \"DocEntry\", " +
            "\"@BDOSFASTR1\".\"U_ItemCode\" AS \"ItemCode\", " +
            "\"OITM\".\"CodeBars\" AS \"CodeBars\", " +
            "\"OITM\".\"SWW\" AS \"AdditionalIdentifier\", " +
            "\"@BDOSFASTR1\".\"U_ItemName\" AS \"W_NAME\", " +
            "CASE WHEN \"BDO_RSUOM\".\"U_RSCode\" is null THEN '99' ELSE \"BDO_RSUOM\".\"U_RSCode\" END AS \"UNIT_ID\", " +
            "CASE WHEN \"OITM\".\"InvntryUom\"='' THEN 'სხვა' ELSE \"OITM\".\"InvntryUom\" END  AS \"UNIT_TXT\", " +
            "0 AS \"VAT_TYPE\", " +
            " '' AS \"VatGroup\", " +
            "'0' AS \"A_ID\", " +
            "CASE WHEN \"@BDOSFASTR1\".\"U_Quantity\" is null or \"@BDOSFASTR1\".\"U_Quantity\" = '0' THEN '1' ELSE \"@BDOSFASTR1\".\"U_Quantity\" END AS \"Quantity\", " +
            "'0' AS \"AMOUNT\"," +
            "'0' AS \"PRICE\" " +

            "FROM \"@BDOSFASTR1\" AS \"@BDOSFASTR1\" " +

            "LEFT JOIN \"OITM\" AS \"OITM\" " +
            "ON \"@BDOSFASTR1\".\"U_ItemCode\" = \"OITM\".\"ItemCode\" " +

            "LEFT JOIN \"OUOM\" AS \"OUOM\" " +
            "ON \"OITM\".\"InvntryUom\" = \"OUOM\".\"UomName\" " +

            "LEFT JOIN \"@BDO_RSUOM\" AS \"BDO_RSUOM\" " +
            "ON \"OUOM\".\"UomEntry\" = \"BDO_RSUOM\".\"U_UomEntry\" " +

            "WHERE \"@BDOSFASTR1\".\"DocEntry\" = '" + baseDocEntry + "' AND ((\"OITM\".\"ItemType\" = 'I' AND \"OITM\".\"InvntItem\" = 'Y') OR \"OITM\".\"ItemType\" = 'F' ) ";

            try
            {
                oRecordSet.DoQuery(query);
                int recordCount = oRecordSet.RecordCount;

                NumberFormatInfo Nfi = new NumberFormatInfo() { NumberDecimalSeparator = "." };
                int i = 0;

                //წასაშლელი Goods --->      
                string[] array_HEADER = null;
                string[][] array_GOODS_RS = null;
                string[][] arry_SUB_WAYBILLS = null;

                if ((ID == "0" || ID == "") == false)
                {
                    int get_waybill_result_int = oWayBill.get_waybill(Convert.ToInt32(ID), out array_HEADER, out array_GOODS_RS, out arry_SUB_WAYBILLS, out errorText);
                    if (get_waybill_result_int != 1)
                    {
                        return;
                    }
                    if (array_HEADER != null)
                    {
                        QUANTITYRS = Convert.ToDouble(array_HEADER[44], CultureInfo.InvariantCulture);
                    }
                }

                int j = 0;
                int countRS = array_GOODS_RS == null ? 0 : array_GOODS_RS.Count();
                array_GOODS = new string[recordCount + countRS][];
                for (j = 0; j < countRS; j++)
                {
                    array_GOODS[j] = new string[13];
                    array_GOODS[j][0] = array_GOODS_RS[j][0]; //ID
                    array_GOODS[j][1] = array_GOODS_RS[j][1]; //W_NAME
                    array_GOODS[j][2] = array_GOODS_RS[j][2]; //UNIT_ID 
                    array_GOODS[j][3] = ""; //ერთეულის სახელი UNIT_TXT
                    array_GOODS[j][4] = array_GOODS_RS[j][3]; //QUANTITY
                    array_GOODS[j][5] = array_GOODS_RS[j][4]; //PRICE
                    array_GOODS[j][6] = "-1"; //STATUS 1 ან -1 თუ გადაეცით -1 შესაბამისი საქონელი წაიშლება
                    array_GOODS[j][7] = array_GOODS_RS[j][5]; //AMOUNT
                    array_GOODS[j][8] = array_GOODS_RS[j][6]; //პროგრამის კოდი
                    array_GOODS[j][9] = array_GOODS_RS[j][7]; //A_ID
                    array_GOODS[j][10] = array_GOODS_RS[j][8]; //VAT_TYPE 0 - ჩეულებრივი; 1 - ნულოვალი; 2 - დაუბეგრავი
                    array_GOODS[j][11] = array_GOODS_RS[j][9]; //QUANTITY_EXT
                }
                //<--- წასაშლელი Goods

                i = j;
                while (!oRecordSet.EoF)
                {
                    array_GOODS[i] = new string[13];
                    array_GOODS[i][0] = "0"; //ID ზედნადებში საქონლის ჩანაწერის ID გადაეცემა 0 თუ ახალი იქმნება
                    array_GOODS[i][1] = oRecordSet.Fields.Item("W_NAME").Value.ToString(); //W_NAME
                    array_GOODS[i][2] = oRecordSet.Fields.Item("UNIT_ID").Value.ToString(); //UNIT_ID 1
                    string UNIT_TXT = oRecordSet.Fields.Item("UNIT_TXT").Value.ToString();
                    array_GOODS[i][3] = array_GOODS[i][2] == "99" ? (UNIT_TXT == "" ? "სხვა" : UNIT_TXT) : "";//ერთეულის სახელი აუცილებელია როდესაც UNIT_ID=99 („სხვა“)UNIT_TXT
                    array_GOODS[i][4] = oRecordSet.Fields.Item("QUANTITY").Value.ToString(Nfi); //QUANTITY
                    array_GOODS[i][5] = oRecordSet.Fields.Item("PRICE").Value.ToString(Nfi); //PRICE
                    array_GOODS[i][6] = "1"; //STATUS 1 ან -1 თუ გადაეცით -1 შესაბამისი საქონელი წაიშლება
                    array_GOODS[i][7] = oRecordSet.Fields.Item("AMOUNT").Value.ToString(Nfi); //AMOUNT
                    switch (itemCode)
                    {
                        case "0":
                            array_GOODS[i][8] = oRecordSet.Fields.Item("ItemCode").Value.ToString(); //პროგრამის კოდი //BAR_CODE
                            break;
                        case "1":
                            array_GOODS[i][8] = oRecordSet.Fields.Item("AdditionalIdentifier").Value.ToString(); //არტიკული       //BAR_CODE  
                            break;
                        case "2":
                            array_GOODS[i][8] = oRecordSet.Fields.Item("CodeBars").Value.ToString(); //ძირითადი შტრიხკოდი  //BAR_CODE
                            break;
                    }
                    array_GOODS[i][9] = oRecordSet.Fields.Item("A_ID").Value.ToString(); //A_ID თუ აქციზური არ არის გადაეცით 0.
                    string VAT_TYPE = oRecordSet.Fields.Item("VAT_TYPE").Value.ToString(); //VAT_TYPE 0 - ჩეულებრივი; 1 - ნულოვალი; 2 - დაუბეგრავი
                    string VatGroup = oRecordSet.Fields.Item("VatGroup").Value.ToString();
                    VAT_TYPE = VatGroup == "X0" ? "" : VAT_TYPE;
                    switch (VAT_TYPE)
                    {
                        //case "18": VAT_TYPE = "0"; //ჩეულებრივი 18%
                        //    break;
                        case "0":
                            VAT_TYPE = "1"; //ნულოვანი 0%
                            break;
                        case "":
                            VAT_TYPE = "2"; //დაუბეგრავი
                            break;
                        default:
                            VAT_TYPE = "0"; //ჩეულებრივი 18%
                            break;
                    }
                    array_GOODS[i][10] = VAT_TYPE; //VAT_TYPE 0 - ჩეულებრივი; 1 - ნულოვალი; 2 - დაუბეგრავი
                    array_GOODS[i][11] = ""; //QUANTITY_EXT          

                    i = i + 1;
                    oRecordSet.MoveNext();
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                GC.Collect();
            }
        }

        private static void getArrayGoodsGoodsIssueType(WayBill oWayBill, int baseDocEntry, string ID, out string[][] array_GOODS, out double QUANTITYRS, out string errorText)
        {
            errorText = null;
            array_GOODS = null;
            QUANTITYRS = 0;

            Dictionary<string, string> rsSettings = CompanyDetails.getRSSettings(out errorText);
            if (errorText != null)
            {
                return;
            }
            string itemCode = rsSettings["ItemCode"];

            Recordset oRecordSet = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            string query = "SELECT " +
            "\"IGE1\".\"LineNum\", " +
            "'" + ID + "'" + " AS \"ID\", " +
            "\"IGE1\".\"DocEntry\" AS \"DocEntry\", " +
            "\"IGE1\".\"ItemCode\" AS \"ItemCode\", " +
            "\"OITM\".\"CodeBars\" AS \"CodeBars\", " +
            "\"OITM\".\"SWW\" AS \"AdditionalIdentifier\", " +
            "\"IGE1\".\"Dscription\" AS \"W_NAME\", " +
            "CASE WHEN \"BDO_RSUOM\".\"U_RSCode\" is null THEN '99' ELSE \"BDO_RSUOM\".\"U_RSCode\" END AS \"UNIT_ID\", " +
            "CASE WHEN \"IGE1\".\"unitMsr\"='' THEN 'სხვა' ELSE \"IGE1\".\"unitMsr\" END  AS \"UNIT_TXT\", " +
            "\"IGE1\".\"VatPrcnt\" AS \"VAT_TYPE\", " +
            "\"IGE1\".\"VatGroup\"AS \"VatGroup\", " +
            "'0' AS \"A_ID\", " +
            "\"IGE1\".\"Quantity\" * \"IGE1\".\"NumPerMsr\" AS \"Quantity\", " +
            "'0' AS \"AMOUNT\"," +
            "'0' AS \"PRICE\" " +

            "FROM \"IGE1\" AS \"IGE1\" " +

            "LEFT JOIN \"OITM\" AS \"OITM\" " +
            "ON \"IGE1\".\"ItemCode\" = \"OITM\".\"ItemCode\" " +

            "LEFT JOIN \"OUOM\" AS \"OUOM\" " +
            "ON \"IGE1\".\"unitMsr\" = \"OUOM\".\"UomName\" " +

            "LEFT JOIN \"@BDO_RSUOM\" AS \"BDO_RSUOM\" " +
            "ON \"OUOM\".\"UomEntry\" = \"BDO_RSUOM\".\"U_UomEntry\" " +

            "WHERE \"IGE1\".\"DocEntry\" = '" + baseDocEntry + "' AND (\"OITM\".\"ItemType\" = 'I' OR \"OITM\".\"ItemType\" = 'F' ) AND \"OITM\".\"InvntItem\" = 'Y'";

            try
            {
                oRecordSet.DoQuery(query);
                int recordCount = oRecordSet.RecordCount;

                NumberFormatInfo Nfi = new NumberFormatInfo() { NumberDecimalSeparator = "." };
                int i = 0;

                //წასაშლელი Goods --->      
                string[] array_HEADER = null;
                string[][] array_GOODS_RS = null;
                string[][] arry_SUB_WAYBILLS = null;

                if ((ID == "0" || ID == "") == false)
                {
                    int get_waybill_result_int = oWayBill.get_waybill(Convert.ToInt32(ID), out array_HEADER, out array_GOODS_RS, out arry_SUB_WAYBILLS, out errorText);
                    if (get_waybill_result_int != 1)
                    {
                        return;
                    }
                    if (array_HEADER != null)
                    {
                        QUANTITYRS = Convert.ToDouble(array_HEADER[44], CultureInfo.InvariantCulture);
                    }
                }

                int j = 0;
                int countRS = array_GOODS_RS == null ? 0 : array_GOODS_RS.Count();
                array_GOODS = new string[recordCount + countRS][];
                for (j = 0; j < countRS; j++)
                {
                    array_GOODS[j] = new string[13];
                    array_GOODS[j][0] = array_GOODS_RS[j][0]; //ID
                    array_GOODS[j][1] = array_GOODS_RS[j][1]; //W_NAME
                    array_GOODS[j][2] = array_GOODS_RS[j][2]; //UNIT_ID 
                    array_GOODS[j][3] = ""; //ერთეულის სახელი UNIT_TXT
                    array_GOODS[j][4] = array_GOODS_RS[j][3]; //QUANTITY
                    array_GOODS[j][5] = array_GOODS_RS[j][4]; //PRICE
                    array_GOODS[j][6] = "-1"; //STATUS 1 ან -1 თუ გადაეცით -1 შესაბამისი საქონელი წაიშლება
                    array_GOODS[j][7] = array_GOODS_RS[j][5]; //AMOUNT
                    array_GOODS[j][8] = array_GOODS_RS[j][6]; //პროგრამის კოდი
                    array_GOODS[j][9] = array_GOODS_RS[j][7]; //A_ID
                    array_GOODS[j][10] = array_GOODS_RS[j][8]; //VAT_TYPE 0 - ჩეულებრივი; 1 - ნულოვალი; 2 - დაუბეგრავი
                    array_GOODS[j][11] = array_GOODS_RS[j][9]; //QUANTITY_EXT
                }
                //<--- წასაშლელი Goods

                i = j;
                while (!oRecordSet.EoF)
                {
                    array_GOODS[i] = new string[13];
                    array_GOODS[i][0] = "0"; //ID ზედნადებში საქონლის ჩანაწერის ID გადაეცემა 0 თუ ახალი იქმნება
                    array_GOODS[i][1] = oRecordSet.Fields.Item("W_NAME").Value.ToString(); //W_NAME
                    array_GOODS[i][2] = oRecordSet.Fields.Item("UNIT_ID").Value.ToString(); //UNIT_ID 1
                    string UNIT_TXT = oRecordSet.Fields.Item("UNIT_TXT").Value.ToString();
                    array_GOODS[i][3] = array_GOODS[i][2] == "99" ? (UNIT_TXT == "" ? "სხვა" : UNIT_TXT) : "";//ერთეულის სახელი აუცილებელია როდესაც UNIT_ID=99 („სხვა“)UNIT_TXT
                    array_GOODS[i][4] = oRecordSet.Fields.Item("QUANTITY").Value.ToString(Nfi); //QUANTITY
                    array_GOODS[i][5] = oRecordSet.Fields.Item("PRICE").Value.ToString(Nfi); //PRICE
                    array_GOODS[i][6] = "1"; //STATUS 1 ან -1 თუ გადაეცით -1 შესაბამისი საქონელი წაიშლება
                    array_GOODS[i][7] = oRecordSet.Fields.Item("AMOUNT").Value.ToString(Nfi); //AMOUNT
                    switch (itemCode)
                    {
                        case "0":
                            array_GOODS[i][8] = oRecordSet.Fields.Item("ItemCode").Value.ToString(); //პროგრამის კოდი //BAR_CODE
                            break;
                        case "1":
                            array_GOODS[i][8] = oRecordSet.Fields.Item("AdditionalIdentifier").Value.ToString(); //არტიკული       //BAR_CODE  
                            break;
                        case "2":
                            array_GOODS[i][8] = oRecordSet.Fields.Item("CodeBars").Value.ToString(); //ძირითადი შტრიხკოდი  //BAR_CODE
                            break;
                    }
                    array_GOODS[i][9] = oRecordSet.Fields.Item("A_ID").Value.ToString(); //A_ID თუ აქციზური არ არის გადაეცით 0.
                    string VAT_TYPE = oRecordSet.Fields.Item("VAT_TYPE").Value.ToString(); //VAT_TYPE 0 - ჩეულებრივი; 1 - ნულოვალი; 2 - დაუბეგრავი
                    string VatGroup = oRecordSet.Fields.Item("VatGroup").Value.ToString();
                    VAT_TYPE = VatGroup == "X0" ? "" : VAT_TYPE;
                    switch (VAT_TYPE)
                    {
                        //case "18": VAT_TYPE = "0"; //ჩეულებრივი 18%
                        //    break;
                        case "0":
                            VAT_TYPE = "1"; //ნულოვანი 0%
                            break;
                        case "":
                            VAT_TYPE = "2"; //დაუბეგრავი
                            break;
                        default:
                            VAT_TYPE = "0"; //ჩეულებრივი 18%
                            break;
                    }
                    array_GOODS[i][10] = VAT_TYPE; //VAT_TYPE 0 - ჩეულებრივი; 1 - ნულოვალი; 2 - დაუბეგრავი
                    array_GOODS[i][11] = ""; //QUANTITY_EXT          

                    i = i + 1;
                    oRecordSet.MoveNext();
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                GC.Collect();
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="Program.oCompany"></param>
        /// <param name="oWayBill"></param>
        /// <param name="baseDocEntry"></param>
        /// <param name="baseDocType"></param>
        /// <param name="ID"></param>
        /// <param name="errorText"></param>
        /// <returns>1 - თუ ყველაფერი კარგადაა,  0 თუ სინქრონიზაცია დარღვეულია</returns>
        private static int checkSync(WayBill oWayBill, int baseDocEntry, string baseDocType, string ID, out string errorText)
        {
            errorText = null;

            if (ID != "0")
            {
                string[][] array_GOODS = null;
                double QUANTITYRS = 0;
                double AMOUNTRS = 0;

                if (baseDocType == "13") //A/R Invoice
                {
                    getArrayGoodsARInvoiceType(oWayBill, baseDocEntry, ID, out array_GOODS, out QUANTITYRS, out AMOUNTRS, out errorText);
                    if (array_GOODS != null)
                    {
                        double QUANTITY = 0;
                        double AMOUNT = 0;

                        for (int i = 0; i < array_GOODS.Count(); i++)
                        {
                            if (array_GOODS[i][6] == "1")
                            {
                                QUANTITY = QUANTITY + Convert.ToDouble(array_GOODS[i][4], CultureInfo.InvariantCulture);
                                AMOUNT = AMOUNT + Convert.ToDouble(array_GOODS[i][7], CultureInfo.InvariantCulture);
                            }
                        }
                        if (QUANTITY != QUANTITYRS || AMOUNT != AMOUNTRS)
                        {
                            return 0; //სინქრონიზაცია დარღვეულია
                        }
                    }
                }

                else if (baseDocType == "15") //Delivery
                {
                    getArrayGoodsDeliveryType(oWayBill, baseDocEntry, ID, out array_GOODS, out QUANTITYRS, out AMOUNTRS, out errorText);
                    if (array_GOODS != null)
                    {
                        double QUANTITY = 0;
                        double AMOUNT = 0;

                        for (int i = 0; i < array_GOODS.Count(); i++)
                        {
                            if (array_GOODS[i][6] == "1")
                            {
                                QUANTITY = QUANTITY + Convert.ToDouble(array_GOODS[i][4], CultureInfo.InvariantCulture);
                                AMOUNT = AMOUNT + Convert.ToDouble(array_GOODS[i][7], CultureInfo.InvariantCulture);
                            }
                        }
                        if (QUANTITY != QUANTITYRS || AMOUNT != AMOUNTRS)
                        {
                            return 0; //სინქრონიზაცია დარღვეულია
                        }
                    }
                }

                else if (baseDocType == "67") //Inventory Transfer
                {
                    getArrayGoodsInventoryTransferType(oWayBill, baseDocEntry, ID, out array_GOODS, out QUANTITYRS, out errorText);
                    if (array_GOODS != null)
                    {
                        double QUANTITY = 0;

                        for (int i = 0; i < array_GOODS.Count(); i++)
                        {
                            if (array_GOODS[i][6] == "1")
                            {
                                QUANTITY = QUANTITY + Convert.ToDouble(array_GOODS[i][4], CultureInfo.InvariantCulture);
                            }
                        }
                        if (QUANTITY != QUANTITYRS)
                        {
                            return 0; //სინქრონიზაცია დარღვეულია
                        }
                    }
                }

                else if (baseDocType == "UDO_F_BDOSFASTRD_D") //Fixed Asset Transfer
                {
                    getArrayGoodsFixedAssetTransferType(oWayBill, baseDocEntry, ID, out array_GOODS, out QUANTITYRS, out errorText);
                    if (array_GOODS != null)
                    {
                        double QUANTITY = 0;

                        for (int i = 0; i < array_GOODS.Count(); i++)
                        {
                            if (array_GOODS[i][6] == "1")
                            {
                                QUANTITY = QUANTITY + Convert.ToDouble(array_GOODS[i][4], CultureInfo.InvariantCulture);
                            }
                        }
                        if (QUANTITY != QUANTITYRS)
                        {
                            return 0; //სინქრონიზაცია დარღვეულია
                        }
                    }
                }
                else if (baseDocType == "14") //A/R Credit Memo
                {
                    getArrayGoodsInvoiceCreditMemoType(oWayBill, baseDocEntry, ID, out array_GOODS, out QUANTITYRS, out AMOUNTRS, out errorText);
                    if (array_GOODS != null)
                    {
                        double QUANTITY = 0;
                        double AMOUNT = 0;

                        for (int i = 0; i < array_GOODS.Count(); i++)
                        {
                            if (array_GOODS[i][6] == "1")
                            {
                                QUANTITY = QUANTITY + Convert.ToDouble(array_GOODS[i][4], CultureInfo.InvariantCulture);
                                AMOUNT = AMOUNT + Convert.ToDouble(array_GOODS[i][7], CultureInfo.InvariantCulture);
                            }
                        }
                        if (QUANTITY != QUANTITYRS || AMOUNT != AMOUNTRS)
                        {
                            return 0; //სინქრონიზაცია დარღვეულია
                        }
                    }
                }

                else if (baseDocType == "165") //A/R Correction Invoice
                {
                    getArrayGoodsInvoiceCorrectionType(oWayBill, baseDocEntry, ID, out array_GOODS, out QUANTITYRS, out AMOUNTRS, out errorText);
                    if (array_GOODS != null)
                    {
                        double QUANTITY = 0;
                        double AMOUNT = 0;

                        for (int i = 0; i < array_GOODS.Count(); i++)
                        {
                            if (array_GOODS[i][6] == "1")
                            {
                                QUANTITY = QUANTITY + Convert.ToDouble(array_GOODS[i][4], CultureInfo.InvariantCulture);
                                AMOUNT = AMOUNT + Convert.ToDouble(array_GOODS[i][7], CultureInfo.InvariantCulture);
                            }
                        }
                        if (QUANTITY != QUANTITYRS || AMOUNT != AMOUNTRS)
                        {
                            return 0; //სინქრონიზაცია დარღვეულია
                        }
                    }
                }

                else if (baseDocType == "60") //Goods Issue
                {
                    getArrayGoodsGoodsIssueType(oWayBill, baseDocEntry, ID, out array_GOODS, out QUANTITYRS, out errorText);
                    if (array_GOODS != null)
                    {
                        double QUANTITY = 0;

                        for (int i = 0; i < array_GOODS.Count(); i++)
                        {
                            if (array_GOODS[i][6] == "1")
                            {
                                QUANTITY = QUANTITY + Convert.ToDouble(array_GOODS[i][4], CultureInfo.InvariantCulture);
                            }
                        }
                        if (QUANTITY != QUANTITYRS)
                        {
                            return 0; //სინქრონიზაცია დარღვეულია
                        }
                    }
                }
            }
            return 1;
        }

        public static void closeWaybill(int docEntry, int baseDocEntry, out string errorText)
        {
            errorText = null;

            Dictionary<string, string> rsSettings = CompanyDetails.getRSSettings(out errorText);
            if (errorText != null)
            {
                return;
            }

            string su = rsSettings["SU"];
            string sp = rsSettings["SP"];

            WayBill oWayBill = new WayBill(rsSettings["ProtocolType"]);

            bool chek_service_user = oWayBill.chek_service_user(su, sp, out errorText);
            if (chek_service_user == false)
            {
                errorText = BDOSResources.getTranslate("ServiceUserPasswordNotCorrect");
                return;
            }

            Recordset oRecordSet = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            string query = "SELECT " +
            "\"BDO_WBLD\".\"U_wblID\", " +
            "\"BDO_WBLD\".\"U_recpInfN\", " +
            "\"BDO_WBLD\".\"U_recvInfN\", " +
            "\"BDO_WBLD\".\"U_baseDocT\", " +
            "\"BDO_WBLD\".\"U_delvDate\", " +
            "\"BDO_WBLD\".\"U_beginTime\", " +
            "\"BDO_WBLD\".\"U_begDate\" " +
            "FROM \"@BDO_WBLD\" AS \"BDO_WBLD\" " +
            "WHERE \"BDO_WBLD\".\"DocEntry\" = '" + docEntry + "'";

            try
            {
                oRecordSet.DoQuery(query);
                int ID = 0;
                DateTime DELIVERY_DATE = DateTime.MinValue;
                DateTime BEGIN_DATE = DateTime.MinValue;
                string baseDocType = null;

                while (!oRecordSet.EoF)
                {
                    ID = Convert.ToInt32(oRecordSet.Fields.Item("U_wblID").Value);
                    baseDocType = oRecordSet.Fields.Item("U_baseDocT").Value.ToString();
                    BEGIN_DATE = DateTime.TryParse(oRecordSet.Fields.Item("U_begDate").Value.ToString("yyyyMMdd") == "18991230" ? "" : oRecordSet.Fields.Item("U_begDate").Value.ToString(), out BEGIN_DATE) == false ? DateTime.MinValue : BEGIN_DATE;  //BEGIN_DATE - ტრანსპორტირების დაწყების თარიღი
                    BEGIN_DATE = BEGIN_DATE == DateTime.MinValue ? DateTime.Today : BEGIN_DATE;
                    DELIVERY_DATE = DateTime.TryParse(oRecordSet.Fields.Item("U_delvDate").Value.ToString("yyyyMMdd") == "18991230" ? "" : oRecordSet.Fields.Item("U_delvDate").Value.ToString(), out DELIVERY_DATE) == false ? DateTime.MinValue : DELIVERY_DATE; //DELIVERY_DATE - მიწოდების თარიღი გადასცემთ უკვე აქტიურს დახურვის წინ
                    DELIVERY_DATE = DELIVERY_DATE <= BEGIN_DATE ? DateTime.Now : DELIVERY_DATE;


                    ///////////////////
                    decimal U_beginTime = Convert.ToDecimal(oRecordSet.Fields.Item("U_beginTime").Value);
                    int Hour = Convert.ToInt32(Math.Round(U_beginTime / 100));
                    int Min = Convert.ToInt32(U_beginTime - Hour * 100);

                    BEGIN_DATE = DateTime.TryParse(oRecordSet.Fields.Item("U_begDate").Value.ToString("yyyyMMdd") == "18991230" ? "" : oRecordSet.Fields.Item("U_begDate").Value.ToString(), out BEGIN_DATE) == false ? DateTime.MinValue : BEGIN_DATE;//BEGIN_DATE - ტრანსპორტირების დაწყების თარიღი
                    BEGIN_DATE = BEGIN_DATE == DateTime.MinValue || BEGIN_DATE < DateTime.Today ? DateTime.Now : BEGIN_DATE;

                    BEGIN_DATE = new DateTime(BEGIN_DATE.Year, BEGIN_DATE.Month, BEGIN_DATE.Day, Hour, Min, 0);
                    DELIVERY_DATE = DELIVERY_DATE < BEGIN_DATE ? BEGIN_DATE.AddMinutes(1) : DELIVERY_DATE;
                    /////////////////

                    oRecordSet.MoveNext();
                    break;
                }

                //სინქრონიზაციის შემოწმება --->
                if (checkSync(oWayBill, baseDocEntry, baseDocType, ID.ToString(), out errorText) == 0)
                {
                    errorText = BDOSResources.getTranslate("SynchronisationViolatedCorrectWaybill");
                    return;
                }
                //<--- სინქრონიზაციის შემოწმება

                int close_waybill_result_int = oWayBill.close_waybill(ID, out errorText);
                if (close_waybill_result_int != 1)
                {
                    return;
                }

                CompanyService oCompanyService = null;
                GeneralService oGeneralService = null;
                GeneralData oGeneralData = null;
                GeneralDataParams oGeneralParams = null;
                oCompanyService = oCompany.GetCompanyService();
                oGeneralService = oCompanyService.GetGeneralService("UDO_F_BDO_WBLD_D");
                oGeneralData = ((GeneralData)(oGeneralService.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralData)));

                //Get UDO record
                oGeneralParams = oGeneralService.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralDataParams);
                oGeneralParams.SetProperty("DocEntry", docEntry);
                oGeneralData = oGeneralService.GetByParams(oGeneralParams);

                //Update UDO record
                string STATUS_RS = "3";  //"დასრულებული";
                oGeneralData.SetProperty("U_status", STATUS_RS);
                oGeneralData.SetProperty("U_delvDate", DELIVERY_DATE);
                oGeneralService.Update(oGeneralData);
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                GC.Collect();
            }
        }

        public static string getInitFromTIN(string tin, out string errorText)
        {
            errorText = null;

            Dictionary<string, string> rsSettings = CompanyDetails.getRSSettings(out errorText);
            if (errorText != null)
            {
                return "";
            }

            WayBill oWayBill = new WayBill(rsSettings["ProtocolType"]);

            string su = rsSettings["SU"];
            string sp = rsSettings["SP"];

            bool chek_service_user = oWayBill.chek_service_user(su, sp, out errorText);
            if (chek_service_user == false)
            {
                errorText = BDOSResources.getTranslate("ServiceUserPasswordNotCorrect");
                return "";
            }

            string result = oWayBill.get_name_from_tin(tin, out errorText);

            return result;
        }

        public static void getWaybill(int docEntry, int baseDocEntry, out string errorText)
        {
            errorText = null;

            Dictionary<string, string> rsSettings = CompanyDetails.getRSSettings(out errorText);
            if (errorText != null)
            {
                return;
            }

            string su = rsSettings["SU"];
            string sp = rsSettings["SP"];

            WayBill oWayBill = new WayBill(rsSettings["ProtocolType"]);

            bool chek_service_user = oWayBill.chek_service_user(su, sp, out errorText);
            if (chek_service_user == false)
            {
                errorText = BDOSResources.getTranslate("ServiceUserPasswordNotCorrect");
                return;
            }

            Recordset oRecordSet = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            string query = "SELECT " +
            "\"BDO_WBLD\".\"U_wblID\", " +
            "\"BDO_WBLD\".\"U_baseDocT\" " +
            "FROM \"@BDO_WBLD\" AS \"BDO_WBLD\" " +
            "WHERE \"BDO_WBLD\".\"DocEntry\" = '" + docEntry + "' AND \"BDO_WBLD\".\"U_baseDoc\" = '" + baseDocEntry + "'";

            try
            {
                oRecordSet.DoQuery(query);
                int ID = 0;
                string baseDocType = null;

                while (!oRecordSet.EoF)
                {
                    baseDocType = oRecordSet.Fields.Item("U_baseDocT").Value.ToString();
                    ID = oRecordSet.Fields.Item("U_wblID").Value == "" ? 0 : Convert.ToInt32(oRecordSet.Fields.Item("U_wblID").Value);
                    oRecordSet.MoveNext();
                    break;
                }

                string[] array_HEADER = null;
                string[][] array_GOODS = null;
                string[][] arry_SUB_WAYBILLS = null;

                if (ID == 0)
                {
                    errorText = BDOSResources.getTranslate("ForStatusUpdateFillID");
                    return;
                }

                int get_waybill_result_int = oWayBill.get_waybill(ID, out array_HEADER, out array_GOODS, out arry_SUB_WAYBILLS, out errorText);
                if (get_waybill_result_int != 1)
                {
                    return;
                }

                string STATUS_RS = null;
                switch (array_HEADER[15]) //STATUS
                {
                    case "0":
                        STATUS_RS = "1"; //"შენახული"
                        break;
                    case "1":
                        STATUS_RS = "2";  //"აქტიური"
                        break;
                    case "2":
                        STATUS_RS = "3";  //"დასრულებული"
                        break;
                    case "-1":
                        STATUS_RS = "4";  //"წაშლილი"
                        break;
                    case "-2":
                        STATUS_RS = "5";  //"გაუქმებული"
                        break;
                    case "8":
                        STATUS_RS = "6";  //"გადამზიდავთან გადაგზავნილი"
                        break;
                }

                CompanyService oCompanyService = null;
                GeneralService oGeneralService = null;
                GeneralData oGeneralData = null;
                GeneralDataParams oGeneralParams = null;
                oCompanyService = oCompany.GetCompanyService();
                oGeneralService = oCompanyService.GetGeneralService("UDO_F_BDO_WBLD_D");
                oGeneralData = ((GeneralData)(oGeneralService.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralData)));

                //Get UDO record
                oGeneralParams = oGeneralService.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralDataParams);
                oGeneralParams.SetProperty("DocEntry", docEntry);
                oGeneralData = oGeneralService.GetByParams(oGeneralParams);

                //Update UDO record 
                double QUANTITY = 0;
                double AMOUNT = 0;

                DateTime DELIVERY_DATE = DateTime.TryParse(array_HEADER[14], out DELIVERY_DATE) == false ? DateTime.MinValue : DELIVERY_DATE;
                DateTime BEGIN_DATE = DateTime.TryParse(array_HEADER[25], out BEGIN_DATE) == false ? DateTime.MinValue : BEGIN_DATE;
                DateTime ACTIVATE_DATE = DateTime.TryParse(array_HEADER[17], out ACTIVATE_DATE) == false ? DateTime.MinValue : ACTIVATE_DATE;


                QUANTITY = Convert.ToDouble(array_HEADER[44], CultureInfo.InvariantCulture);
                AMOUNT = Convert.ToDouble(array_HEADER[45], CultureInfo.InvariantCulture);

                oGeneralData.SetProperty("U_status", STATUS_RS); //STATUS
                oGeneralData.SetProperty("U_number", array_HEADER[22]); //WAYBILL_NUMBER
                oGeneralData.SetProperty("U_delvDate", DELIVERY_DATE); //DELIVERY_DATE
                oGeneralData.SetProperty("U_begDate", BEGIN_DATE); //BEGIN_DATE
                //oGeneralData.SetProperty("", array_HEADER[30]); //IS_CONFIRMED
                oGeneralData.SetProperty("U_actDate", ACTIVATE_DATE); //ACTIVATE_DATE
                oGeneralData.SetProperty("U_endAddrs", array_HEADER[7]); //end add
                oGeneralData.SetProperty("U_strAddrs", array_HEADER[6]); //start add

                //თუ ტრანსპორტირებითაა
                if (oGeneralData.GetProperty("U_type") == "0")
                {
                    oGeneralData.SetProperty("U_drivTin", array_HEADER[8]); //DRIVER_TIN
                    oGeneralData.SetProperty("U_notRsdnt", array_HEADER[9] == "0" ? "Y" : "N"); //CHEK_DRIVER_TIN
                    oGeneralData.SetProperty("U_vehicNum", array_HEADER[21]); //CAR_NUMBER
                    oGeneralData.SetProperty("U_trailNum", array_HEADER[28]); //TRANS_TXT
                    oGeneralData.SetProperty("U_drvCode", array_HEADER[10]); //DRIVER_NAME
                }

                //თუ გაუქმებულია ან წაშლილია
                if (STATUS_RS == "5" || STATUS_RS == "4")
                {
                    oGeneralData.SetProperty("U_number", ""); //НомерТН
                    oGeneralData.SetProperty("U_wblID", ""); //ID
                    oGeneralData.SetProperty("U_begDate", DateTime.MinValue); //BEGIN_DATE
                    oGeneralData.SetProperty("U_delvDate", DateTime.MinValue); //DELIVERY_DATE
                    //oGeneralData.SetProperty("", false); //IS_CONFIRMED
                }

                oGeneralService.Update(oGeneralData);

                //სინქრონიზაციის შემოწმება --->
                if (checkSync(oWayBill, baseDocEntry, baseDocType, ID.ToString(), out errorText) == 0)
                {
                    errorText = BDOSResources.getTranslate("SynchronisationViolatedCorrectWaybill");
                    return;
                }
                //<--- სინქრონიზაციის შემოწმება
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                GC.Collect();
            }
        }

        public static void refWaybill(int docEntry, int baseDocEntry, out string errorText)
        {
            errorText = null;

            Dictionary<string, string> rsSettings = CompanyDetails.getRSSettings(out errorText);
            if (errorText != null)
            {
                return;
            }

            string su = rsSettings["SU"];
            string sp = rsSettings["SP"];

            WayBill oWayBill = new WayBill(rsSettings["ProtocolType"]);

            bool chek_service_user = oWayBill.chek_service_user(su, sp, out errorText);
            if (chek_service_user == false)
            {
                errorText = BDOSResources.getTranslate("ServiceUserPasswordNotCorrect");
                return;
            }

            Recordset oRecordSet = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            string query = "SELECT " +
            "\"BDO_WBLD\".\"U_wblID\", " +
            "\"BDO_WBLD\".\"U_comment\" " +
            "FROM \"@BDO_WBLD\" AS \"BDO_WBLD\" " +
            "WHERE \"BDO_WBLD\".\"DocEntry\" = '" + docEntry + "'";

            try
            {
                oRecordSet.DoQuery(query);
                int ID = 0;
                string comment = "";
                while (!oRecordSet.EoF)
                {
                    ID = Convert.ToInt32(oRecordSet.Fields.Item("U_wblID").Value);
                    comment = oRecordSet.Fields.Item("U_comment").Value;
                    oRecordSet.MoveNext();
                    break;
                }

                int close_waybill_result_int = oWayBill.ref_waybill_vd(ID, comment, out errorText);
                if (close_waybill_result_int != 1)
                {
                    return;
                }

                CompanyService oCompanyService = null;
                GeneralService oGeneralService = null;
                GeneralData oGeneralData = null;
                GeneralDataParams oGeneralParams = null;
                oCompanyService = oCompany.GetCompanyService();
                oGeneralService = oCompanyService.GetGeneralService("UDO_F_BDO_WBLD_D");
                oGeneralData = ((GeneralData)(oGeneralService.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralData)));

                //Get UDO record
                oGeneralParams = oGeneralService.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralDataParams);
                oGeneralParams.SetProperty("DocEntry", docEntry);
                oGeneralData = oGeneralService.GetByParams(oGeneralParams);

                //Update UDO record
                string STATUS_RS = "5";  //"გაუქმებული";
                oGeneralData.SetProperty("U_status", STATUS_RS);
                oGeneralData.SetProperty("U_wblID", "");
                oGeneralData.SetProperty("U_number", "");
                oGeneralData.SetProperty("U_delvDate", DateTime.MinValue);
                oGeneralData.SetProperty("U_begDate", DateTime.MinValue);
                oGeneralData.SetProperty("U_actDate", DateTime.MinValue);

                oGeneralService.Update(oGeneralData);
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                GC.Collect();
            }
        }
        #endregion

        public static bool canCreateDocument(int docEntry, string objectType)
        {
            Recordset oRecordSet = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            StringBuilder query = new StringBuilder();
            query.Append("select \"U_baseDoc\" from \"@BDO_WBLD\" \n");
            query.Append("where \"Canceled\" = 'N' and \"U_baseDocT\" = '" + objectType + "' and \"U_baseDoc\" = '" + docEntry + "'");

            try
            {
                oRecordSet.DoQuery(query.ToString());

                if (!oRecordSet.EoF)
                {
                    uiApp.SetStatusBarMessage(BDOSResources.getTranslate("WaybillAlreadyExistsForThisDocument"), SAPbouiCOM.BoMessageTime.bmt_Short);
                    return false;
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                oRecordSet = null;
                GC.Collect();
            }
            return true;
        }
    }
}