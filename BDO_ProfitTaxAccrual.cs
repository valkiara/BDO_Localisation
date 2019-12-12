using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace BDO_Localisation_AddOn
{
    static partial class BDO_ProfitTaxAccrual
    {
        public static decimal profitTaxRate = 0;
        public static bool ProfitTaxTypeIsSharing = false;
        public static int removeRecordRow = 0;

        public static void createDocumentUDO( out string errorText)
        {
            errorText = null;
            string tableName = "BDO_TAXP";
            string description = "Profit Tax Accrual";

            SAPbobsCOM.UserObjectsMD oUserObjectsMD = null;
            oUserObjectsMD = (SAPbobsCOM.UserObjectsMD)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectsMD);

            int result = UDO.addUserTable( tableName, description, SAPbobsCOM.BoUTBTableType.bott_Document, out errorText);

            if (result != 0)
            {
                return;
            }

            Dictionary<string, object> fieldskeysMap;

            fieldskeysMap = new Dictionary<string, object>(); // დაბეგვრის ობიექტი 
            fieldskeysMap.Add("Name", "prBase");
            fieldskeysMap.Add("TableName", "BDO_TAXP");
            fieldskeysMap.Add("Description", "Profit Base");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 15);
            fieldskeysMap.Add("Mandatory", SAPbobsCOM.BoYesNoEnum.tYES);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "prBsDscr");
            fieldskeysMap.Add("TableName", "BDO_TAXP");
            fieldskeysMap.Add("Description", "Base Type Description");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 100);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // DocDate 
            fieldskeysMap.Add("Name", "docDate");
            fieldskeysMap.Add("TableName", "BDO_TAXP");
            fieldskeysMap.Add("Description", "Posting Date");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Date);
            fieldskeysMap.Add("EditSize", 15);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            tableName = "BDO_TXP1";
            description = "Profit Tax Accrual Child1";

            result = UDO.addUserTable( tableName, description, SAPbobsCOM.BoUTBTableType.bott_DocumentLines, out errorText);

            if (result != 0)
            {
                return;
            }

            Dictionary<string, string> listValidValuesDict = new Dictionary<string, string>();
            listValidValuesDict = new Dictionary<string, string>();
            listValidValuesDict.Add("Uncrediting", "Uncrediting");
            listValidValuesDict.Add("Accrual", "Accrual");

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "txType"); //დაბეგვრის ტიპი
            fieldskeysMap.Add("TableName", "BDO_TXP1");
            fieldskeysMap.Add("Description", "Base Type");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);
            fieldskeysMap.Add("ValidValues", listValidValuesDict);
            fieldskeysMap.Add("Mandatory", SAPbobsCOM.BoYesNoEnum.tYES);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "amtTx"); //დასაბეგრი თანხა
            fieldskeysMap.Add("TableName", "BDO_TXP1");
            fieldskeysMap.Add("Description", "Amount Taxable");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);
            fieldskeysMap.Add("Mandatory", SAPbobsCOM.BoYesNoEnum.tYES);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "amtPrTx"); //მოგების გადასახადი
            fieldskeysMap.Add("TableName", "BDO_TXP1");
            fieldskeysMap.Add("Description", "Profit Tax Amount");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);
            fieldskeysMap.Add("Mandatory", SAPbobsCOM.BoYesNoEnum.tYES);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            GC.Collect();
        }

        public static void registerUDO( out string errorText)
        {
            errorText = null;
            string code = "UDO_F_BDO_TAXP_D"; //20 characters (must include at least one alphabetical character).
            Dictionary<string, object> formProperties;

            formProperties = new Dictionary<string, object>();
            formProperties.Add("Name", "Profit Tax Accrual"); //100 characters
            formProperties.Add("TableName", "BDO_TAXP");
            formProperties.Add("ObjectType", SAPbobsCOM.BoUDOObjType.boud_Document);
            formProperties.Add("CanCancel", SAPbobsCOM.BoYesNoEnum.tYES);
            formProperties.Add("CanClose", SAPbobsCOM.BoYesNoEnum.tYES);
            formProperties.Add("CanCreateDefaultForm", SAPbobsCOM.BoYesNoEnum.tYES);
            formProperties.Add("CanDelete", SAPbobsCOM.BoYesNoEnum.tNO);
            formProperties.Add("CanFind", SAPbobsCOM.BoYesNoEnum.tYES);
            formProperties.Add("CanYearTransfer", SAPbobsCOM.BoYesNoEnum.tYES);
            formProperties.Add("ManageSeries", SAPbobsCOM.BoYesNoEnum.tYES);
            formProperties.Add("EnableEnhancedForm", SAPbobsCOM.BoYesNoEnum.tYES);

            List<Dictionary<string, object>> listFindColumns = new List<Dictionary<string, object>>();
            List<Dictionary<string, object>> listFormColumns = new List<Dictionary<string, object>>();
            List<Dictionary<string, object>> listChildTables = new List<Dictionary<string, object>>();
            List<Dictionary<string, object>> listEnhancedFormColumns = new List<Dictionary<string, object>>();

            Dictionary<string, object> fieldskeysMap;

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "U_prBase");
            fieldskeysMap.Add("ColumnDescription", "Profit Base"); //30 characters
            listFindColumns.Add(fieldskeysMap);
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "U_prBsDscr");
            fieldskeysMap.Add("ColumnDescription", "Profit Base Description"); //30 characters
            listFindColumns.Add(fieldskeysMap);
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "U_docDate");
            fieldskeysMap.Add("ColumnDescription", "Posting Date"); //30 characters
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "CreateDate");
            fieldskeysMap.Add("ColumnDescription", "Create Date"); //30 characters
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
            fieldskeysMap.Add("ColumnAlias", "Remark");
            fieldskeysMap.Add("ColumnDescription", "Remark"); //30 characters
            listFindColumns.Add(fieldskeysMap);

            formProperties.Add("FindColumns", listFindColumns);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("FormColumnAlias", "DocEntry");
            fieldskeysMap.Add("FormColumnDescription", "DocEntry"); //30 characters
            listFormColumns.Add(fieldskeysMap);

            formProperties.Add("FormColumns", listFormColumns);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("TableName", "BDO_TXP1");
            fieldskeysMap.Add("ObjectName", "BDO_TXP1"); //30 characters
            listChildTables.Add(fieldskeysMap);

            formProperties.Add("ChildTables", listChildTables);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ChildNumber", 1);
            fieldskeysMap.Add("ColumnAlias", "LineId");
            fieldskeysMap.Add("ColumnDescription", "Line ID");
            fieldskeysMap.Add("ColumnIsUsed", SAPbobsCOM.BoYesNoEnum.tYES);
            fieldskeysMap.Add("ColumnNumber", 1);
            fieldskeysMap.Add("Editable", SAPbobsCOM.BoYesNoEnum.tNO);
            listEnhancedFormColumns.Add(fieldskeysMap);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ChildNumber", 1);
            fieldskeysMap.Add("ColumnAlias", "U_txType");
            fieldskeysMap.Add("ColumnDescription", "Base Type");
            fieldskeysMap.Add("ColumnIsUsed", SAPbobsCOM.BoYesNoEnum.tYES);
            fieldskeysMap.Add("ColumnNumber", 2);
            fieldskeysMap.Add("Editable", SAPbobsCOM.BoYesNoEnum.tYES);
            listEnhancedFormColumns.Add(fieldskeysMap);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ChildNumber", 1);
            fieldskeysMap.Add("ColumnAlias", "U_amtTx");
            fieldskeysMap.Add("ColumnDescription", "Amount Taxable");
            fieldskeysMap.Add("ColumnIsUsed", SAPbobsCOM.BoYesNoEnum.tYES);
            fieldskeysMap.Add("ColumnNumber", 3);
            fieldskeysMap.Add("Editable", SAPbobsCOM.BoYesNoEnum.tYES);
            listEnhancedFormColumns.Add(fieldskeysMap);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ChildNumber", 1);
            fieldskeysMap.Add("ColumnAlias", "U_amtPrTx");
            fieldskeysMap.Add("ColumnDescription", "Profit Tax Amount");
            fieldskeysMap.Add("ColumnIsUsed", SAPbobsCOM.BoYesNoEnum.tYES);
            fieldskeysMap.Add("ColumnNumber", 4);
            fieldskeysMap.Add("Editable", SAPbobsCOM.BoYesNoEnum.tNO);
            listEnhancedFormColumns.Add(fieldskeysMap);

            formProperties.Add("EnhancedFormColumns", listEnhancedFormColumns);

            UDO.registerUDO( code, formProperties, out errorText);

            GC.Collect();
        }

        public static void addMenus()
        {
            SAPbouiCOM.MenuItem menuItem;
            SAPbouiCOM.MenuItem fatherMenuItem;
            SAPbouiCOM.MenuCreationParams oCreationPackage;

            try
            {
                oCreationPackage = (SAPbouiCOM.MenuCreationParams)Program.uiApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                fatherMenuItem = Program.uiApp.Menus.Item("1536");

                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                oCreationPackage.UniqueID = "TAX1536";
                oCreationPackage.String = BDOSResources.getTranslate("Tax");
                oCreationPackage.Position = -1;

                menuItem = fatherMenuItem.SubMenus.AddEx(oCreationPackage);
            }
            catch 
            {
                
            }

            try
            {
                fatherMenuItem = Program.uiApp.Menus.Item("TAX1536");
                // Add a pop-up menu item
                oCreationPackage = (SAPbouiCOM.MenuCreationParams)Program.uiApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                oCreationPackage.Checked = false;
                oCreationPackage.Enabled = true;
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "UDO_F_BDO_TAXP_D";
                oCreationPackage.String = BDOSResources.getTranslate("ProfitTaxAccrual");
                oCreationPackage.Position = -1;

                menuItem = fatherMenuItem.SubMenus.AddEx(oCreationPackage);
            }
            catch
            {

            }
        }

        public static void createFormItems(  SAPbouiCOM.Form oForm, out bool BubbleEvent, out string errorText)
        {
            BubbleEvent = true;
            errorText = null;

            profitTaxRate = ProfitTax.GetProfitTaxRate();
            ProfitTaxTypeIsSharing = ProfitTax.ProfitTaxTypeIsSharing();

            if (ProfitTaxTypeIsSharing == false)
            {
                errorText = BDOSResources.getTranslate("ProfitTaxSystem") + " " + BDOSResources.getTranslate("IsNot") + " " + BDOSResources.getTranslate("ProfitSharingTax");
                Program.uiApp.MessageBox(errorText);
                BubbleEvent = false;
            }         

            Dictionary<string, object> formItems;
            string itemName = "";

            int left_s = 6;
            int left_e = 127;
            int height = 15;
            int top = 6;
            int width_s = 121;
            int width_e = 148;
            int fontSize = 10;

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "PrBaseS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("TaxableObject"));
            formItems.Add("LinkTo", "PrBaseE");
            formItems.Add("FontSize", fontSize);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            bool multiSelection = false;
            string objectType = "UDO_F_BDO_PTBS_D";
            string uniqueID_lf_ProfitBaseCFL = "CFL_ProfitBase";
            FormsB1.addChooseFromList( oForm, multiSelection, objectType, uniqueID_lf_ProfitBaseCFL);

            formItems = new Dictionary<string, object>();
            itemName = "PrBaseE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXP");
            formItems.Add("Alias", "U_prBase");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", 30);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("ChooseFromListUID", uniqueID_lf_ProfitBaseCFL);
            formItems.Add("ChooseFromListAlias", "Code");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "PrBsDscr"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXP");
            formItems.Add("Alias", "U_prBsDscr");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e + 32);
            formItems.Add("Width", width_e - 32);
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

            //golden errow
            formItems = new Dictionary<string, object>();
            itemName = "prBaseLB"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            formItems.Add("Left", left_e - 20);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("LinkTo", "PrBaseE");
            formItems.Add("LinkedObjectType", objectType);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 1;            

            //მარჯვენა
            top = 6;
            left_s = 295;
            left_e = left_s + 121;

            formItems = new Dictionary<string, object>();
            itemName = "No.S"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s / 3);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Number"));
            formItems.Add("LinkTo", "SeriesC");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "SeriesC"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXP");
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
            itemName = "DocNumE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXP");
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
            itemName = "StatusS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Status"));
            formItems.Add("LinkTo", "StatusC");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "StatusC"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXP");
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
            itemName = "CreateDatS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("CreateDate"));
            formItems.Add("LinkTo", "CreateDatE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "CreateDatE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXP");
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
            formItems.Add("TableName", "@BDO_TAXP");
            formItems.Add("Alias", "U_docDate");
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
            
            left_s = 6;
            left_e = 127;
            top = top +  height;

            formItems = new Dictionary<string, object>();
            itemName = "addMTRB"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            formItems.Add("Left", left_s);
            formItems.Add("Width", 100);
            formItems.Add("Top", top + 35);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Add"));

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + 2 * height + 15;            
             
            //სარდაფი
            left_s = 6;
            left_e = 127;
            top = 100;

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

            formItems = new Dictionary<string, object>();
            itemName = "CreatorE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXP");
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

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

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
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "RemarksE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_TAXP");
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

            GC.Collect();
        }

        public static void reArrangeFormItems(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.Item oItem = null;
            int height = 15;
            int top = 6;
            top = top + height + 1;

            oItem = oForm.Items.Item("PrBaseS");
            oItem.Top = top;
            oItem = oForm.Items.Item("PrBsDscr");
            oItem.Top = top;            
            top = top + height + 1;
            

            top = 6;

            oItem = oForm.Items.Item("No.S");
            oItem.Top = top;
            oItem = oForm.Items.Item("SeriesC");
            oItem.Top = top;
            oItem = oForm.Items.Item("DocNumE");
            oItem.Top = top;
            top = top + height + 1;

            oItem = oForm.Items.Item("StatusS");
            oItem.Top = top;
            oItem = oForm.Items.Item("StatusC");
            oItem.Top = top;
            top = top + height + 1;

            oItem = oForm.Items.Item("CreateDatS");
            oItem.Top = top;
            oItem = oForm.Items.Item("CreateDatE");
            oItem.Top = top;
                        
            top = top + height + 1;

            oItem = oForm.Items.Item("DocDateS");
            oItem.Top = top;
            oItem = oForm.Items.Item("DocDateE");
            oItem.Top = top;
            top = top + height + 1;
            oItem = oForm.Items.Item("BDOSJrnEnS");
            oItem.Top = top;

            oItem = oForm.Items.Item("BDOSJEntLB");
            oItem.Top = top;
            oItem = oForm.Items.Item("BDOSJrnEnt");
            oItem.Top = top;
            oItem = oForm.Items.Item("BDOSJrnEnS");
            oItem.Top = top;
            top = top + 5 * height + 1;

            int wblMTRWidth = oForm.Width;// - 28
            oItem = oForm.Items.Item("0_U_G");
            oItem.Top = top;
            oItem.Width = wblMTRWidth - 15;
            oItem.Height = oForm.Height / 3;

            oItem = oForm.Items.Item("0_U_FD");
            oItem.Top = top - 10;
            oItem = oForm.Items.Item("U_RC");
            oItem.Top = top - 5;
            oItem.Width = wblMTRWidth;
            oItem.Height = oForm.Height / 3 + 10;

            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("0_U_G").Specific));
            SAPbouiCOM.Columns oColumns = oMatrix.Columns;
            SAPbouiCOM.Column oColumn;

            oColumn = oMatrix.Columns.Item("#");
            oColumn.Visible = false;

            oColumn = oMatrix.Columns.Item("C_0_1");
            oColumn.Width = 20;
            oColumn.TitleObject.Caption = "#";

            wblMTRWidth = wblMTRWidth - 20 - 1;
            oColumn = oMatrix.Columns.Item("C_0_2");
            oColumn.Width = wblMTRWidth / 3;
            oColumn = oMatrix.Columns.Item("C_0_3");
            oColumn.Width = wblMTRWidth / 3;
            oColumn = oMatrix.Columns.Item("C_0_4");
            oColumn.Width = wblMTRWidth / 3;
            oColumn.Editable = false;

            //სარდაფი
            top = oForm.Items.Item("U_RC").Top + oForm.Items.Item("U_RC").Height + 20;

            oItem = oForm.Items.Item("CreatorS");
            oItem.Top = top;
            oItem = oForm.Items.Item("CreatorE");
            oItem.Top = top;
            top = top + height + 1;

            oItem = oForm.Items.Item("RemarksS");
            oItem.Top = top;
            oItem = oForm.Items.Item("RemarksE");
            oItem.Top = top;
            
            top = top + 4*height;
            
            ////ღილაკები
            //oItem = oForm.Items.Item("1");
            //oItem.Top = oForm.ClientHeight - 25;

            //oItem = oForm.Items.Item("2");
            //oItem.Top = oForm.ClientHeight - 25;

            oForm.Items.Item("1").Top = top;
            oForm.Items.Item("2").Top = top;

        }

        public static void setSizeForm( SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            try
            {
                oForm.ClientHeight = Program.uiApp.Desktop.Width / 6;

                oForm.Height = Program.uiApp.Desktop.Width / 4;

                oForm.Left = (Program.uiApp.Desktop.Width - oForm.Width) / 2;
                oForm.Top = (Program.uiApp.Desktop.Height - oForm.Height) / 2;
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
        }

        public static void resizeForm( SAPbouiCOM.Form oForm, out string errorText)
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

        public static void setVisibleFormItems(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            try
            {
                SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("0_U_G").Specific));
                oMatrix.Columns.Item("C_0_3").Editable = (oForm.Mode != SAPbouiCOM.BoFormMode.fm_OK_MODE);
                 //matrix.Col   (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                oForm.Items.Item("BDOSJrnEnt").Enabled = false;
                string docEntry = oForm.DataSources.DBDataSources.Item("@BDO_TAXP").GetValue("DocEntry", 0).Trim();
                bool docEntryIsEmpty = string.IsNullOrEmpty(docEntry);

                oForm.Items.Item("DocDateE").Enabled = (docEntryIsEmpty == true);
                oForm.Items.Item("PrBaseE").Enabled = (docEntryIsEmpty == true);

                oForm.Items.Item("0_U_G").Enabled = (docEntryIsEmpty == true);
                oForm.Items.Item("DocNumE").Enabled = (docEntryIsEmpty == true);
                oForm.Items.Item("CreateDatE").Enabled = (docEntryIsEmpty == true);
                oForm.Items.Item("StatusC").Enabled = (docEntryIsEmpty == true);
                oForm.Items.Item("SeriesC").Enabled = (docEntryIsEmpty == true);
                oForm.Items.Item("PrBsDscr").Enabled = (docEntryIsEmpty == true);
                oForm.Items.Item("CreatorE").Enabled = (docEntryIsEmpty == true);
                
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

        public static void formDataLoad( SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            try
            {
                string DocEntry = oForm.DataSources.DBDataSources.Item("@BDO_TAXP").GetValue("DocEntry", 0).Trim();

                setVisibleFormItems(oForm, out errorText);

                // გატარებები
                SAPbouiCOM.DBDataSource DocDBSourceTAXP = oForm.DataSources.DBDataSources.Item(0);
                string Ref1 = DocDBSourceTAXP.GetValue("DocEntry", 0);
                string Ref2 = "UDO_F_BDO_TAXP_D";

                SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string query = "SELECT " +
                                "*  " +
                                "FROM \"OJDT\"  " +
                                "WHERE \"StornoToTr\" IS NULL " +
                                "AND \"Ref1\" = '" + Ref1 + "' " +
                                "AND \"Ref2\" = '" + Ref2 + "' ";
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

        public static void chooseFromList( SAPbouiCOM.Form oForm, SAPbouiCOM.IChooseFromListEvent oCFLEvento, string itemUID, bool beforeAction, out string errorText)
        {
            errorText = null;

            NumberFormatInfo Nfi = new NumberFormatInfo() { NumberDecimalSeparator = "." };

            try
            {
                string sCFL_ID = oCFLEvento.ChooseFromListUID;
                SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(sCFL_ID);

                if (beforeAction == false)
                {
                    SAPbouiCOM.DataTable oDataTable = null;
                    oDataTable = oCFLEvento.SelectedObjects;

                    if (oDataTable != null)
                    {
                        if (sCFL_ID == "CFL_ProfitBase")
                        {
                            string ProfitBaseCode = Convert.ToString(oDataTable.GetValue("Code", 0));
                            string ProfitBaseName = Convert.ToString(oDataTable.GetValue("Name", 0));

                            oForm.DataSources.DBDataSources.Item("@BDO_TAXP").SetValue("U_prBase", 0, ProfitBaseCode);
                            oForm.DataSources.DBDataSources.Item("@BDO_TAXP").SetValue("U_prBsDscr", 0, ProfitBaseName);

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

        public static void fillMatrixRow(  SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            oForm.Freeze(true);
            NumberFormatInfo Nfi = new NumberFormatInfo() { NumberDecimalSeparator = "." };

            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("0_U_G").Specific));
            SAPbouiCOM.CellPosition cellPos = oMatrix.GetCellFocus();
            if (cellPos == null)
            {
                errorText = "Error";
            }

            try
            {
                profitTaxRate = ProfitTax.GetProfitTaxRate();
                ProfitTaxTypeIsSharing = ProfitTax.ProfitTaxTypeIsSharing();

                SAPbouiCOM.EditText oEditText = oMatrix.Columns.Item("C_0_3").Cells.Item(cellPos.rowIndex).Specific;
                string amtTxVal = oEditText.Value;
                decimal amtTx = Convert.ToDecimal(amtTxVal, CultureInfo.InvariantCulture);
                decimal amtPrTx = CommonFunctions.roundAmountByGeneralSettings(amtTx / (100 - profitTaxRate) * profitTaxRate, "Sum");
                oEditText = oMatrix.Columns.Item("C_0_4").Cells.Item(cellPos.rowIndex).Specific;
                oEditText.Value = FormsB1.ConvertDecimalToString(Convert.ToDecimal(amtPrTx, CultureInfo.InvariantCulture));
            }
            catch
            { }
            finally
            {
                GC.Collect();
                oForm.Update();
                oForm.Freeze(false);
            }
        }

        public static void fillMatrixRowCount(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            oForm.Freeze(true);

            int lineNum = 0;

            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("0_U_G").Specific));
            try
            {
                for (int i = 1; i < oMatrix.RowCount + 1; i++)
                {
                    lineNum = i;
                    if (Program.removeLineTrans == true & removeRecordRow != 0 & i > removeRecordRow)
                    {
                        lineNum--;
                    }
                    oMatrix.Columns.Item("C_0_1").Cells.Item(i).Specific.Value = lineNum;
                }
            }
            catch
            { }
            finally
            {
                GC.Collect();
                oForm.Update();
                oForm.Freeze(false);
            }
        }

        public static void addMatrixRow( SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            oForm.Freeze(true);

            try
            {
                SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("0_U_G").Specific));
                //if (oMatrix.RowCount > 0)
                //{
                    //string txType = oMatrix.Columns.Item("C_0_2").Cells.Item(oMatrix.RowCount).Specific.Value;
                    //if (txType != "")
                    //{
                        oMatrix.AddRow();
                        
                        oMatrix.Columns.Item("C_0_3").Cells.Item(oMatrix.RowCount).Specific.Value = 0;
                        oMatrix.Columns.Item("C_0_4").Cells.Item(oMatrix.RowCount).Specific.Value = 0;
                        oMatrix.Columns.Item("C_0_1").Cells.Item(oMatrix.RowCount).Specific.Value = oMatrix.RowCount;
                    //}
                //}
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

        public static void delMatrixRow( SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            oForm.Freeze(true);

            try
            {
                SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("txpMTR").Specific));
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
                        oForm.DataSources.DBDataSources.Item("@BDO_TXP1").RemoveRecord(row - deletedRowCount);
                        firstRow = row;
                    }
                }

                oMatrix.LoadFromDataSource();
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE && deletedRowCount > 0)
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                }
                
                for (int i = 1; i < oMatrix.RowCount + 1; i++)
                {
                    oMatrix.Columns.Item("LineID").Cells.Item(i).Specific.Value = i;

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

        public static void JrnEntry( string DocEntry, string DocNum, DateTime DocDate, DataTable JrnLinesDT, DataTable reLines, out string errorText)
        {
            errorText = null;

            try
            {
                JournalEntry.JrnEntry( DocEntry, "UDO_F_BDO_TAXP_D", "Profit Tax Accrual:  " + DocNum, DocDate, JrnLinesDT,  out errorText);

                if (errorText != null)
                {
                    return;
                }

                for (int i = 0; i < reLines.Rows.Count; i++)
                {
                    reLines.Rows[i]["DocEntry"] = string.IsNullOrEmpty(DocEntry) ? 0 : Convert.ToInt32(DocEntry); //DocEntry == "" ? 0 : Convert.ToInt32(DocEntry);
                    reLines.Rows[i]["DocNum"] = string.IsNullOrEmpty(DocNum) ? 0 : Convert.ToInt32(DocNum); //DocNum.ToString();
                    reLines.Rows[i]["docDate"] = DocDate;
                }

                ProfitTax.AddRecord( reLines, "UDO_F_BDO_TAXP_D", "Profit Tax Accrual: " + DocNum, out errorText);
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
        }

        public static DataTable createAdditionalEntries( SAPbouiCOM.Form oForm, SAPbobsCOM.GeneralData oGeneralData, DataTable DTSource, string DocCurrency, out DataTable reLines, decimal DocRate)
        {
            DataTable jeLines = JournalEntry.JournalEntryTable();

            reLines = ProfitTax.ProfitTaxTable();
            DataRow reLinesRow = null;
            DataTable AccountTable = CommonFunctions.GetOACTTable();
            SAPbouiCOM.DBDataSource DBDataSourceTable = null;
            int JEcount = 0;

            if (oForm == null)
            {
                JEcount = DTSource.Rows.Count;
            }
            else
            {
                DBDataSourceTable = oForm.DataSources.DBDataSources.Item("@BDO_TXP1");
                JEcount = DBDataSourceTable.Size;
            }




            string prBase = oForm.DataSources.DBDataSources.Item("@BDO_TAXP").GetValue("U_prBase", 0).Trim();
            string DebitAccount="";
            string CreditAccount = "";

            for (int i = 0; i < JEcount; i++)
            {
                string TxType = CommonFunctions.getChildOrDbDataSourceValue(DBDataSourceTable, null, DTSource, "U_txType", i).ToString();
                
                if (TxType != "")
                {
                    if (TxType != "Accrual")
                    {
                        DebitAccount = CommonFunctions.getOADM( "U_BDO_TaxAcc").ToString();                        
                        CreditAccount = CommonFunctions.getOADM( "U_BDO_CapAcc").ToString();                        
                    }
                    else
                    {
                        DebitAccount = CommonFunctions.getOADM( "U_BDO_CapAcc").ToString();
                        CreditAccount = CommonFunctions.getOADM( "U_BDO_TaxAcc").ToString();                        
                    }


                    decimal TaxAmount =  Convert.ToDecimal(CommonFunctions.getChildOrDbDataSourceValue(DBDataSourceTable, null, DTSource, "U_AmtPrTx", i),CultureInfo.InvariantCulture);
                    decimal AmtTx = Convert.ToDecimal(CommonFunctions.getChildOrDbDataSourceValue(DBDataSourceTable, null, DTSource, "U_AmtTx", i), CultureInfo.InvariantCulture);

                    DocCurrency = DocCurrency == CommonFunctions.getLocalCurrency() ? "" : DocCurrency;
                    decimal TaxAmountFC = DocCurrency == "" ? 0 : TaxAmount / DocRate;

                    JournalEntry.AddJournalEntryRow(AccountTable, jeLines, "Full", DebitAccount, CreditAccount, TaxAmount, TaxAmountFC, DocCurrency, "", "", "", "", "", "", "", "");

                    reLinesRow = reLines.Rows.Add();

                    reLinesRow["debitAccount"] = DebitAccount;
                    reLinesRow["creditAccount"] = CreditAccount;
                    reLinesRow["prBase"] = prBase;
                    reLinesRow["txType"] = TxType;
                    reLinesRow["amtTx"] = AmtTx;
                    reLinesRow["amtPrTx"] = TaxAmount;
                }
            }

            return jeLines;

        }

         public static void uiApp_FormDataEvent(  ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(BusinessObjectInfo.FormTypeEx, Program.currentFormCount);

            if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD)
            {
                if (BusinessObjectInfo.BeforeAction == true)
                {
                    // მოგების გადასახადი
                    if (oForm.DataSources.DBDataSources.Item("@BDO_TAXP").GetValue("U_docDate", 0) == "")
                    {
                        Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("PostingDate") + " " + BDOSResources.getTranslate("YouCantLeaveEmpty"));
                        Program.uiApp.MessageBox(BDOSResources.getTranslate("OperationUnsuccesfullSeeLog"));
                        BubbleEvent = false;
                    }

                    bool TaxAccountsIsEmpty = ProfitTax.TaxAccountsIsEmpty();
                    if (TaxAccountsIsEmpty == true || oForm.DataSources.DBDataSources.Item("@BDO_TAXP").GetValue("U_prBase", 0) == "")
                    {
                        Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("TaxAccounts") + ", " + BDOSResources.getTranslate("TaxableObject") + " " + BDOSResources.getTranslate("YouCantLeaveEmpty"));
                        Program.uiApp.MessageBox(BDOSResources.getTranslate("OperationUnsuccesfullSeeLog"));
                        BubbleEvent = false;
                    }

                    SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("0_U_G").Specific));
                    for (int i = 1; i < oMatrix.RowCount + 1; i++)
                    {
                        string AmtTx = FormsB1.cleanStringOfNonDigits( oMatrix.Columns.Item("C_0_3").Cells.Item(i).Specific.Value).ToString();
                        string AmtPrTx = FormsB1.cleanStringOfNonDigits( oMatrix.Columns.Item("C_0_4").Cells.Item(i).Specific.Value).ToString();
                        if (Convert.ToDouble(AmtTx) <= 0 || Convert.ToDouble(AmtPrTx) <= 0)
                        {
                            Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("Amount") + " " + BDOSResources.getTranslate("YouCantLeaveEmpty"));
                            Program.uiApp.MessageBox(BDOSResources.getTranslate("OperationUnsuccesfullSeeLog"));
                            BubbleEvent = false;
                            break;
                        }

                    }
                }
                if (BubbleEvent)
                {
                JournalEntryTransaction(  oForm, BusinessObjectInfo.ActionSuccess, BusinessObjectInfo.BeforeAction, out BubbleEvent, out errorText);
                }              
            }

            if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE & BusinessObjectInfo.BeforeAction == false & BusinessObjectInfo.ActionSuccess == true)
            {
                if (Program.cancellationTrans == true & Program.canceledDocEntry != 0)
                {
                    cancellation( oForm, Program.canceledDocEntry, out errorText);
                    Program.canceledDocEntry = 0;
                }
            }

                if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD & BusinessObjectInfo.BeforeAction == false)
                {
                    formDataLoad( oForm, out errorText);
                }
        }

         public static void JournalEntryTransaction(  SAPbouiCOM.Form oForm, bool ActionSuccess, bool BeforeAction, out bool BubbleEvent, out string errorText)
         {
             errorText = null;
             BubbleEvent = true;


             if (ActionSuccess != BeforeAction)
             {
                 SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("0_U_G").Specific));
                 oMatrix.FlushToDataSource();

                 //დოკუმენტის გატარების დროს გატარდეს ბუღლტრული გატარება
                 SAPbouiCOM.DBDataSource DocDBSourcePAYR = oForm.DataSources.DBDataSources.Item(0);

                 if (DocDBSourcePAYR.GetValue("CANCELED", 0) == "N")
                 {
                     string DocEntry = DocDBSourcePAYR.GetValue("DocEntry", 0);
                     //string DocCurrency = DocDBSourcePAYR.GetValue("DocCur", 0);
                     //decimal DocRate = FormsB1.cleanStringOfNonDigits( DocDBSourcePAYR.GetValue("DocRate", 0));
                     string DocNum = DocDBSourcePAYR.GetValue("DocNum", 0);
                     DateTime DocDate = DateTime.ParseExact(DocDBSourcePAYR.GetValue("U_docDate", 0), "yyyyMMdd", CultureInfo.InvariantCulture);

                     CommonFunctions.StartTransaction();

                     Program.JrnLinesGlobal = new DataTable();
                     DataTable reLines = null;
                     DataTable JrnLinesDT = createAdditionalEntries( oForm, null, null, "", out reLines, 0);

                     JrnEntry( DocEntry, DocNum, DocDate, JrnLinesDT, reLines, out errorText);
                     if (errorText != null)
                     {
                         Program.uiApp.MessageBox(errorText);
                         BubbleEvent = false;
                     }
                     else
                     {
                         if (ActionSuccess == false)
                         {
                             Program.JrnLinesGlobal = JrnLinesDT;
                         }
                     }

                     if (Program.oCompany.InTransaction)
                     {
                         //თუ დოკუმენტი გატარდა, მერე ვაკეთებს ბუღალტრულ გატარებას
                         if (ActionSuccess == true & BeforeAction == false)
                         {
                             CommonFunctions.EndTransaction( SAPbobsCOM.BoWfTransOpt.wf_Commit);
                         }
                         else
                         {
                             CommonFunctions.EndTransaction( SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                         }
                     }
                     else
                     {
                         Program.uiApp.MessageBox("ტრანზაქციის დასრულებს შეცდომა");
                         BubbleEvent = false;

                     }
                 }


             }
         }


        public static void cancellation(  SAPbouiCOM.Form oForm, int docEntry, out string errorText)
        {
            errorText = null;

            try
            {
                JournalEntry.cancellation( oForm, docEntry, "UDO_F_BDO_TAXP_D", out errorText);
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

        public static void uiApp_MenuEvent(  ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent, out string errorText)
        {
            errorText = null;
            BubbleEvent = true;

            //----------------------------->Preview <-----------------------------
            try
            {
                if (pVal.BeforeAction && pVal.MenuUID == "PreviewUDOJrE")
                {
                    SAPbouiCOM.Form oDocForm = Program.uiApp.Forms.ActiveForm;
                    JournalEntryTransaction(  oDocForm, false, true, out BubbleEvent, out errorText);

                    if (BubbleEvent)
                    {
                        SAPbouiCOM.Form oJournalForm = Program.uiApp.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_JournalPosting, "", "");
                    }
                }
            }
            catch (Exception ex)
            {
                Program.uiApp.MessageBox(ex.ToString(), 1, "", "");
            }
        }


        public static void uiApp_ItemEvent(  string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                /*if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK & pVal.BeforeAction == true)
                {
                    BubbleEvent = false;
                }*/
                
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD & pVal.BeforeAction == true)
                {
                    createFormItems( oForm, out BubbleEvent, out errorText);
                    Program.FORM_LOAD_FOR_VISIBLE = true;
                    Program.FORM_LOAD_FOR_ACTIVATE = true;
                    setVisibleFormItems(oForm, out errorText);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE & pVal.BeforeAction == false)
                {
                    oForm.Freeze(true);
                    if (Program.FORM_LOAD_FOR_VISIBLE == true)
                    {
                        setSizeForm( oForm, out errorText);
                        oForm.Title = BDOSResources.getTranslate("ProfitTaxAccrual");
                        Program.FORM_LOAD_FOR_VISIBLE = false;
                    }
                    //setVisibleFormItems(oForm, out errorText);
                    oForm.Freeze(false);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_RESIZE & pVal.BeforeAction == false)
                {
                    oForm.Freeze(true);
                    resizeForm( oForm, out errorText);
                    oForm.Freeze(false);
                }

                if (pVal.ItemUID == "PrBaseE" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_KEY_DOWN && pVal.BeforeAction == false)
                {
                    oForm.Freeze(true);
                    if (string.IsNullOrEmpty(oForm.Items.Item("PrBaseE").Specific.Value))
                    {
                        oForm.Items.Item("PrBsDscr").Specific.Value = "";
                        string uniqueID_lf_ProfitBaseCFL = "CFL_ProfitBase";
                        oForm.Items.Item("PrBaseE").Specific.ChooseFromListUID = uniqueID_lf_ProfitBaseCFL;
                        oForm.Items.Item("PrBaseE").Specific.ChooseFromListAlias = "Code";
                    }
                    oForm.Freeze(false);
                }

                if (pVal.ItemUID == "PrBaseE" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                {
                    SAPbouiCOM.IChooseFromListEvent oCFLEvento = null;
                    oCFLEvento = ((SAPbouiCOM.IChooseFromListEvent)(pVal));

                    chooseFromList( oForm, oCFLEvento, pVal.ItemUID, pVal.BeforeAction, out errorText);
                }

                if (pVal.ItemUID == "addMTRB" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && oForm.Mode!=SAPbouiCOM.BoFormMode.fm_OK_MODE && pVal.BeforeAction == false)
                {
                    addMatrixRow( oForm, out errorText);
                }

                if (pVal.ItemUID == "0_U_G")
                {
                    if (pVal.ItemUID == "0_U_G" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_VALIDATE & pVal.BeforeAction == false & pVal.InnerEvent == false)
                    {
                        fillMatrixRow( oForm, out errorText);
                    }                  

                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK)
                    {
                        removeRecordRow = pVal.Row;
                    }

                    if (Program.removeLineTrans == true & removeRecordRow != 0 & pVal.EventType == SAPbouiCOM.BoEventTypes.et_LOST_FOCUS)
                    {
                        fillMatrixRowCount(oForm, out errorText);
                        Program.removeLineTrans = false;
                        removeRecordRow = 0;
                    }
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE & pVal.BeforeAction == false)
                {
                    if (Program.FORM_LOAD_FOR_ACTIVATE == true)
                    {
                        oForm.Freeze(true);
                        formDataLoad( oForm, out errorText);
                        oForm.Freeze(false);
                        Program.FORM_LOAD_FOR_ACTIVATE = false;
                    }
                }
            }
        }
    }
}
