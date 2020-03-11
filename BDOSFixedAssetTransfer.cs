using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;

namespace BDO_Localisation_AddOn
{
    class BDOSFixedAssetTransfer
    {
        public static bool openFormEvent = false;

        public static void createDocumentUDO(out string errorText)
        {
            errorText = null;
            string tableName = "BDOSFASTRD";
            string description = "Fixed Asset Transfer Document";

            SAPbobsCOM.UserObjectsMD oUserObjectsMD = null;
            oUserObjectsMD = (SAPbobsCOM.UserObjectsMD)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectsMD);

            int result = UDO.addUserTable(tableName, description, SAPbobsCOM.BoUTBTableType.bott_Document, out errorText);

            if (result != 0)
            {
                return;
            }

            Dictionary<string, object> fieldskeysMap;

            fieldskeysMap = new Dictionary<string, object>(); // DocDate 
            fieldskeysMap.Add("Name", "DocDate");
            fieldskeysMap.Add("TableName", "BDOSFASTRD");
            fieldskeysMap.Add("Description", "Posting Date");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Date);
            fieldskeysMap.Add("EditSize", 15);
            fieldskeysMap.Add("Mandatory", SAPbobsCOM.BoYesNoEnum.tYES);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //ბიზნესპარტნიორი (კოდი)
            fieldskeysMap.Add("Name", "CardCode");
            fieldskeysMap.Add("TableName", "BDOSFASTRD");
            fieldskeysMap.Add("Description", "Card Code");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 15);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //ბიზნესპარტნიორი (სახელი)
            fieldskeysMap.Add("Name", "CardName");
            fieldskeysMap.Add("TableName", "BDOSFASTRD");
            fieldskeysMap.Add("Description", "Card Name");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 100);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //From Location (კოდი)
            fieldskeysMap.Add("Name", "FLocCode");
            fieldskeysMap.Add("TableName", "BDOSFASTRD");
            fieldskeysMap.Add("Description", "From Location Code");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 11);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //From Location (სახელი)
            fieldskeysMap.Add("Name", "FLocName");
            fieldskeysMap.Add("TableName", "BDOSFASTRD");
            fieldskeysMap.Add("Description", "From Location Name");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 100);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //To Location (კოდი)
            fieldskeysMap.Add("Name", "TLocCode");
            fieldskeysMap.Add("TableName", "BDOSFASTRD");
            fieldskeysMap.Add("Description", "To Location Code");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 11);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //To Location (სახელი)
            fieldskeysMap.Add("Name", "TLocName");
            fieldskeysMap.Add("TableName", "BDOSFASTRD");
            fieldskeysMap.Add("Description", "To Location Name");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 100);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // From Employee (კოდი)
            fieldskeysMap.Add("Name", "FEmplID");
            fieldskeysMap.Add("TableName", "BDOSFASTRD");
            fieldskeysMap.Add("Description", "From Employee");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 11);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // From Employee (სახელი)
            fieldskeysMap.Add("Name", "FEmplName");
            fieldskeysMap.Add("TableName", "BDOSFASTRD");
            fieldskeysMap.Add("Description", "From Employee Name");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 100);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // To Employee (კოდი)
            fieldskeysMap.Add("Name", "TEmplID");
            fieldskeysMap.Add("TableName", "BDOSFASTRD");
            fieldskeysMap.Add("Description", "To Employee");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 11);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // To Employee (სახელი)
            fieldskeysMap.Add("Name", "TEmplName");
            fieldskeysMap.Add("TableName", "BDOSFASTRD");
            fieldskeysMap.Add("Description", "To Employee Name");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 100);

            UDO.addUserTableFields(fieldskeysMap, out errorText);


            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "ChngDstRl1");
            fieldskeysMap.Add("TableName", "BDOSFASTRD");
            fieldskeysMap.Add("Description", "Change Distr.Rule 1");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 1);
            fieldskeysMap.Add("DefaultValue", "N");

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "ChngDstRl2");
            fieldskeysMap.Add("TableName", "BDOSFASTRD");
            fieldskeysMap.Add("Description", "Change Distr.Rule 2");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 1);
            fieldskeysMap.Add("DefaultValue", "N");

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "ChngDstRl3");
            fieldskeysMap.Add("TableName", "BDOSFASTRD");
            fieldskeysMap.Add("Description", "Change Distr.Rule 3");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 1);
            fieldskeysMap.Add("DefaultValue", "N");

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "ChngDstRl4");
            fieldskeysMap.Add("TableName", "BDOSFASTRD");
            fieldskeysMap.Add("Description", "Change Distr.Rule 4");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 1);
            fieldskeysMap.Add("DefaultValue", "N");

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "ChngDstRl5");
            fieldskeysMap.Add("TableName", "BDOSFASTRD");
            fieldskeysMap.Add("Description", "Change Distr.Rule 5");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 1);
            fieldskeysMap.Add("DefaultValue", "N");

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "DistrRule1");
            fieldskeysMap.Add("TableName", "BDOSFASTRD");
            fieldskeysMap.Add("Description", "Distr.Rule 1");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "DistrRule2");
            fieldskeysMap.Add("TableName", "BDOSFASTRD");
            fieldskeysMap.Add("Description", "Distr.Rule 2");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "DistrRule3");
            fieldskeysMap.Add("TableName", "BDOSFASTRD");
            fieldskeysMap.Add("Description", "Distr.Rule 3");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "DistrRule4");
            fieldskeysMap.Add("TableName", "BDOSFASTRD");
            fieldskeysMap.Add("Description", "Distr.Rule 4");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "DistrRule5");
            fieldskeysMap.Add("TableName", "BDOSFASTRD");
            fieldskeysMap.Add("Description", "Distr.Rule 5");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "PrjCode");
            fieldskeysMap.Add("TableName", "BDOSFASTRD");
            fieldskeysMap.Add("Description", "Project");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);


            //ცხრილური ნაწილი
            tableName = "BDOSFASTR1";
            description = "Fixed Asset Transfer Child1";

            result = UDO.addUserTable(tableName, description, SAPbobsCOM.BoUTBTableType.bott_DocumentLines, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // Item (კოდი)
            fieldskeysMap.Add("Name", "ItemCode");
            fieldskeysMap.Add("TableName", "BDOSFASTR1");
            fieldskeysMap.Add("Description", "Item Code");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // Item (სახელი)
            fieldskeysMap.Add("Name", "ItemName");
            fieldskeysMap.Add("TableName", "BDOSFASTR1");
            fieldskeysMap.Add("Description", "Item Name");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 100);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // Asset Serial Number
            fieldskeysMap.Add("Name", "SerNo");
            fieldskeysMap.Add("TableName", "BDOSFASTR1");
            fieldskeysMap.Add("Description", "Asset Serial Number");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 254);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "Quantity"); //რაოდენობა
            fieldskeysMap.Add("TableName", "BDOSFASTR1");
            fieldskeysMap.Add("Description", "Quantity");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Quantity);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "APC"); //APC
            fieldskeysMap.Add("TableName", "BDOSFASTR1");
            fieldskeysMap.Add("Description", "Acquisition and Production Cost");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //From Location (კოდი)
            fieldskeysMap.Add("Name", "FLocCode");
            fieldskeysMap.Add("TableName", "BDOSFASTR1");
            fieldskeysMap.Add("Description", "From Location Code");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 11);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //From Location (სახელი)
            fieldskeysMap.Add("Name", "FLocName");
            fieldskeysMap.Add("TableName", "BDOSFASTR1");
            fieldskeysMap.Add("Description", "From Location Name");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 100);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //To Location (კოდი)
            fieldskeysMap.Add("Name", "TLocCode");
            fieldskeysMap.Add("TableName", "BDOSFASTR1");
            fieldskeysMap.Add("Description", "To Location Code");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 11);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //To Location (სახელი)
            fieldskeysMap.Add("Name", "TLocName");
            fieldskeysMap.Add("TableName", "BDOSFASTR1");
            fieldskeysMap.Add("Description", "To Location Name");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 100);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // From Employee (კოდი)
            fieldskeysMap.Add("Name", "FEmplID");
            fieldskeysMap.Add("TableName", "BDOSFASTR1");
            fieldskeysMap.Add("Description", "From Employee");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 11);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // From Employee (სახელი)
            fieldskeysMap.Add("Name", "FEmplName");
            fieldskeysMap.Add("TableName", "BDOSFASTR1");
            fieldskeysMap.Add("Description", "From Employee Name");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 100);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // To Employee (კოდი)
            fieldskeysMap.Add("Name", "TEmplID");
            fieldskeysMap.Add("TableName", "BDOSFASTR1");
            fieldskeysMap.Add("Description", "To Employee");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 11);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // To Employee (სახელი)
            fieldskeysMap.Add("Name", "TEmplName");
            fieldskeysMap.Add("TableName", "BDOSFASTR1");
            fieldskeysMap.Add("Description", "To Employee Name");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 100);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // Uom (კოდი)
            fieldskeysMap.Add("Name", "UomCode");
            fieldskeysMap.Add("TableName", "BDOSFASTR1");
            fieldskeysMap.Add("Description", "Uom Code");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 20);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // Uom (სახელი)
            fieldskeysMap.Add("Name", "UomName");
            fieldskeysMap.Add("TableName", "BDOSFASTR1");
            fieldskeysMap.Add("Description", "Uom Name");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 100);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // Remark 
            fieldskeysMap.Add("Name", "Remark");
            fieldskeysMap.Add("TableName", "BDOSFASTR1");
            fieldskeysMap.Add("Description", "Remark");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 254);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            GC.Collect();

        }

        public static void registerUDO(out string errorText)
        {
            errorText = null;
            string code = "UDO_F_BDOSFASTRD_D"; //20 characters (must include at least one alphabetical character).
            Dictionary<string, object> formProperties;

            formProperties = new Dictionary<string, object>();
            formProperties.Add("Name", "Fixed Asset Transfer Document"); //100 characters
            formProperties.Add("TableName", "BDOSFASTRD");
            formProperties.Add("ObjectType", SAPbobsCOM.BoUDOObjType.boud_Document);
            formProperties.Add("CanCancel", SAPbobsCOM.BoYesNoEnum.tYES);
            formProperties.Add("CanClose", SAPbobsCOM.BoYesNoEnum.tYES);
            formProperties.Add("CanCreateDefaultForm", SAPbobsCOM.BoYesNoEnum.tYES);
            formProperties.Add("CanDelete", SAPbobsCOM.BoYesNoEnum.tNO);
            formProperties.Add("CanFind", SAPbobsCOM.BoYesNoEnum.tYES);
            formProperties.Add("CanYearTransfer", SAPbobsCOM.BoYesNoEnum.tYES);
            formProperties.Add("ManageSeries", SAPbobsCOM.BoYesNoEnum.tNO);
            formProperties.Add("CanLog", SAPbobsCOM.BoYesNoEnum.tYES);

            //string fatherMenuID = "9201";
            //SAPbouiCOM.MenuItem fatherMenuItem = Program.uiApp.Menus.Item(fatherMenuID);
            //formProperties.Add("MenuItem", SAPbobsCOM.BoYesNoEnum.tYES);
            //formProperties.Add("FatherMenuID", fatherMenuID);
            //formProperties.Add("MenuUID", code);
            //formProperties.Add("MenuCaption", BDOSResources.getTranslate("FixedAssetTransferDocument"));
            //formProperties.Add("Position", fatherMenuItem.SubMenus.Count - 1);

            List<Dictionary<string, object>> listFindColumns = new List<Dictionary<string, object>>();
            List<Dictionary<string, object>> listFormColumns = new List<Dictionary<string, object>>();
            List<Dictionary<string, object>> listChildTables = new List<Dictionary<string, object>>();

            Dictionary<string, object> fieldskeysMap;

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "U_DocDate");
            fieldskeysMap.Add("ColumnDescription", "Posting Date");
            listFindColumns.Add(fieldskeysMap);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "U_CardCode");
            fieldskeysMap.Add("ColumnDescription", "Card Code");
            listFindColumns.Add(fieldskeysMap);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "U_CardName");
            fieldskeysMap.Add("ColumnDescription", "Card Name");
            listFindColumns.Add(fieldskeysMap);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "U_FLocCode");
            fieldskeysMap.Add("ColumnDescription", "From Location Code");
            listFindColumns.Add(fieldskeysMap);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "U_FLocName");
            fieldskeysMap.Add("ColumnDescription", "From Location Name");
            listFindColumns.Add(fieldskeysMap);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "U_TLocCode");
            fieldskeysMap.Add("ColumnDescription", "To Location Code");
            listFindColumns.Add(fieldskeysMap);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "U_TLocName");
            fieldskeysMap.Add("ColumnDescription", "To Location Name");
            listFindColumns.Add(fieldskeysMap);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "U_FEmplID");
            fieldskeysMap.Add("ColumnDescription", "From Employee");
            listFindColumns.Add(fieldskeysMap);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "U_FEmplName");
            fieldskeysMap.Add("ColumnDescription", "From Employee Name");
            listFindColumns.Add(fieldskeysMap);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "U_TEmplID");
            fieldskeysMap.Add("ColumnDescription", "To Employee");
            listFindColumns.Add(fieldskeysMap);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "U_TEmplName");
            fieldskeysMap.Add("ColumnDescription", "To Employee Name");
            listFindColumns.Add(fieldskeysMap);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "DocEntry");
            fieldskeysMap.Add("ColumnDescription", "DocEntry");
            listFindColumns.Add(fieldskeysMap);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "Remark");
            fieldskeysMap.Add("ColumnDescription", "Remark");
            listFindColumns.Add(fieldskeysMap);

            formProperties.Add("FindColumns", listFindColumns);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("FormColumnAlias", "DocEntry");
            fieldskeysMap.Add("FormColumnDescription", "DocEntry");
            listFormColumns.Add(fieldskeysMap);

            formProperties.Add("FormColumns", listFormColumns);

            //ცხრილური ნაწილები
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("TableName", "BDOSFASTR1");
            fieldskeysMap.Add("ObjectName", "BDOSFASTR1");
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
                oCreationPackage.UniqueID = "UDO_F_BDOSFASTRD_D";
                oCreationPackage.String = BDOSResources.getTranslate("FixedAssetTransferDocument");
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

                if (pVal.BeforeAction && pVal.MenuUID == "1284")
                {
                    ////oDocForm.Items.Item("RemarksE").Click();

                    //// პროექტზე კორექტირება არ უნდა არსებობდეს
                    //string prjCode = oDocForm.DataSources.DBDataSources.Item("@BDOSLPSCDC").GetValue("U_PrjCode", 0).Trim();
                    //string docType = oDocForm.DataSources.DBDataSources.Item("@BDOSLPSCDC").GetValue("U_DocType", 0).Trim();
                    //if (prjCode != "" && docType == "Basic")
                    //{
                    //    string query = @"SELECT ""DocEntry"" FROM ""@BDOSLPSCDC"" WHERE ""U_PrjCode"" = N'" + prjCode + @"' AND ""U_DocType"" = 'Correction' AND ""Canceled"" = 'N'";

                    //    SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    //    oRecordSet.DoQuery(query);
                    //    if (!oRecordSet.EoF)
                    //    {
                    //        Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("YouCanNotCancelBasicCorrectedDocument"), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    //        BubbleEvent = false;
                    //    }
                    //}
                }
                else if (pVal.BeforeAction == false && pVal.MenuUID == "BDOSAddRow")
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

            if (oForm.TypeEx == "UDO_FT_UDO_F_BDOSFASTRD_D")
            {
                if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD)
                {
                    if (BusinessObjectInfo.BeforeAction == true)
                    {
                        checkDoc(oForm, out errorText);
                        if (errorText != null)
                        {
                            Program.uiApp.StatusBar.SetSystemMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            Program.uiApp.MessageBox(BDOSResources.getTranslate("OperationUnsuccesfullSeeLog"));
                            BubbleEvent = false;
                        }
                        else
                        {
                            updateAsset(oForm, false, out errorText);
                            if (errorText != null)
                            {
                                Program.uiApp.StatusBar.SetSystemMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                Program.uiApp.MessageBox(BDOSResources.getTranslate("OperationUnsuccesfullSeeLog"));
                                BubbleEvent = false;
                            }
                        }
                    }
                }

                if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE & BusinessObjectInfo.BeforeAction == true) //false & BusinessObjectInfo.ActionSuccess ==
                {
                    if (Program.cancellationTrans == true & Program.canceledDocEntry != 0)
                    {
                        updateAsset(oForm, true, out errorText);
                        if (errorText != null)
                        {
                            Program.uiApp.StatusBar.SetSystemMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            Program.uiApp.MessageBox(BDOSResources.getTranslate("OperationUnsuccesfullSeeLog"));
                            BubbleEvent = false;
                        }
                        cancellation(Program.canceledDocEntry, out errorText);
                        Program.canceledDocEntry = 0;
                    }
                }

                if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD & BusinessObjectInfo.BeforeAction == false)
                {
                    formDataLoad(oForm, out errorText);
                }

            }
        }

        public static void uiApp_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (pVal.ItemUID.Length > 1)
                {
                    if (pVal.ItemUID.Substring(0, pVal.ItemUID.Length - 1) == "Folder" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED & pVal.BeforeAction == false)
                    {
                        oForm.PaneLevel = Convert.ToInt32(pVal.ItemUID.Substring(6, 1));
                    }
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD & pVal.BeforeAction == true)
                {
                    createFormItems(oForm, out BubbleEvent, out errorText);
                    Program.FORM_LOAD_FOR_VISIBLE = true;
                    Program.FORM_LOAD_FOR_ACTIVATE = true;
                    //setVisibleFormItems(oForm, out errorText);
                }

                //if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE && Program.FORM_LOAD_FOR_VISIBLE) //&& !pVal.BeforeAction
                //{
                //    setVisibleFormItems(oForm, out errorText);
                //}

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_GOT_FOCUS && oForm.Visible == true && oForm.VisibleEx == true && openFormEvent == false)
                {
                    setVisibleFormItems(oForm, out errorText);

                    string docEntry = oForm.DataSources.DBDataSources.Item("@BDOSFASTRD").GetValue("DocEntry", 0).Trim();
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
                        oForm.Title = BDOSResources.getTranslate("FixedAssetTransferDocument");
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

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                {
                    SAPbouiCOM.ChooseFromListEvent oCFLEvento = null;
                    oCFLEvento = ((SAPbouiCOM.ChooseFromListEvent)(pVal));
                    chooseFromList(oForm, oCFLEvento, pVal, out errorText);
                    //---------------------------------------------------------------------------

                    if (pVal.BeforeAction == true)
                    {
                        oCFLEvento = ((SAPbouiCOM.ChooseFromListEvent)(pVal));
                        string sCFL_ID = oCFLEvento.ChooseFromListUID;
                        if (sCFL_ID == "ItemMTR_CFL" && oForm.Items.Item("DocDateE").Specific.Value == "")
                        {
                            BubbleEvent = false;
                            Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("TheFollowingFieldIsMandatory")+ ": "+ BDOSResources.getTranslate("PostingDate"), SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                        }
                        SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(sCFL_ID);
                        SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();

                        if (oCFLEvento.ChooseFromListUID == "Waybill_CFL")
                        {
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                            {
                                goto SkipToEnd;
                            }
                            string query = @"Select ""DocEntry"" from ""@BDO_WBLD"" where ""U_baseDoc"" =0";

                            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                            oRecordSet.DoQuery(query);

                            SAPbouiCOM.Condition oCon = null;
                            while (!oRecordSet.EoF)
                            {
                                oCon = oCons.Add();
                                oCon.Alias = "DocEntry";
                                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                                oCon.CondVal = oRecordSet.Fields.Item("DocEntry").Value.ToString();

                                oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR;


                                oRecordSet.MoveNext();
                            }
                            oCon = oCons.Add();
                            oCon.Alias = "DocEntry";
                            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                            //oCon.CondVal = "";


                            oCFL.SetConditions(oCons);

                        SkipToEnd:;
                        }
                    }
                    else
                    {
                        oCFLEvento = null;
                        oCFLEvento = ((SAPbouiCOM.ChooseFromListEvent)(pVal));

                        if (oCFLEvento.ChooseFromListUID == "Waybill_CFL")
                        {
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                            {
                                goto SkipToEnd;
                            }

                            SAPbouiCOM.DataTable oDataTableSelectedObjects = oCFLEvento.SelectedObjects;

                            if (oDataTableSelectedObjects == null)
                            {
                                return;
                            }

                            int answer = Program.uiApp.MessageBox(BDOSResources.getTranslate("LinkWaybillToDocument"), 1, BDOSResources.getTranslate("Yes"), BDOSResources.getTranslate("No"), "");
                            if (answer == 2)
                            {
                                return;
                            }


                            oForm.Freeze(true);

                            int newDocEntry = oDataTableSelectedObjects.GetValue("DocEntry", 0);
                            int docEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("@BDOSFASTRD").GetValue("DocEntry", 0));

                            //არჩეულ ზედნადებში უნდა ჩავწეროთ ამ დოკუმენტის ნომრები
                            SAPbobsCOM.CompanyService oCompanyService = null;
                            SAPbobsCOM.GeneralService oGeneralService = null;
                            SAPbobsCOM.GeneralData oGeneralData = null;
                            SAPbobsCOM.GeneralDataParams oGeneralParams = null;
                            oCompanyService = Program.oCompany.GetCompanyService();
                            oGeneralService = oCompanyService.GetGeneralService("UDO_F_BDO_WBLD_D");
                            oGeneralData = ((SAPbobsCOM.GeneralData)(oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)));
                            //Get UDO record
                            oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                            oGeneralParams.SetProperty("DocEntry", newDocEntry);
                            oGeneralData = oGeneralService.GetByParams(oGeneralParams);

                            oGeneralData.SetProperty("U_baseDoc", docEntry);
                            oGeneralData.SetProperty("U_baseDTxt", docEntry.ToString());
                            oGeneralService.Update(oGeneralData);

                            oForm.DataSources.UserDataSources.Item("BDO_WblDoc").ValueEx = newDocEntry.ToString();
                            StockTransfer.formDataLoad(oForm, out errorText);

                            BubbleEvent = true;
                            oForm.Freeze(false);
                            oForm.Update();

                        SkipToEnd:;

                        }
                    }



                    //---------------------------------------------------------------------------
                }


                if (pVal.ItemUID == "BDO_WblTxt" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED & pVal.BeforeAction == false)
                {
                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                    {
                        goto SkipToEnd;
                    }

                    oForm.Freeze(true);

                    int docEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("@BDOSFASTRD").GetValue("DocEntry", 0));
                    string cancelled = oForm.DataSources.DBDataSources.Item("@BDOSFASTRD").GetValue("CANCELED", 0).Trim();
                    string BDO_WblDoc = oForm.DataSources.UserDataSources.Item("BDO_WblDoc").ValueEx;
                    int newDocEntry = 0;

                    if (docEntry != 0 & (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_VIEW_MODE))
                    {
                        if (BDO_WblDoc == "" && cancelled == "N")
                        {

                            string objectType = "UDO_F_BDOSFASTRD_D";
                            BDO_Waybills.createDocument(objectType, docEntry, null, null, null, null, out newDocEntry, out errorText);
                            oForm.DataSources.UserDataSources.Item("BDO_WblDoc").ValueEx = newDocEntry.ToString();
                            if (errorText == null & newDocEntry != 0)
                            {
                                Program.uiApp.MessageBox(BDOSResources.getTranslate("WaybillCreatedSuccesfully") + " DocEntry : " + newDocEntry);
                                StockTransfer.formDataLoad(oForm, out errorText);
                            }
                            else
                            {
                                Program.uiApp.MessageBox(errorText);
                            }
                        }
                        else if (cancelled != "N")
                        {
                            errorText = BDOSResources.getTranslate("DocumentMustNotBeCancelledOrCancellation");
                        }
                        else if (BDO_WblDoc != "")
                        {
                            errorText = BDOSResources.getTranslate("DocumentLinkedToWaybill");
                        }
                        BubbleEvent = true;
                    }
                    else
                    {
                        errorText = BDOSResources.getTranslate("ToCreateWaybillWriteDocument");
                    }

                    oForm.Freeze(false);
                    oForm.Update();

                    if (newDocEntry != 0)
                    {
                        Program.uiApp.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_UserDefinedObject, "UDO_F_BDO_WBLD_D", newDocEntry.ToString());
                    }

                SkipToEnd:;
                    
                }

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("BDO_WblDoc").Enabled = false;
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_GOT_FOCUS && pVal.BeforeAction == false)
                {
                    

                    if (oForm.Items.Item("0_U_E").ToPane != 2)
                    {
                        oForm.Visible = false;
                        oForm.VisibleEx = false;
                        oForm.Freeze(true);
                        oForm.Items.Item("0_U_E").FromPane = 1;
                        oForm.Items.Item("0_U_E").ToPane = 2;
                        createFolder(oForm, out errorText);

                        oForm.Items.Item("Folder2").Click();
                        oForm.PaneLevel = 2;

                        if (oForm.Items.Item("0_U_E").Specific.Value != "")
                        {
                            oForm.DataSources.DBDataSources.Item(0).Clear();
                        }

                        oForm.Visible = true;
                        oForm.VisibleEx = true;
                        oForm.Update();
                        oForm.Refresh();
                        oForm.Freeze(false);
                    }
                }

                if (pVal.ItemUID == "AssetMTR" && pVal.ColUID == "LineID" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK & pVal.BeforeAction == true)
                {
                    BubbleEvent = false;
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "DistrRul1S" || pVal.ItemUID == "DistrRul2S" || pVal.ItemUID == "DistrRul3S"
                         || pVal.ItemUID == "DistrRul4S" || pVal.ItemUID == "DistrRul5S")
                    {
                        setVisibleFormItems(oForm, out errorText);
                    }

                }

                

                    if (pVal.ItemUID != "" && pVal.ItemUID == "FillMTR" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.BeforeAction == false)
                {
                    if (oForm.Items.Item(pVal.ItemUID).Enabled == false)
                    {
                        return;
                    }

                    fillAssets(oForm, out errorText);
                }
				//if (pVal.ItemUID == "CardCodeE" && pVal.BeforeAction == false)
				//{
				//	SAPbouiCOM.EditText oEdit = oForm.Items.Item("CardCodeE").Specific;
				//	string cardCode = oEdit.Value;
				//	if (string.IsNullOrEmpty(cardCode))
				//	{
				//		oForm.Items.Item("CardNameE").Specific.Value = "";
				//	}
				//}
				//if (pVal.ItemUID == "FLocCodeE" && pVal.BeforeAction == false)
				//{
				//	SAPbouiCOM.EditText oEdit = oForm.Items.Item("FLocCodeE").Specific;
				//	string fLocCode = oEdit.Value;
				//	if (string.IsNullOrEmpty(fLocCode))
				//	{
				//		oForm.Items.Item("FLocNameE").Specific.Value = "";
				//	}
				//}
				//if (pVal.ItemUID == "FEmplIDE" && pVal.BeforeAction == false)
				//{
				//	SAPbouiCOM.EditText oEdit = oForm.Items.Item("FEmplIDE").Specific;
				//	string fEmplID = oEdit.Value;
				//	if (string.IsNullOrEmpty(fEmplID))
				//	{
				//		oForm.Items.Item("FEmplNameE").Specific.Value = "";
				//	}
				//}
				//if (pVal.ItemUID == "TLocCodeE" && pVal.BeforeAction == false)
				//{
				//	SAPbouiCOM.EditText oEdit = oForm.Items.Item("TLocCodeE").Specific;
				//	string tLocCode = oEdit.Value;
				//	if (string.IsNullOrEmpty(tLocCode))
				//	{
				//		oForm.Items.Item("TLocNameE").Specific.Value = "";
				//	}
				//}
				//if (pVal.ItemUID == "TEmplIDE" && pVal.BeforeAction == false)
				//{
				//	SAPbouiCOM.EditText oEdit = oForm.Items.Item("TEmplIDE").Specific;
				//	string tEmplID = oEdit.Value;
				//	if (string.IsNullOrEmpty(tEmplID))
				//	{
				//		oForm.Items.Item("TEmplNameE").Specific.Value = "";
				//	}
				//}

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

            if (eventInfo.ItemUID == "AssetMTR")
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

        public static void createFolder(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            try
            {
                SAPbouiCOM.UserDataSource FolderDS = oForm.DataSources.UserDataSources.Item("FolderDS");
            }
            catch
            {
                SAPbouiCOM.UserDataSource FolderDS = oForm.DataSources.UserDataSources.Add("FolderDS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
            }

            try
            {
                for (int i = 1; i <= 2; i++)
                {
                    string folderName = "";
                    if (i == 1)
                    {
                        folderName = BDOSResources.getTranslate("CostAccounting");
                    }
                    else if (i == 2)
                    {
                        folderName = BDOSResources.getTranslate("Contents");
                    }

                    Dictionary<string, object> formItems = new Dictionary<string, object>();
                    string itemName = "Folder" + i.ToString();
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_FOLDER);
                    formItems.Add("Bound", true);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", "FolderDS");
                    formItems.Add("Width", oForm.Width / 3);
                    formItems.Add("Top", oForm.Items.Item("BDO_WblTxt").Top + oForm.Items.Item("BDO_WblTxt").Height + 1);
                    formItems.Add("Height", oForm.Items.Item("AssetMTR").Height + 30);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", folderName);
                    formItems.Add("Pane", i);
                    formItems.Add("ValOn", "0");
                    formItems.Add("ValOff", itemName);

                    if (i != 1)
                    {
                        formItems.Add("GroupWith", "Folder" + (i - 1).ToString());
                    }

                    formItems.Add("AffectsFormMode", false);
                    formItems.Add("Description", folderName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        continue;
                    }
                }
            }
            catch
            {
                string errMsg;
                int errCode;
                Program.oCompany.GetLastError(out errCode, out errMsg);
                Program.uiApp.StatusBar.SetSystemMessage(errMsg);
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

            //მარცხენა რიგი
            top = top + height + 1;
            oForm.AutoManaged = true;

            formItems = new Dictionary<string, object>();
            itemName = "CardCodeS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("CustomerCode"));
            formItems.Add("LinkTo", "CardCodeE");
            formItems.Add("FromPane", 1);
            formItems.Add("ToPane", 2);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            string objectTypeEmployee = "171";
            string objectTypeLocation = "144";
            string objectTypeItem = "4";
            string objectTypeBP = "2";

            bool multiSelection = false;
            string uniqueID_lf_BusinessPartnerCFL = "BusinessPartner_CFL";
            FormsB1.addChooseFromList(oForm, multiSelection, objectTypeBP, uniqueID_lf_BusinessPartnerCFL);

            //პირობის დადება ბიზნესპარტნიორის არჩევის სიაზე
            SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(uniqueID_lf_BusinessPartnerCFL);
            SAPbouiCOM.Conditions oCons = oCFL.GetConditions();
            SAPbouiCOM.Condition oCon = oCons.Add();
            oCon.Alias = "CardType";
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCon.CondVal = "C"; //მყიდველი
            oCFL.SetConditions(oCons);

            formItems = new Dictionary<string, object>();
            itemName = "CardCodeE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSFASTRD");
            formItems.Add("Alias", "U_CardCode");
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
            formItems.Add("SetAutoManaged", true);
            formItems.Add("FromPane", 1);
            formItems.Add("ToPane", 2);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "CardNameE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSFASTRD");
            formItems.Add("Alias", "U_CardName");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e + width_e / 4);
            formItems.Add("Width", width_e * 3 / 4);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", false);
            formItems.Add("SetAutoManaged", true);
            formItems.Add("FromPane", 1);
            formItems.Add("ToPane", 2);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "CardLB"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            formItems.Add("Left", left_e - 20);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("LinkTo", "CardCodeE");
            formItems.Add("LinkedObjectType", objectTypeBP);
            formItems.Add("FromPane", 1);
            formItems.Add("ToPane", 2);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "FLocCodeS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("FromLocation"));
            formItems.Add("LinkTo", "FLocCodeE");
            formItems.Add("FromPane", 1);
            formItems.Add("ToPane", 2);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            string uniqueID_lf_FromLoc = "FromLoc_CFL";
            FormsB1.addChooseFromList(oForm, multiSelection, objectTypeLocation, uniqueID_lf_FromLoc);

            formItems = new Dictionary<string, object>();
            itemName = "FLocCodeE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSFASTRD");
            formItems.Add("Alias", "U_FLocCode");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e / 4);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("ChooseFromListUID", uniqueID_lf_FromLoc);
            formItems.Add("ChooseFromListAlias", "Code");
            formItems.Add("SetAutoManaged", true);
            formItems.Add("FromPane", 1);
            formItems.Add("ToPane", 2);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "FLocNameE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSFASTRD");
            formItems.Add("Alias", "U_FLocName");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e + width_e / 4);
            formItems.Add("Width", width_e * 3 / 4);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", false);
            formItems.Add("SetAutoManaged", true);
            formItems.Add("FromPane", 1);
            formItems.Add("ToPane", 2);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "FLocLB"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            formItems.Add("Left", left_e - 20);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("LinkTo", "FLocCodeE");
            formItems.Add("LinkedObjectType", objectTypeLocation);
            formItems.Add("FromPane", 1);
            formItems.Add("ToPane", 2);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "FEmplIDS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("FromEmployee"));
            formItems.Add("LinkTo", "FEmplIDE");
            formItems.Add("FromPane", 1);
            formItems.Add("ToPane", 2);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            string uniqueID_lf_FromEmpl = "FromEmpl_CFL";
            FormsB1.addChooseFromList(oForm, multiSelection, objectTypeEmployee, uniqueID_lf_FromEmpl);

            formItems = new Dictionary<string, object>();
            itemName = "FEmplIDE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSFASTRD");
            formItems.Add("Alias", "U_FEmplID");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e / 4);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("ChooseFromListUID", uniqueID_lf_FromEmpl);
            formItems.Add("ChooseFromListAlias", "empID");
            formItems.Add("SetAutoManaged", true);
            formItems.Add("FromPane", 1);
            formItems.Add("ToPane", 2);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "FEmplNameE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSFASTRD");
            formItems.Add("Alias", "U_FEmplName");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e + width_e / 4);
            formItems.Add("Width", width_e * 3 / 4);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", false);
            formItems.Add("SetAutoManaged", true);
            formItems.Add("FromPane", 1);
            formItems.Add("ToPane", 2);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "FEmplLB"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            formItems.Add("Left", left_e - 20);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("LinkTo", "FEmplIDE");
            formItems.Add("LinkedObjectType", objectTypeEmployee);
            formItems.Add("FromPane", 1);
            formItems.Add("ToPane", 2);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }
            //----------------------------------------------------------------
            
            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            string caption = BDOSResources.getTranslate("CreateWaybill");
            itemName = "BDO_WblTxt"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", caption);
            formItems.Add("TextStyle", 4);
            formItems.Add("FontSize", 10);
            formItems.Add("Enabled", true);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            multiSelection = false;
            string objectType = "UDO_F_BDO_WBLD_D"; //Waybill document
            string uniqueID_WaybillCFL = "Waybill_CFL";
            FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_WaybillCFL);

            formItems = new Dictionary<string, object>();
            itemName = "BDO_WblDoc"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "UserDataSources");
            formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
            formItems.Add("Length", 11);
            formItems.Add("TableName", "");
            formItems.Add("Alias", itemName);
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e / 2);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("AffectsFormMode", false);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", false);
            formItems.Add("ChooseFromListUID", uniqueID_WaybillCFL);
            formItems.Add("ChooseFromListAlias", "DocEntry");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "BDO_WblLB"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            formItems.Add("Left", left_e - 20);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("LinkTo", "BDO_WblDoc");
            formItems.Add("LinkedObjectType", objectType);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "BDO_WblID"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", "");
            //formItems.Add("TextStyle", 4);
            formItems.Add("FontSize", 10);
            formItems.Add("Enabled", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "BDO_WblNum"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", "");
            //formItems.Add("TextStyle", 4);
            formItems.Add("FontSize", 10);
            formItems.Add("Enabled", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "BDO_WblSts"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_e * 2);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", "");
            //formItems.Add("TextStyle", 4);
            formItems.Add("FontSize", 10);
            formItems.Add("Enabled", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }



            //----------------------------------------------------------------
            top = top + height + 1;
            int topLeft = top;

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
            formItems.Add("TableName", "@BDOSFASTRD");
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
            itemName = "DocDateS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("PostingDate"));
            formItems.Add("LinkTo", "DocDateE");
            formItems.Add("FromPane", 1);
            formItems.Add("ToPane", 2);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "DocDateE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSFASTRD");
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
            formItems.Add("FromPane", 1);
            formItems.Add("ToPane", 2);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "TLocCodeS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("ToLocation"));
            formItems.Add("LinkTo", "TLocCodeE");
            formItems.Add("FromPane", 1);
            formItems.Add("ToPane", 2);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            string uniqueID_lf_ToLoc = "ToLoc_CFL";
            FormsB1.addChooseFromList(oForm, multiSelection, objectTypeLocation, uniqueID_lf_ToLoc);

            formItems = new Dictionary<string, object>();
            itemName = "TLocCodeE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSFASTRD");
            formItems.Add("Alias", "U_TLocCode");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e / 4);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("ChooseFromListUID", uniqueID_lf_ToLoc);
            formItems.Add("ChooseFromListAlias", "Code");
            formItems.Add("SetAutoManaged", true);
            formItems.Add("FromPane", 1);
            formItems.Add("ToPane", 2);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "TLocNameE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSFASTRD");
            formItems.Add("Alias", "U_TLocName");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e + width_e / 4);
            formItems.Add("Width", width_e * 3 / 4);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", false);
            formItems.Add("SetAutoManaged", true);
            formItems.Add("FromPane", 1);
            formItems.Add("ToPane", 2);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "TLocLB"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            formItems.Add("Left", left_e - 20);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("LinkTo", "TLocCodeE");
            formItems.Add("LinkedObjectType", objectTypeLocation);
            formItems.Add("FromPane", 1);
            formItems.Add("ToPane", 2);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "TEmplIDS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("ToEmployee"));
            formItems.Add("LinkTo", "TEmplIDE");
            formItems.Add("FromPane", 1);
            formItems.Add("ToPane", 2);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            string uniqueID_lf_ToEmpl = "ToEmpl_CFL";
            FormsB1.addChooseFromList(oForm, multiSelection, objectTypeEmployee, uniqueID_lf_ToEmpl);

            formItems = new Dictionary<string, object>();
            itemName = "TEmplIDE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSFASTRD");
            formItems.Add("Alias", "U_TEmplID");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e / 4);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("ChooseFromListUID", uniqueID_lf_ToEmpl);
            formItems.Add("ChooseFromListAlias", "empID");
            formItems.Add("SetAutoManaged", true);
            formItems.Add("FromPane", 1);
            formItems.Add("ToPane", 2);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "TEmplNameE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSFASTRD");
            formItems.Add("Alias", "U_TEmplName");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e + width_e / 4);
            formItems.Add("Width", width_e * 3 / 4);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", false);
            formItems.Add("SetAutoManaged", true);
            formItems.Add("FromPane", 1);
            formItems.Add("ToPane", 2);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "TEmplLB"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            formItems.Add("Left", left_e - 20);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("LinkTo", "TEmplIDE");
            formItems.Add("LinkedObjectType", objectTypeEmployee);
            formItems.Add("FromPane", 1);
            formItems.Add("ToPane", 2);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            //Add Delete
            left_s = 6;
            left_e = 127;
            top = Math.Max(topLeft, top) + 2 * height + 1;

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
            formItems.Add("FromPane", 2);
            formItems.Add("ToPane", 2);

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
            itemName = "AssetMTR"; //10 characters
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
            formItems.Add("FromPane", 2);
            formItems.Add("ToPane", 2);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            string uniqueID_lf_ItemMTR_CFL = "ItemMTR_CFL";
            FormsB1.addChooseFromList(oForm, true, objectTypeItem, uniqueID_lf_ItemMTR_CFL);
            //პირობის დადება ძს არჩევის სიაზე
            SAPbouiCOM.ChooseFromList oCFL_Item = oForm.ChooseFromLists.Item(uniqueID_lf_ItemMTR_CFL);
            SAPbouiCOM.Conditions oCons_Item = oCFL_Item.GetConditions();
            SAPbouiCOM.Condition oCon_Item = oCons_Item.Add();
            oCon_Item.Alias = "ItemType";
            oCon_Item.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCon_Item.CondVal = "F"; //Fixed Assets
            oCFL_Item.SetConditions(oCons_Item);

            string uniqueID_lf_ToLocMTR_CFL = "ToLocMTR_CFL";
            FormsB1.addChooseFromList(oForm, multiSelection, objectTypeLocation, uniqueID_lf_ToLocMTR_CFL);

            string uniqueID_lf_ToEmplMTR_CFL = "ToEmplMTR_CFL";
            FormsB1.addChooseFromList(oForm, multiSelection, objectTypeEmployee, uniqueID_lf_ToEmplMTR_CFL);

            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("AssetMTR").Specific));
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

            oColumn = oColumns.Add("ItemName", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("ItemName");
            oColumn.Width = 60;
            oColumn.Editable = false;
            oColumn.Visible = true;

            oColumn = oColumns.Add("SerNo", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("SerNo");
            oColumn.Width = 60;
            oColumn.Editable = false;
            oColumn.Visible = true;

            oColumn = oColumns.Add("Quantity", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Quantity");
            oColumn.Width = 60;
            oColumn.Editable = false;
            oColumn.Visible = true;

            oColumn = oColumns.Add("APC", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("AcquisitionAndProductionCost");
            oColumn.Width = 60;
            oColumn.Editable = false;
            oColumn.Visible = true;

            oColumn = oColumns.Add("FLocCode", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("FromLocation");
            oColumn.Width = 60;
            oColumn.Editable = false;
            oColumn.Visible = true;
            oColumn.ExtendedObject.LinkedObjectType = objectTypeLocation;

            oColumn = oColumns.Add("FLocName", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("FromLocationName");
            oColumn.Width = 60;
            oColumn.Editable = false;
            oColumn.Visible = true;

            oColumn = oColumns.Add("FEmplID", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("FromEmployee");
            oColumn.Width = 60;
            oColumn.Editable = false;
            oColumn.Visible = true;
            oColumn.ExtendedObject.LinkedObjectType = objectTypeEmployee;

            oColumn = oColumns.Add("FEmplName", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("FromEmployeeName");
            oColumn.Width = 60;
            oColumn.Editable = false;
            oColumn.Visible = true;

            oColumn = oColumns.Add("TLocCode", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("ToLocation");
            oColumn.Width = 60;
            oColumn.Editable = true;
            oColumn.Visible = true;
            oColumn.ExtendedObject.LinkedObjectType = objectTypeLocation;

            oColumn = oColumns.Add("TLocName", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("ToLocationName");
            oColumn.Width = 60;
            oColumn.Editable = false;
            oColumn.Visible = true;

            oColumn = oColumns.Add("TEmplID", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("ToEmployee");
            oColumn.Width = 60;
            oColumn.Editable = true;
            oColumn.Visible = true;
            oColumn.ExtendedObject.LinkedObjectType = objectTypeEmployee;

            oColumn = oColumns.Add("TEmplName", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("ToEmployeeName");
            oColumn.Width = 60;
            oColumn.Editable = false;
            oColumn.Visible = true;

            oColumn = oColumns.Add("Remark", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Remark");
            oColumn.Width = 60;
            oColumn.Editable = true;
            oColumn.Visible = true;

            SAPbouiCOM.DBDataSource oDBDataSource;
            oDBDataSource = oForm.DataSources.DBDataSources.Add("@BDOSFASTR1");

            oColumn = oColumns.Item("LineID");
            oColumn.DataBind.SetBound(true, "@BDOSFASTR1", "LineID");

            oColumn = oColumns.Item("ItemCode");
            oColumn.DataBind.SetBound(true, "@BDOSFASTR1", "U_ItemCode");
            oColumn.ChooseFromListUID = uniqueID_lf_ItemMTR_CFL;
            oColumn.ChooseFromListAlias = "ItemCode";

            oColumn = oColumns.Item("ItemName");
            oColumn.DataBind.SetBound(true, "@BDOSFASTR1", "U_ItemName");

            oColumn = oColumns.Item("SerNo");
            oColumn.DataBind.SetBound(true, "@BDOSFASTR1", "U_SerNo");

            oColumn = oColumns.Item("Quantity");
            oColumn.DataBind.SetBound(true, "@BDOSFASTR1", "U_Quantity");

            oColumn = oColumns.Item("APC");
            oColumn.DataBind.SetBound(true, "@BDOSFASTR1", "U_APC");

            oColumn = oColumns.Item("FLocCode");
            oColumn.DataBind.SetBound(true, "@BDOSFASTR1", "U_FLocCode");

            oColumn = oColumns.Item("FLocName");
            oColumn.DataBind.SetBound(true, "@BDOSFASTR1", "U_FLocName");

            oColumn = oColumns.Item("TLocCode");
            oColumn.DataBind.SetBound(true, "@BDOSFASTR1", "U_TLocCode");
            oColumn.ChooseFromListUID = uniqueID_lf_ToLocMTR_CFL;
            oColumn.ChooseFromListAlias = "Code";

            oColumn = oColumns.Item("TLocName");
            oColumn.DataBind.SetBound(true, "@BDOSFASTR1", "U_TLocName");

            oColumn = oColumns.Item("FEmplID");
            oColumn.DataBind.SetBound(true, "@BDOSFASTR1", "U_FEmplID");

            oColumn = oColumns.Item("FEmplName");
            oColumn.DataBind.SetBound(true, "@BDOSFASTR1", "U_FEmplName");

            oColumn = oColumns.Item("TEmplID");
            oColumn.DataBind.SetBound(true, "@BDOSFASTR1", "U_TEmplID");
            oColumn.ChooseFromListUID = uniqueID_lf_ToEmplMTR_CFL;
            oColumn.ChooseFromListAlias = "empID";

            oColumn = oColumns.Item("TEmplName");
            oColumn.DataBind.SetBound(true, "@BDOSFASTR1", "U_TEmplName");

            oColumn = oColumns.Item("Remark");
            oColumn.DataBind.SetBound(true, "@BDOSFASTR1", "U_Remark");

            //მეორე ჩანართი
            int pane = 1;

            formItems = new Dictionary<string, object>();
            itemName = "ProjectS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Project"));
            formItems.Add("LinkTo", "ProjectE");
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            objectType = "63";
            string uniqueID_lf_Project = "Project_CFL";
            FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_lf_Project);

            formItems = new Dictionary<string, object>();
            itemName = "ProjectE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSFASTRD");
            formItems.Add("Alias", "U_PrjCode");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("ChooseFromListUID", uniqueID_lf_Project);
            formItems.Add("ChooseFromListAlias", "PrjCode");
            formItems.Add("SetAutoManaged", true);
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "ProjectLB"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            formItems.Add("Left", left_e - 20);
            formItems.Add("Top", top);
            formItems.Add("Height", 14);
            formItems.Add("UID", itemName);
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);
            formItems.Add("LinkTo", "ProjectE");
            formItems.Add("LinkedObjectType", objectType);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }


            Dictionary<string, string> activeDimensionsList = CommonFunctions.getActiveDimensionsList(out errorText);

            for (int i = 1; i <= activeDimensionsList.Count; i++)
            {
                top = top + height + 1;
                //ChngDstRl2
                formItems = new Dictionary<string, object>();
                itemName = "DistrRul" + i + "S"; //10 characters
                formItems.Add("isDataSource", true);
                formItems.Add("DataSource", "DBDataSources");
                formItems.Add("TableName", "@BDOSFASTRD");
                formItems.Add("Alias", "U_ChngDstRl" + i);
                formItems.Add("Bound", true);
                formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
                formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                formItems.Add("Length", 1);
                formItems.Add("Left", left_s);
                formItems.Add("Width", width_s);
                formItems.Add("Top", top);
                formItems.Add("Height", height);
                formItems.Add("UID", itemName);
                formItems.Add("Caption", activeDimensionsList[i.ToString()]);
                formItems.Add("ValOff", "N");
                formItems.Add("ValOn", "Y");
                formItems.Add("DisplayDesc", true);
                formItems.Add("SetAutoManaged", true);
                formItems.Add("FromPane", pane);
                formItems.Add("ToPane", pane);

                FormsB1.createFormItem(oForm, formItems, out errorText);
                if (errorText != null)
                {
                    return;
                }

                objectType = "62";
                string uniqueID_lf_DistrRule = "Rule_CFL" + i.ToString() + "A";
                FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_lf_DistrRule);

                formItems = new Dictionary<string, object>();
                itemName = "DistrRul" + i + "E"; //10 characters
                formItems.Add("isDataSource", true);
                formItems.Add("DataSource", "DBDataSources");
                formItems.Add("TableName", "@BDOSFASTRD");
                formItems.Add("Alias", "U_DistrRule" + i);
                formItems.Add("Bound", true);
                formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                formItems.Add("Left", left_e);
                formItems.Add("Width", width_e);
                formItems.Add("Top", top);
                formItems.Add("Height", height);
                formItems.Add("UID", itemName);
                formItems.Add("DisplayDesc", true);
                formItems.Add("ChooseFromListUID", uniqueID_lf_DistrRule);
                formItems.Add("ChooseFromListAlias", "OcrCode");
                formItems.Add("SetAutoManaged", true);
                formItems.Add("FromPane", pane);
                formItems.Add("ToPane", pane);

                FormsB1.createFormItem(oForm, formItems, out errorText);
                if (errorText != null)
                {
                    return;
                }

                formItems = new Dictionary<string, object>();
                itemName = "DstrRul" + i + "LB"; //10 characters
                formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                formItems.Add("Left", left_e - 20);
                formItems.Add("Top", top);
                formItems.Add("Height", 14);
                formItems.Add("UID", itemName);
                formItems.Add("FromPane", pane);
                formItems.Add("ToPane", pane);
                formItems.Add("LinkTo", "DistrRul" + i + "E");
                formItems.Add("LinkedObjectType", objectType);

                FormsB1.createFormItem(oForm, formItems, out errorText);
                if (errorText != null)
                {
                    return;
                }
            }
            //მეორე ჩანართი

            //სარდაფი
            left_s = 6;
            left_e = 127;
            top = oForm.Items.Item("AssetMTR").Top + oForm.Items.Item("AssetMTR").Height + 40;

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
            formItems.Add("FromPane", 1);
            formItems.Add("ToPane", 2);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "CreatorE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSFASTRD");
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
            formItems.Add("FromPane", 1);
            formItems.Add("ToPane", 2);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            //შენიშვნა
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
            formItems.Add("FromPane", 1);
            formItems.Add("ToPane", 2);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "RemarksE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSFASTRD");
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
            formItems.Add("FromPane", 1);
            formItems.Add("ToPane", 2);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            GC.Collect();
        }

        public static void formDataLoad(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            try
            {
                openFormEvent = false;

                setVisibleFormItems(oForm, out errorText);
                //-------------------------------------------

                int docEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("@BDOSFASTRD").GetValue("DocEntry", 0));
                string caption = BDOSResources.getTranslate("CreateWaybill");
                int wblDocEntry = 0;
                string wblID = "";
                string wblNum = "";
                string wblSts = "";

                if (docEntry != 0)
                {
                    Dictionary<string, string> wblDocInfo = BDO_Waybills.getWaybillDocumentInfo(docEntry, "UDO_F_BDOSFASTRD_D", out errorText);
                    wblDocEntry = Convert.ToInt32(wblDocInfo["DocEntry"]);
                    wblID = wblDocInfo["wblID"];
                    wblNum = wblDocInfo["number"];
                    wblSts = wblDocInfo["status"];

                    if (wblDocEntry != 0)
                    {
                        caption = BDOSResources.getTranslate("WaybillDocEntry");
                    }
                }
                else
                {
                    caption = BDOSResources.getTranslate("CreateWaybill");
                    wblDocEntry = 0;
                }

                oForm.DataSources.UserDataSources.Item("BDO_WblDoc").ValueEx = wblDocEntry == 0 ? "" : wblDocEntry.ToString();
                SAPbouiCOM.StaticText oStaticText = (SAPbouiCOM.StaticText)oForm.Items.Item("BDO_WblTxt").Specific;
                oStaticText.Caption = caption; oStaticText = (SAPbouiCOM.StaticText)oForm.Items.Item("BDO_WblID").Specific;
                oStaticText.Caption = wblID != "" ? "ID : " + wblID : "";
                oStaticText = (SAPbouiCOM.StaticText)oForm.Items.Item("BDO_WblNum").Specific;
                oStaticText.Caption = wblNum != "" ? "№ " + wblNum : "";
                oStaticText = (SAPbouiCOM.StaticText)oForm.Items.Item("BDO_WblSts").Specific;
                oStaticText.Caption = wblSts != "" ? BDOSResources.getTranslate("Status") + " : " + wblSts : "";


                oForm.Items.Item("BDO_WblDoc").Enabled = (oForm.DataSources.DBDataSources.Item(0).GetValue("CANCELED", 0) == "N");

                
                //--------------------------------------------

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

        public static void cancellation(int docEntry, out string errorText)
        {
            errorText = null;

            try
            {
                Dictionary<string, string> wblDocInfo = BDO_Waybills.getWaybillDocumentInfo(docEntry, "UDO_F_BDOSFASTRD_D", out errorText);
                int wblDocEntry = Convert.ToInt32(wblDocInfo["DocEntry"]);

                if (wblDocEntry != 0)
                {
                    int answer = Program.uiApp.MessageBox(BDOSResources.getTranslate("DocumentLinkedToWaybillCancel"), 1, BDOSResources.getTranslate("Yes"), BDOSResources.getTranslate("No"), "");
                    string operation = answer == 1 ? "Update" : "Cancel";
                    BDO_Waybills.cancellation(wblDocEntry, operation, out errorText);
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
            top = top + height + 1;

            oForm.Items.Item("0_U_E").Left = 130;

            oForm.Items.Item("CardCodeS").Top = top;
            oForm.Items.Item("CardCodeE").Top = top;
            oForm.Items.Item("CardNameE").Top = top;
            oForm.Items.Item("CardLB").Top = top;

            top = top + height + 1;
            oForm.Items.Item("FLocCodeS").Top = top;
            oForm.Items.Item("FLocCodeE").Top = top;
            oForm.Items.Item("FLocNameE").Top = top;
            oForm.Items.Item("FLocLB").Top = top;

            top = top + height + 1;
            oForm.Items.Item("FEmplIDS").Top = top;
            oForm.Items.Item("FEmplIDE").Top = top;
            oForm.Items.Item("FEmplNameE").Top = top;
            oForm.Items.Item("FEmplLB").Top = top;

            top = top + height + 1;

            int topLeft = top;

            top = 6;
            //top = oItem.Top;
            oForm.Items.Item("CanceledS").Top = top;
            oForm.Items.Item("CanceledE").Top = top;

            top = top + height + 1;
            oForm.Items.Item("DocDateS").Top = top;
            oForm.Items.Item("DocDateE").Top = top;

            top = top + height + 1;
            oForm.Items.Item("TLocCodeS").Top = top;
            oForm.Items.Item("TLocCodeE").Top = top;
            oForm.Items.Item("TLocNameE").Top = top;
            oForm.Items.Item("TLocLB").Top = top;

            top = top + height + 1;
            oForm.Items.Item("TEmplIDS").Top = top;
            oForm.Items.Item("TEmplIDE").Top = top;
            oForm.Items.Item("TEmplNameE").Top = top;
            oForm.Items.Item("TEmplLB").Top = top;

            top = Math.Max(topLeft, top) + 2 * height + 1;

            top = top + height + 1;
            oForm.Items.Item("FillMTR").Top = top;
            //oForm.Items.Item("addMTR").Top = top;
            //oForm.Items.Item("delMTR").Top = top;

            int MTRWidth = oForm.Width - 15;
            top = top + height + 1;
            oItem = oForm.Items.Item("AssetMTR");
            oItem.Top = top;
            oItem.Width = MTRWidth;
            oItem.Height = oForm.Height / 3;

            // სვეტების ზომები 
            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("AssetMTR").Specific));
            SAPbouiCOM.Columns oColumns = oMatrix.Columns;
            SAPbouiCOM.Column oColumn;
            oColumn = oMatrix.Columns.Item("LineID");
            oColumn.Width = 20 - 1;
            MTRWidth = MTRWidth - 20 - 1;

            //სარდაფი
            top = oItem.Top + oItem.Height + 20;

            oItem = oForm.Items.Item("CreatorS");
            oItem.Top = top;
            oItem = oForm.Items.Item("CreatorE");
            oItem.Top = top;
            top = top + height + 1;

            oItem = oForm.Items.Item("RemarksS");
            oItem.Top = top;
            oItem = oForm.Items.Item("RemarksE");
            oItem.Top = top;

            top = top + 4 * height;

            oForm.Items.Item("1").Top = top;
            oForm.Items.Item("2").Top = top;

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

        public static void setVisibleFormItems(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            oForm.Freeze(true);

            try
            {
                SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("AssetMTR").Specific));
                SAPbouiCOM.Column oColumn = oMatrix.Columns.Item("LineID");
                SAPbouiCOM.DBDataSource oDBDataSource = oForm.DataSources.DBDataSources.Item("@BDOSFASTRD");
                oColumn.Editable = false;

                oForm.Items.Item("CardNameE").Enabled = false;
                oForm.Items.Item("FLocNameE").Enabled = false;
                oForm.Items.Item("FEmplNameE").Enabled = false;
                oForm.Items.Item("TLocNameE").Enabled = false;
                oForm.Items.Item("TEmplNameE").Enabled = false;

                if (openFormEvent)
                {
                    oForm.Items.Item("RemarksE").Click();
                }

                if (oDBDataSource.GetValue("DocStatus", 0) == "O")
                {
                    SAPbouiCOM.Item oItem = oForm.Items.Item("BDO_WblTxt");
                    oItem.Enabled = true;
                }

                
                Dictionary<string, string> activeDimensionsList = CommonFunctions.getActiveDimensionsList(out errorText);
                for (int i = 1; i <= activeDimensionsList.Count; i++)
                {
                    oForm.Items.Item("DistrRul" + i + "E").Enabled = (oDBDataSource.GetValue("U_ChngDstRl" + i, 0) == "Y");
                }
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

        public static void chooseFromList(SAPbouiCOM.Form oForm, SAPbouiCOM.IChooseFromListEvent oCFLEvento, SAPbouiCOM.ItemEvent pVal, out string errorText)
        {
            errorText = null;

            bool beforeAction = pVal.BeforeAction;
            int row = pVal.Row;
            string sCFL_ID = oCFLEvento.ChooseFromListUID;

            try
            {
                SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(sCFL_ID);
                SAPbouiCOM.DBDataSources oDBDataSources = oForm.DataSources.DBDataSources;

                if (beforeAction == false)
                {
                    SAPbouiCOM.DataTable oDataTable = null;
                    oDataTable = oCFLEvento.SelectedObjects;

                    if (oDataTable != null)
                    {
                        if (sCFL_ID == "BusinessPartner_CFL")
                        {
                            oDBDataSources.Item("@BDOSFASTRD").SetValue("U_CardCode", 0, oDataTable.GetValue("CardCode", 0));
                            oDBDataSources.Item("@BDOSFASTRD").SetValue("U_CardName", 0, oDataTable.GetValue("CardName", 0));
                        }
                        else if (sCFL_ID == "FromLoc_CFL")
                        {
                            oDBDataSources.Item("@BDOSFASTRD").SetValue("U_FLocCode", 0, oDataTable.GetValue("Code", 0));
                            oDBDataSources.Item("@BDOSFASTRD").SetValue("U_FLocName", 0, oDataTable.GetValue("Location", 0));

                            //addMatrixRow(oForm, out errorText);
                        }
                        else if (sCFL_ID == "ToLoc_CFL")
                        {
                            oDBDataSources.Item("@BDOSFASTRD").SetValue("U_TLocCode", 0, oDataTable.GetValue("Code", 0));
                            oDBDataSources.Item("@BDOSFASTRD").SetValue("U_TLocName", 0, oDataTable.GetValue("Location", 0));

                            fillToValues(oForm, "Location", out errorText);
                        }
                        else if (sCFL_ID == "FromEmpl_CFL")
                        {
                            oDBDataSources.Item("@BDOSFASTRD").SetValue("U_FEmplID", 0, oDataTable.GetValue("empID", 0));
                            oDBDataSources.Item("@BDOSFASTRD").SetValue("U_FEmplName", 0, oDataTable.GetValue("firstName", 0) + " " + oDataTable.GetValue("lastName", 0));
                        }
                        else if (sCFL_ID == "ToEmpl_CFL")
                        {
                            oDBDataSources.Item("@BDOSFASTRD").SetValue("U_TEmplID", 0, oDataTable.GetValue("empID", 0));
                            oDBDataSources.Item("@BDOSFASTRD").SetValue("U_TEmplName", 0, oDataTable.GetValue("firstName", 0) + " " + oDataTable.GetValue("lastName", 0));

                            fillToValues(oForm, "Employee", out errorText);
                        }
                        else if (sCFL_ID == "Project_CFL")
                        {
                            oDBDataSources.Item("@BDOSFASTRD").SetValue("U_PrjCode", 0, oDataTable.GetValue("PrjCode", 0));
                        }
                        else if (sCFL_ID.Length >= 2 && sCFL_ID.Substring(0, sCFL_ID.Length - 2) == "Rule_CFL")
                        {
                            SAPbouiCOM.DataTable oDataTableSelectedObjects = oCFLEvento.SelectedObjects;
                            string val = oDataTableSelectedObjects.GetValue("OcrCode", 0);

                            oDBDataSources.Item("@BDOSFASTRD").SetValue("U_DistrRule" + sCFL_ID.Substring(sCFL_ID.Length - 2, 1), 0, oDataTable.GetValue("OcrCode", 0));
                        }

                        else if (sCFL_ID == "ItemMTR_CFL")
                        {
                            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("AssetMTR").Specific));
                            oMatrix.FlushToDataSource();
                            oDBDataSources.Item("@BDOSFASTR1").SetValue("U_ItemCode", pVal.Row - 1, oDataTable.GetValue("ItemCode", 0));
                            oMatrix.LoadFromDataSource();

                            fillAssetValues(oForm, pVal.Row - 1, out errorText);

                            addMatrixRow(oForm, out errorText);
                        }
                        else if (sCFL_ID == "ToLocMTR_CFL")
                        {
                            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("AssetMTR").Specific));
                            oMatrix.FlushToDataSource();
                            oDBDataSources.Item("@BDOSFASTR1").SetValue("U_TLocCode", pVal.Row - 1, oDataTable.GetValue("Code", 0));
                            oDBDataSources.Item("@BDOSFASTR1").SetValue("U_TLocName", pVal.Row - 1, oDataTable.GetValue("Location", 0));
                            oMatrix.LoadFromDataSource();
                        }
                        else if (sCFL_ID == "ToEmplMTR_CFL")
                        {
                            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("AssetMTR").Specific));
                            oMatrix.FlushToDataSource();
                            oDBDataSources.Item("@BDOSFASTR1").SetValue("U_TEmplID", pVal.Row - 1, oDataTable.GetValue("empID", 0));
                            oDBDataSources.Item("@BDOSFASTR1").SetValue("U_TEmplName", pVal.Row - 1, oDataTable.GetValue("firstName", 0) + " " + oDataTable.GetValue("lastName", 0));
                            oMatrix.LoadFromDataSource();
                        }

                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                        }
                    }
                }
                else
                {
                    if (sCFL_ID.Length > 1 && sCFL_ID.Substring(0, sCFL_ID.Length - 2) == "Rule_CFL")
                    {
                        oForm.Freeze(true);
                        string dimensionCode = sCFL_ID.Substring(sCFL_ID.Length - 2, 1);
                        SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();

                        string strDocDate = oDBDataSources.Item("@BDOSFASTRD").GetValue("U_DocDate", 0);
                        DateTime DocDate = DateTime.TryParseExact(strDocDate, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DocDate) == false ? DateTime.Now : DocDate;

                        SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        string query = @"SELECT
	                                     ""OCR1"".""OcrCode"",
	                                     ""OOCR"".""DimCode"" 
                                    FROM ""OCR1"" 
                                    LEFT JOIN ""OOCR"" ON ""OCR1"".""OcrCode"" = ""OOCR"".""OcrCode"" 
                                    WHERE ""OOCR"".""DimCode"" = " + dimensionCode + @" AND ""ValidFrom"" <= '" + DocDate.ToString("yyyyMMdd") +
                                                                                         @"' AND (""ValidTo"" > '" + DocDate.ToString("yyyyMMdd") + @"' OR " + @" ""ValidTo"" IS NULL)";

                        try
                        {
                            oRecordSet.DoQuery(query);
                            int recordCount = oRecordSet.RecordCount;
                            int i = 1;

                            while (!oRecordSet.EoF)
                            {
                                SAPbouiCOM.Condition oCon = oCons.Add();
                                oCon.Alias = "OcrCode";
                                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                                oCon.CondVal = oRecordSet.Fields.Item("OcrCode").Value.ToString();
                                oCon.Relationship = (i == recordCount) ? SAPbouiCOM.BoConditionRelationship.cr_NONE : SAPbouiCOM.BoConditionRelationship.cr_OR;

                                i = i + 1;
                                oRecordSet.MoveNext();
                            }

                            //თუ არცერთი შეესაბამება ცარიელზე გავიდეს
                            if (oCons.Count == 0)
                            {
                                SAPbouiCOM.Condition oCon = oCons.Add();
                                oCon.Alias = "OcrCode";
                                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                                oCon.CondVal = "";
                            }

                            oCFL.SetConditions(oCons);
                        }
                        catch (Exception ex)
                        {
                            errorText = ex.Message;
                        }

                        oForm.Freeze(false);

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
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                }
                GC.Collect();
            }
        }

        public static void addMatrixRow(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            oForm.Freeze(true);

            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("AssetMTR").Specific));
            SAPbouiCOM.DBDataSource mtrDataSource = oForm.DataSources.DBDataSources.Item("@BDOSFASTR1");

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
                SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("AssetMTR").Specific));
                SAPbouiCOM.DBDataSource mtrDataSource = oForm.DataSources.DBDataSources.Item("@BDOSFASTR1");

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

        public static void fillToValues(SAPbouiCOM.Form oForm, string toField, out string errorText)
        {
            errorText = null;

            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("AssetMTR").Specific));
            SAPbouiCOM.DBDataSource oDataSource = oForm.DataSources.DBDataSources.Item("@BDOSFASTRD");
            SAPbouiCOM.DBDataSource mtrDataSource = oForm.DataSources.DBDataSources.Item("@BDOSFASTR1");

            try
            {
                oForm.Freeze(true);
                oMatrix.FlushToDataSource();

                string toCode;
                string toName;
                if (toField == "Location")
                {
                    toCode = oDataSource.GetValue("U_TLocCode", 0);
                    toName = oDataSource.GetValue("U_TLocName", 0);
                }
                else
                {
                    toCode = oDataSource.GetValue("U_TEmplID", 0);
                    toName = oDataSource.GetValue("U_TEmplName", 0);
                }

                for (int i = 0; i < mtrDataSource.Size; i++)
                {
                    if (!string.IsNullOrEmpty(mtrDataSource.GetValue("U_ItemCode", i)))
                    {
                        if (toField == "Location")
                        {
                            mtrDataSource.SetValue("U_TLocCode", i, toCode);
                            mtrDataSource.SetValue("U_TLocName", i, toName);
                        }
                        else
                        {
                            mtrDataSource.SetValue("U_TEmplID", i, toCode);
                            mtrDataSource.SetValue("U_TEmplName", i, toName);
                        }
                    }
                }

                oMatrix.LoadFromDataSource();
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
            finally
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                }

                oForm.Freeze(false);
            }

        }

        public static void fillAssetValues(SAPbouiCOM.Form oForm, int row, out string errorText)
        {
            errorText = null;
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("AssetMTR").Specific));
            SAPbouiCOM.DBDataSource oDataSource = oForm.DataSources.DBDataSources.Item("@BDOSFASTRD");
            SAPbouiCOM.DBDataSource mtrDataSource = oForm.DataSources.DBDataSources.Item("@BDOSFASTR1");

            try
            {
                oForm.Freeze(true);
                oMatrix.FlushToDataSource();

                string itemcode = mtrDataSource.GetValue("U_ItemCode", row);
                if (!string.IsNullOrEmpty(itemcode))
                {
                    DateTime dt= DateTime.TryParseExact(oForm.Items.Item("DocDateE").Specific.Value.ToString(), "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out dt) == false ? DateTime.Now : dt;
                    string query = getAssetQuery(dt, itemcode, "", "");

                    if (Program.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                    {
                        query = query.Replace("ISNULL", "IFNULL");
                    }

                    oRecordSet.DoQuery(query);

                    if (!oRecordSet.EoF)
                    {
                        mtrDataSource.SetValue("U_ItemName", row, oRecordSet.Fields.Item("ItemName").Value);
                        mtrDataSource.SetValue("U_SerNo", row, oRecordSet.Fields.Item("AssetSerNo").Value);
                        mtrDataSource.SetValue("U_Quantity", row, oRecordSet.Fields.Item("Quantity").Value);
                        mtrDataSource.SetValue("U_APC", row, oRecordSet.Fields.Item("APC").Value);
                        mtrDataSource.SetValue("U_FLocCode", row, oRecordSet.Fields.Item("LocCode").Value);
                        mtrDataSource.SetValue("U_FLocName", row, oRecordSet.Fields.Item("LocName").Value);
                        mtrDataSource.SetValue("U_TLocCode", row, oDataSource.GetValue("U_TLocCode", 0));
                        mtrDataSource.SetValue("U_TLocName", row, oDataSource.GetValue("U_TLocName", 0));
                        mtrDataSource.SetValue("U_FEmplID", row, oRecordSet.Fields.Item("EmplID").Value);
                        mtrDataSource.SetValue("U_FEmplName", row, oRecordSet.Fields.Item("EmplName").Value);
                        mtrDataSource.SetValue("U_TEmplID", row, oDataSource.GetValue("U_TEmplID", 0));
                        mtrDataSource.SetValue("U_TEmplName", row, oDataSource.GetValue("U_TEmplName", 0));
                        //mtrDataSource.SetValue("U_UomCode", row, oRecordSet.Fields.Item("ItemName").Value);
                        mtrDataSource.SetValue("U_UomName", row, oRecordSet.Fields.Item("InvntryUom").Value);
                    }
                }

                oMatrix.LoadFromDataSource();
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                oRecordSet = null;

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                }

                oForm.Freeze(false);
            }

        }

        public static string getAssetQuery(DateTime docDate, string itemCode, string fLocCode, string fEmplID)
        {
            string ThisYear = docDate.Year.ToString();
            string query = @"SELECT
	                                 ""Assets"".* 
                                FROM (SELECT
	                                 ""OITM"".""ItemCode"",
	                                 ""ItemName"",
	                                 ""AssetSerNo"",
	                                 ""Employee"",
	                                 ""InvntryUom"",
	                                 ISNULL(CAST(""OLCT"".""Code"" AS nvarchar), '') AS ""LocCode"",
	                                 ""OLCT"".""Location"" AS ""LocName"",
	                                 ISNULL(CAST(""OHEM"".""empID"" AS nvarchar), '') AS ""EmplID"",
	                                 CONCAT(CONCAT(""OHEM"".""firstName"", ' '), ""OHEM"".""lastName"") AS ""EmplName"",
	 
	                                 ""AssetCost"".""APC"",
	                                 ""AssetCost"".""Quantity""
		 
	                                FROM ""OITM""	
	                                LEFT JOIN ""OLCT"" ON ""OITM"".""Location"" = ""OLCT"".""Code"" 
	                                LEFT JOIN ""OHEM"" ON ""OITM"".""Employee"" = ""OHEM"".""empID"" 
	                                LEFT JOIN (SELECT
				                                 ""ItemCode"",
				                                 SUM (""APC"") AS ""APC"",
				                                 SUM(""Quantity"") AS ""Quantity""
			                                FROM
			                                (SELECT
				                                 ""ItemCode"",
				                                 ""APC"" ,
				                                 ""Quantity"" ,
                                                   ""PeriodCat""
			                                FROM (SELECT
					                                 ""ItemCode"",
					                                 ""APC"" ,
					                                 ""Quantity"",
					                                 ""PeriodCat"",
				 	                                 RANK ( ) OVER (PARTITION BY ""ItemCode"" ORDER BY ""PeriodCat"" Desc) AS ""Rank"" 
				                                FROM ""ITM8"" ) AS ""Items""
				                                WHERE ""Rank"" = 1 and ""PeriodCat""='" + ThisYear+ @"' 
			
			                                UNION ALL 
			                                SELECT
				                                 ""ItemCode"",
				                                 ""APC"" ,
				                                 ""Qty"" ,
                                            ""PeriodCat""
			                                FROM ""FIX1"" 
			                                INNER JOIN ""OFIX"" ON ""FIX1"".""AbsEntry"" = ""OFIX"".""AbsEntry"" AND ""Canceled"" = 'N'
			                                where ""PeriodCat""='" + ThisYear + @"' 
			                                UNION ALL 
			                                SELECT
				                                 ""RecvAsst"",
				                                 -1*""APC"" ,
				                                 -1*""Qty"" ,
                                            ""PeriodCat""
			                                FROM ""FIX1"" 
			                                INNER JOIN ""OFIX"" ON ""FIX1"".""AbsEntry"" = ""OFIX"".""AbsEntry"" AND ""Canceled"" = 'N' AND ISNULL(""RecvAsst"",
				                                 '') <> ''
                                              where ""PeriodCat""='" + ThisYear + @"' 
                                            ) AS ""UnionAssets""			
			                                GROUP BY ""ItemCode"") AS ""AssetCost""	ON ""OITM"".""ItemCode"" = ""AssetCost"".""ItemCode""
	
	                                WHERE ""OITM"".""ItemType"" = 'F' AND ""OITM"".""AsstStatus"" = 'A' " +

                                     (docDate == new DateTime() ? "" : @" AND ""OITM"".""CapDate"" <= '" + docDate.ToString("yyyyMMdd") + "'") +
                                    (string.IsNullOrEmpty(itemCode) ? "" : @" AND ""OITM"".""ItemCode"" = N'" + itemCode + "'") +
                                    (string.IsNullOrEmpty(fLocCode) ? "" : @" AND CAST(""OLCT"".""Code"" AS nvarchar) = N'" + fLocCode + "'") +
                                    (string.IsNullOrEmpty(fEmplID) ? "" : @" AND CAST(""OHEM"".""empID"" AS nvarchar) = N'" + fEmplID + "'") +                                    

                                    @" ) AS ""Assets""                                    
                                    ORDER BY ""ItemCode""";

            if (Program.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
            {
                query = query.Replace("ISNULL", "IFNULL");
                query = query.Replace("ROW_NUMBER", "RANK");
            }

            return query;
        }

        public static void fillAssets(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            SAPbouiCOM.DBDataSource oDataSource = oForm.DataSources.DBDataSources.Item("@BDOSFASTRD");
            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("AssetMTR").Specific));
            SAPbouiCOM.DBDataSource mtrDataSource = oForm.DataSources.DBDataSources.Item("@BDOSFASTR1");

            string fLocCode = oDataSource.GetValue("U_FLocCode", 0);
            string fEmplID = oDataSource.GetValue("U_FEmplID", 0);

            if (string.IsNullOrEmpty(fLocCode) && string.IsNullOrEmpty(fEmplID))
            {
                errorText = BDOSResources.getTranslate("FromLocation") + " " + BDOSResources.getTranslate("Or") + " " + BDOSResources.getTranslate("FromEmployee") + " " + BDOSResources.getTranslate("YouCantLeaveEmpty");
                Program.uiApp.StatusBar.SetSystemMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return;
            }

            string strDocDate = oDataSource.GetValue("U_DocDate", 0);
            if (string.IsNullOrEmpty(strDocDate))
            {
                errorText = BDOSResources.getTranslate("PostingDate") + " " + BDOSResources.getTranslate("YouCantLeaveEmpty");
                Program.uiApp.StatusBar.SetSystemMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return;
            }
            DateTime DocDate = DateTime.TryParseExact(strDocDate, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DocDate) == false ? DateTime.Now : DocDate;

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                oMatrix.FlushToDataSource();

                if (mtrDataSource.Size > 1)
                {
                    int answer = Program.uiApp.MessageBox(BDOSResources.getTranslate("DoYouWantToClearTheTable") + "?", 1, BDOSResources.getTranslate("Yes"), BDOSResources.getTranslate("No"), "");

                    if (answer == 1)
                    {
                        mtrDataSource.Clear();
                        //while (mtrDataSource.Size > 0)
                        //{
                        //    mtrDataSource.RemoveRecord(mtrDataSource.Size - 1);
                        //}
                    }
                }
                
                string query = getAssetQuery(DocDate, "", fLocCode, fEmplID);
                oRecordSet.DoQuery(query);
                int i;

                string tLocCode = oDataSource.GetValue("U_TLocCode", 0);
                string tLocName = oDataSource.GetValue("U_TLocName", 0);
                string tEmplID = oDataSource.GetValue("U_TEmplID", 0);
                string tEmplName = oDataSource.GetValue("U_TEmplName", 0);

                while (!oRecordSet.EoF)
                {
                    i = mtrDataSource.Size;
                    mtrDataSource.InsertRecord(i);
                    mtrDataSource.SetValue("LineId", i - 1, i.ToString());

                    mtrDataSource.SetValue("U_ItemCode", i - 1, oRecordSet.Fields.Item("ItemCode").Value);
                    mtrDataSource.SetValue("U_ItemName", i - 1, oRecordSet.Fields.Item("ItemName").Value);
                    mtrDataSource.SetValue("U_SerNo", i - 1, oRecordSet.Fields.Item("AssetSerNo").Value);
                    mtrDataSource.SetValue("U_Quantity", i - 1, oRecordSet.Fields.Item("Quantity").Value);
                    mtrDataSource.SetValue("U_APC", i - 1, oRecordSet.Fields.Item("APC").Value);
                    mtrDataSource.SetValue("U_FLocCode", i - 1, oRecordSet.Fields.Item("LocCode").Value);
                    mtrDataSource.SetValue("U_FLocName", i - 1, oRecordSet.Fields.Item("LocName").Value);
                    mtrDataSource.SetValue("U_TLocCode", i - 1, tLocCode);
                    mtrDataSource.SetValue("U_TLocName", i - 1, tLocName);
                    mtrDataSource.SetValue("U_FEmplID", i - 1, oRecordSet.Fields.Item("EmplID").Value);
                    mtrDataSource.SetValue("U_FEmplName", i - 1, oRecordSet.Fields.Item("EmplName").Value);
                    mtrDataSource.SetValue("U_TEmplID", i - 1, tEmplID);
                    mtrDataSource.SetValue("U_TEmplName", i - 1, tEmplName);
                    mtrDataSource.SetValue("U_UomName", i - 1, oRecordSet.Fields.Item("InvntryUom").Value);

                    oRecordSet.MoveNext();
                }
                oMatrix.LoadFromDataSource();
            }
            catch
            { }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                oRecordSet = null;

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                }

                oForm.Update();
                oForm.Freeze(false);
            }
        }

        public static void updateAsset(SAPbouiCOM.Form oForm, bool cancel, out string errorText)
        {
            errorText = null;

            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("AssetMTR").Specific));
            SAPbouiCOM.DBDataSource oDataSource = oForm.DataSources.DBDataSources.Item("@BDOSFASTRD");
            SAPbouiCOM.DBDataSource mtrDataSource = oForm.DataSources.DBDataSources.Item("@BDOSFASTR1");

            string strDocDate = oDataSource.GetValue("U_DocDate", 0);
            DateTime DocDate = DateTime.TryParseExact(strDocDate, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DocDate) == false ? DateTime.Now : DocDate;

            string prjCode = oDataSource.GetValue("U_PrjCode", 0);

            bool chngRule1 = oDataSource.GetValue("U_ChngDstRl1", 0) == "Y";
            bool chngRule2 = oDataSource.GetValue("U_ChngDstRl2", 0) == "Y";
            bool chngRule3 = oDataSource.GetValue("U_ChngDstRl3", 0) == "Y";
            bool chngRule4 = oDataSource.GetValue("U_ChngDstRl4", 0) == "Y";
            bool chngRule5 = oDataSource.GetValue("U_ChngDstRl5", 0) == "Y";

            string distrRule1 = oDataSource.GetValue("U_DistrRule1", 0);
            string distrRule2 = oDataSource.GetValue("U_DistrRule2", 0);
            string distrRule3 = oDataSource.GetValue("U_DistrRule3", 0);
            string distrRule4 = oDataSource.GetValue("U_DistrRule4", 0);
            string distrRule5 = oDataSource.GetValue("U_DistrRule5", 0);

            bool Commit = true;
            CommonFunctions.StartTransaction();

            try
            {
                oForm.Freeze(true);
                oMatrix.FlushToDataSource();

                for (int i = 0; i < mtrDataSource.Size; i++)
                {
                    string itemCode = mtrDataSource.GetValue("U_ItemCode", i);
                    if (!string.IsNullOrEmpty(itemCode))
                    {
                        SAPbobsCOM.Items oItem = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems);
                        oItem.GetByKey(itemCode);

                        if (!string.IsNullOrEmpty(prjCode))
                        {
                            SAPbobsCOM.ItemsProjects oItemsProjects = oItem.Projects;
                            oItemsProjects.SetCurrentLine(oItemsProjects.Count - 1);
                            if (cancel)
                            {
                                //თუ cancel-ია და ამ დოკუმენტის პროექტი არის აითემის ბოლო პროექტი, უნდა წაიშალოს ეს ჩანაწერი
                                if (oItemsProjects.ValidFrom == DocDate)
                                {
                                    oItemsProjects.Delete();
                                    if (oItemsProjects.Count > 0)
                                    {
                                        oItemsProjects.SetCurrentLine(oItemsProjects.Count - 1);
                                        oItemsProjects.ValidTo = new DateTime();
                                    }
                                }
                            }
                            else
                            {
                                // თუ add-ია და აითემის ბოლო პროექტი განსხვავებულია დოკუმენტის პროექტისგან
                                if (oItem.Projects.Project != prjCode)
                                {
                                    if (!string.IsNullOrEmpty(oItemsProjects.Project))
                                    {
                                        oItemsProjects.ValidTo = DocDate.AddDays(-1);
                                        oItemsProjects.Add();
                                    }
                                    oItemsProjects.SetCurrentLine(oItemsProjects.Count - 1);
                                    oItemsProjects.ValidFrom = DocDate;
                                    oItemsProjects.Project = prjCode;
                                }
                            }
                        }

                        if (chngRule1 || chngRule2 || chngRule3 || chngRule4 || chngRule5)
                        {
                            SAPbobsCOM.ItemsDistributionRules oItemsDistributionRules = oItem.DistributionRules;
                            oItemsDistributionRules.SetCurrentLine(oItemsDistributionRules.Count - 1);
                            if (cancel)
                            {
                                //თუ cancel-ია და ამ დოკუმენტის პროექტი არის აითემის ბოლო პროექტი, უნდა წაიშალოს ეს ჩანაწერი
                                if (oItemsDistributionRules.ValidFrom == DocDate)
                                {
                                    oItemsDistributionRules.Delete();
                                    if (oItemsDistributionRules.Count > 0)
                                    {
                                        oItemsDistributionRules.SetCurrentLine(oItemsDistributionRules.Count - 1);
                                        oItemsDistributionRules.ValidTo = new DateTime();
                                    }
                                }
                            }
                            else
                            {
                                string curDistrRule1 = oItemsDistributionRules.DistributionRule;
                                string curDistrRule2 = oItemsDistributionRules.DistributionRule2;
                                string curDistrRule3 = oItemsDistributionRules.DistributionRule3;
                                string curDistrRule4 = oItemsDistributionRules.DistributionRule4;
                                string curDistrRule5 = oItemsDistributionRules.DistributionRule5;

                                if (!(string.IsNullOrEmpty(curDistrRule1) && string.IsNullOrEmpty(curDistrRule2) && string.IsNullOrEmpty(curDistrRule3) &&
                                      string.IsNullOrEmpty(curDistrRule4) && string.IsNullOrEmpty(curDistrRule5)))
                                {
                                oItemsDistributionRules.ValidTo = DocDate.AddDays(-1);
                                    oItemsDistributionRules.Add();
                                }

                                oItemsDistributionRules.SetCurrentLine(oItemsDistributionRules.Count - 1);
                                oItemsDistributionRules.ValidFrom = DocDate;

                                oItemsDistributionRules.DistributionRule = (chngRule1 ? distrRule1 : curDistrRule1);
                                oItemsDistributionRules.DistributionRule2 = (chngRule2 ? distrRule2 : curDistrRule2);
                                oItemsDistributionRules.DistributionRule3 = (chngRule3 ? distrRule3 : curDistrRule3);
                                oItemsDistributionRules.DistributionRule4 = (chngRule4 ? distrRule4 : curDistrRule4);
                                oItemsDistributionRules.DistributionRule5 = (chngRule5 ? distrRule5 : curDistrRule5);
                            }
                        }

                        string toLocation = (cancel ? mtrDataSource.GetValue("U_FlocCode", i) : mtrDataSource.GetValue("U_TlocCode", i));
                        string toEmployee = (cancel ? mtrDataSource.GetValue("U_FEmplID", i) : mtrDataSource.GetValue("U_TEmplID", i));
                        if (!string.IsNullOrEmpty(toLocation))
                        {
                            oItem.Location = Convert.ToInt32(toLocation);
                        }
                        if (!string.IsNullOrEmpty(toEmployee))
                        {
                            oItem.Employee = Convert.ToInt32(toEmployee);
                        }

                        int succ = oItem.Update();
                        if (succ != 0)
                        {
                            int errCode;
                            string errMsg;
                            Program.oCompany.GetLastError(out errCode, out errMsg);
                            errorText = BDOSResources.getTranslate("CannotUpdateAsset") + ": " + itemCode + "; " + BDOSResources.getTranslate("Reason") + ": " + errMsg;

                            Commit = false;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
                Commit = false;
            }
            finally
            {
                oForm.Freeze(false);

                if (Program.oCompany.InTransaction)
                {
                    if (Commit)
                    {
                        CommonFunctions.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                    }
                    else
                    {
                        CommonFunctions.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                    }
                }
            }
        }

        public static void checkDoc(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("AssetMTR").Specific));
            SAPbouiCOM.DBDataSource oDataSource = oForm.DataSources.DBDataSources.Item("@BDOSFASTRD");
            SAPbouiCOM.DBDataSource mtrDataSource = oForm.DataSources.DBDataSources.Item("@BDOSFASTR1");

            string strDocDate = oDataSource.GetValue("U_DocDate", 0);
            if (string.IsNullOrEmpty(strDocDate))
            {
                errorText = BDOSResources.getTranslate("PostingDate") + " " + BDOSResources.getTranslate("YouCantLeaveEmpty");
                return;
            }

            DateTime docDate = DateTime.TryParseExact(strDocDate, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out docDate) == false ? DateTime.Now : docDate;

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string query;

            try
            {
                oForm.Freeze(true);
                oMatrix.FlushToDataSource();

                for (int i = 0; i < mtrDataSource.Size; i++)
                {
                    string itemCode = mtrDataSource.GetValue("U_ItemCode", i);
                    if (string.IsNullOrEmpty(itemCode))
                    {
                        if (i == mtrDataSource.Size - 1)
                        {
                            mtrDataSource.RemoveRecord(i);
                        }
                        else
                        {
                        errorText = BDOSResources.getTranslate("ItemCode") + " " + BDOSResources.getTranslate("YouCantLeaveEmpty");
                        break;
                    }
                    }
                    else
                    {
                    query = @"SELECT * 
                                    FROM ""@BDOSFASTR1""
                                    INNER JOIN ""@BDOSFASTRD"" ON ""@BDOSFASTR1"".""DocEntry"" = ""@BDOSFASTRD"".""DocEntry""

                                    WHERE ""@BDOSFASTRD"".""Canceled"" = 'N' AND 
	                                        ""@BDOSFASTR1"".""U_ItemCode"" = N'" + itemCode +
                                                @"' AND ""@BDOSFASTRD"".""U_DocDate"" >= '" + docDate.ToString("yyyyMMdd") + "'";

                    oRecordSet.DoQuery(query);

                    if (!oRecordSet.EoF)
                    {
                            errorText = BDOSResources.getTranslate("TransferDocumentAlreadyExistsForAsset") + " '" + itemCode + "' ";
                            break;
                        }
                    }
                }

                oMatrix.LoadFromDataSource();
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
            }
        }

    }
}
