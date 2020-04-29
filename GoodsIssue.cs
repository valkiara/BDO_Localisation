using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Globalization;
using System.Data;

namespace BDO_Localisation_AddOn
{
    static partial class GoodsIssue
    {
        public static bool ProfitTaxTypeIsSharing = false;
        public static string itemCodeOld;
        public static string warehouseOld;

        public static void createUserFields(out string errorText)
        {
            errorText = null;
            Dictionary<string, object> fieldskeysMap;

            //მოგების გადასახადი
            fieldskeysMap = new Dictionary<string, object>(); //იბეგრება განაწილებული მოგებით
            fieldskeysMap.Add("Name", "liablePrTx");
            fieldskeysMap.Add("TableName", "OIGE");
            fieldskeysMap.Add("Description", "Liable to Profit Tax");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 1);
            fieldskeysMap.Add("DefaultValue", "N");

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //იბეგრება დღგ-ით
            fieldskeysMap.Add("Name", "liableVat");
            fieldskeysMap.Add("TableName", "OIGE");
            fieldskeysMap.Add("Description", "Liable to Vat");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 1);
            fieldskeysMap.Add("DefaultValue", "N");

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "VatExpAcct");
            fieldskeysMap.Add("TableName", "OIGE");
            fieldskeysMap.Add("Description", "Vat Expense Account");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // დაბეგვრის ობიექტი 
            fieldskeysMap.Add("Name", "prBase");
            fieldskeysMap.Add("TableName", "OIGE");
            fieldskeysMap.Add("Description", "Profit Base");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 15);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // დაბეგვრის ობიექტის სახელი
            fieldskeysMap.Add("Name", "prBsDscr");
            fieldskeysMap.Add("TableName", "OIGE");
            fieldskeysMap.Add("Description", "Profit Base DEscription");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 100);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSDstRl1");
            fieldskeysMap.Add("TableName", "OIGE");
            fieldskeysMap.Add("Description", "Distr.Rule 1");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);
            fieldskeysMap.Add("Mandatory", SAPbobsCOM.BoYesNoEnum.tNO);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSDstRl2");
            fieldskeysMap.Add("TableName", "OIGE");
            fieldskeysMap.Add("Description", "Distr.Rule 2");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);
            fieldskeysMap.Add("Mandatory", SAPbobsCOM.BoYesNoEnum.tNO);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSDstRl3");
            fieldskeysMap.Add("TableName", "OIGE");
            fieldskeysMap.Add("Description", "Distr.Rule 3");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);
            fieldskeysMap.Add("Mandatory", SAPbobsCOM.BoYesNoEnum.tNO);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSDstRl4");
            fieldskeysMap.Add("TableName", "OIGE");
            fieldskeysMap.Add("Description", "Distr.Rule 4");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);
            fieldskeysMap.Add("Mandatory", SAPbobsCOM.BoYesNoEnum.tNO);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSDstRl5");
            fieldskeysMap.Add("TableName", "OIGE");
            fieldskeysMap.Add("Description", "Distr.Rule 5");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);
            fieldskeysMap.Add("Mandatory", SAPbobsCOM.BoYesNoEnum.tNO);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSPrjCod");
            fieldskeysMap.Add("TableName", "OIGE");
            fieldskeysMap.Add("Description", "Project");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);
            fieldskeysMap.Add("Mandatory", SAPbobsCOM.BoYesNoEnum.tNO);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSVatCod");
            fieldskeysMap.Add("TableName", "OIGE");
            fieldskeysMap.Add("Description", "Vat code");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);
            fieldskeysMap.Add("Mandatory", SAPbobsCOM.BoYesNoEnum.tNO);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSInStck");
            fieldskeysMap.Add("TableName", "IGE1");
            fieldskeysMap.Add("Description", "In Stock");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Quantity);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "amtPrTx"); //მოგების გადასახადი
            fieldskeysMap.Add("TableName", "OIGE");
            fieldskeysMap.Add("Description", "Profit Tax Amount");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "amtVat"); //დღგ-ის თანხა
            fieldskeysMap.Add("TableName", "OIGE");
            fieldskeysMap.Add("Description", "Vat Amount");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);
        }

        public static void createFormItems(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            ProfitTaxTypeIsSharing = ProfitTax.ProfitTaxTypeIsSharing();

            Dictionary<string, object> formItems;
            string itemName = "";

            SAPbouiCOM.Item oItem = oForm.Items.Item("20");
            int left = oItem.Left;
            int height = oItem.Height;
            double top = oItem.Top;
            int width_ = oItem.Width;
            int pane = 3;

            top = top + height + 5;

            formItems = new Dictionary<string, object>();
            itemName = "liableVat";
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OIGE");
            formItems.Add("Alias", "U_liableVat");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
            formItems.Add("Length", 1);
            formItems.Add("Left", left);
            formItems.Add("Width", 150);
            formItems.Add("Top", top);
            formItems.Add("Height", 14);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("LiableToVat"));
            formItems.Add("ValOff", "N");
            formItems.Add("ValOn", "Y");
            formItems.Add("DisplayDesc", true);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 5;

            formItems = new Dictionary<string, object>();
            itemName = "liablePrTx";
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OIGE");
            formItems.Add("Alias", "U_liablePrTx");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
            formItems.Add("Length", 1);
            formItems.Add("Left", left);
            formItems.Add("Width", 200);
            formItems.Add("Top", top);
            formItems.Add("Height", 14);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("LiableToProfitTax"));
            formItems.Add("ValOff", "N");
            formItems.Add("ValOn", "Y");
            formItems.Add("DisplayDesc", true);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            //გადასახადები (ჩანართი)
            SAPbouiCOM.Item oFolder = oForm.Items.Item("1320000080");
            formItems = new Dictionary<string, object>();
            itemName = "Taxes";
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_FOLDER);
            formItems.Add("Left", oFolder.Left + oFolder.Width);
            formItems.Add("Width", oFolder.Width);
            formItems.Add("Top", oFolder.Top);
            formItems.Add("Height", oFolder.Height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Taxes"));
            formItems.Add("Pane", pane);
            formItems.Add("ValOn", "0");
            formItems.Add("ValOff", itemName);
            formItems.Add("GroupWith", "1320000080");
            formItems.Add("AffectsFormMode", false);
            formItems.Add("Description", BDOSResources.getTranslate("Taxes"));

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            oItem = oForm.Items.Item("Taxes");
            top = oItem.Top + 30;
            int left_s = 13;
            int width_s = 100;
            int left_e = 130;
            int width_e = 180;

            formItems = new Dictionary<string, object>();
            itemName = "VtExpAccE";
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("Caption", BDOSResources.getTranslate("VatExpense") + " " + BDOSResources.getTranslate("Acct"));
            formItems.Add("UID", itemName);
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);
            formItems.Add("LinkTo", "VatExpAcct");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            bool multiSelection = false;
            string objectType = "1";
            string uniqueID_CFL = "Acct_CFL";
            FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_CFL);

            //პირობის დადება ანგარიშის არჩევის სიაზე
            SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(uniqueID_CFL);
            SAPbouiCOM.Conditions oCons = oCFL.GetConditions();
            SAPbouiCOM.Condition oCon = oCons.Add();
            oCon.Alias = "Postable";
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCon.CondVal = "Y"; //Active Account
            oCFL.SetConditions(oCons);

            formItems = new Dictionary<string, object>();
            itemName = "VatExpAcct";
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OIGE");
            formItems.Add("Alias", "U_VatExpAcct");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
            formItems.Add("DisplayDesc", true);
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);
            formItems.Add("Description", BDOSResources.getTranslate("VatExpense") + " " + BDOSResources.getTranslate("Account"));
            formItems.Add("ChooseFromListUID", uniqueID_CFL);
            formItems.Add("ChooseFromListAlias", "AcctCode");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            //golden errow
            formItems = new Dictionary<string, object>();
            itemName = "VtExpAccLB"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            formItems.Add("Left", left_e - 20);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);
            formItems.Add("LinkTo", "VatExpAcct");
            formItems.Add("LinkedObjectType", objectType);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 5;

            formItems = new Dictionary<string, object>();
            itemName = "PrBaseS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", 14);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("TaxableObject"));
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);
            formItems.Add("LinkTo", "PrBaseE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            objectType = "UDO_F_BDO_PTBS_D";
            uniqueID_CFL = "CFL_ProfitBase";
            FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_CFL);

            formItems = new Dictionary<string, object>();
            itemName = "PrBaseE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OIGE");
            formItems.Add("Alias", "U_prBase");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", 30);
            formItems.Add("Top", top);
            formItems.Add("Height", 14);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);
            formItems.Add("ChooseFromListUID", uniqueID_CFL);
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
            formItems.Add("TableName", "OIGE");
            formItems.Add("Alias", "U_PrBsDscr");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e + 32);
            formItems.Add("Width", width_e - 32);
            formItems.Add("Top", top);
            formItems.Add("Height", 14);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);
            formItems.Add("Enabled", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            //golden errow
            formItems = new Dictionary<string, object>();
            itemName = "PrBaseLB"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            formItems.Add("Left", left_e - 20);
            formItems.Add("Top", top);
            formItems.Add("Height", 14);
            formItems.Add("UID", itemName);
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);
            formItems.Add("LinkTo", "PrBaseE");
            formItems.Add("LinkedObjectType", objectType);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 5;

            formItems = new Dictionary<string, object>();
            itemName = "FillAmtTxs";
            formItems.Add("Caption", BDOSResources.getTranslate("Fill"));
            formItems.Add("Size", 20);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            formItems.Add("Left", left_s);
            formItems.Add("Width", 100);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            ////////////////////////////////////////////////
            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "VatCodS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("VatGroup"));
            formItems.Add("LinkTo", "VatCodE");
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string query = "select * " +
            "FROM  \"OVTG\" " +
            "WHERE \"Category\"='O'";

            oRecordSet.DoQuery(query);

            Dictionary<string, string> listValidValuesDict = new Dictionary<string, string>();
            while (!oRecordSet.EoF)
            {
                listValidValuesDict.Add(oRecordSet.Fields.Item("Code").Value, oRecordSet.Fields.Item("Name").Value);
                oRecordSet.MoveNext();
            }

            formItems = new Dictionary<string, object>();
            itemName = "VatCodE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OIGE");
            formItems.Add("Alias", "U_BDOSVatCod");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("ValidValues", listValidValuesDict);
            formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_ValueDescription);
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            ////////////////////////////////////////////////
            top = top + height + 1;

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
            formItems.Add("TableName", "OIGE");
            formItems.Add("Alias", "U_BDOSPrjCod");
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

                formItems = new Dictionary<string, object>();
                itemName = "DistrRul" + i + "S"; //10 characters
                formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                formItems.Add("Left", left_s);
                formItems.Add("Width", width_s);
                formItems.Add("Top", top);
                formItems.Add("Height", height);
                formItems.Add("UID", itemName);
                formItems.Add("Caption", activeDimensionsList[i.ToString()]);
                formItems.Add("LinkTo", "DistrRul" + i + "E");
                //formItems.Add("Visible", false);
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
                formItems.Add("TableName", "OIGE");
                formItems.Add("Alias", "U_BDOSDstRl" + i);
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
                //formItems.Add("Visible", false);
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


            //მარჯვენა
            top = 130;
            left_s = 300;
            width_s = 100;
            left_e = 420;
            width_e = 80;

            formItems = new Dictionary<string, object>();
            itemName = "AmtVatS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("VatAmount"));
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);
            formItems.Add("LinkTo", "AmtVatE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "AmtVatE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OIGE");
            formItems.Add("Alias", "U_amtVat");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", false);
            formItems.Add("RightJustified", true);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 5;

            formItems = new Dictionary<string, object>();
            itemName = "AmtPrTxS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("ProfitTaxAmount"));
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);
            formItems.Add("LinkTo", "AmtPrTxE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "AmtPrTxE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OIGE");
            formItems.Add("Alias", "U_amtPrTx");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", false);
            formItems.Add("RightJustified", true);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            //<-------------------------------------------სასაქონლო ზედნადები----------------------------------->
            height = oForm.Items.Item("22").Height;
            top = oForm.Items.Item("22").Top + height * 1.5 + 1;
            left_s = oForm.Items.Item("22").Left;
            left_e = oForm.Items.Item("21").Left;
            width_e = oForm.Items.Item("21").Width;

            string caption = BDOSResources.getTranslate("CreateWaybill");
            formItems = new Dictionary<string, object>();
            itemName = "BDO_WblTxt"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_e * 1.5);
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
            objectType = "UDO_F_BDO_WBLD_D"; //Waybill document
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
            formItems.Add("Left", left_e + width_e - 40);
            formItems.Add("Width", 40);
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
            formItems.Add("Left", left_e + width_e - 40 - 20);
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

            top = top + 2 * height + 5;

            formItems = new Dictionary<string, object>();
            itemName = "StckByDate";
            formItems.Add("Caption", BDOSResources.getTranslate("ShowStockByDate"));
            formItems.Add("Size", 20);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            formItems.Add("Left", left_e - 70);
            formItems.Add("Width", 150);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            oForm.DataSources.UserDataSources.Add("BDO_WblID", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
            oForm.DataSources.UserDataSources.Add("BDO_WblNum", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
            oForm.DataSources.UserDataSources.Add("BDO_WblSts", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
            //<-------------------------------------------სასაქონლო ზედნადები-----------------------------------

            GC.Collect();
        }

        public static void chooseFromList(SAPbouiCOM.Form oForm, SAPbouiCOM.IChooseFromListEvent oCFLEvento, string itemUID, bool beforeAction, out string errorText)
        {
            errorText = null;

            try
            {
                string sCFL_ID = oCFLEvento.ChooseFromListUID;

                if (beforeAction == false)
                {
                    SAPbouiCOM.DataTable oDataTable = null;
                    oDataTable = oCFLEvento.SelectedObjects;

                    if (oDataTable != null)
                    {
                        if (sCFL_ID == "Acct_CFL")
                        {
                            string account = Convert.ToString(oDataTable.GetValue("AcctCode", 0));

                            try
                            {
                                SAPbouiCOM.EditText oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("VatExpAcct").Specific;
                                oEditText.Value = account;
                            }
                            catch { }
                        }


                        else if (sCFL_ID == "Project_CFL")
                        {
                            string prjCode = Convert.ToString(oDataTable.GetValue("PrjCode", 0));

                            try
                            {
                                SAPbouiCOM.EditText oEdit = oForm.Items.Item("ProjectE").Specific;
                                oEdit.Value = prjCode;
                            }
                            catch { }
                        }
                        else if (sCFL_ID.Length >= 2 && sCFL_ID.Substring(0, sCFL_ID.Length - 2) == "Rule_CFL")
                        {
                            SAPbouiCOM.DataTable oDataTableSelectedObjects = oCFLEvento.SelectedObjects;
                            string val = oDataTableSelectedObjects.GetValue("OcrCode", 0);

                            SAPbouiCOM.EditText oEdit = oForm.Items.Item("DistrRul" + sCFL_ID.Substring(sCFL_ID.Length - 2, 1) + "E").Specific;
                            oEdit.Value = val;
                        }


                        if (sCFL_ID == "CFL_ProfitBase")
                        {
                            string ProfitBaseCode = Convert.ToString(oDataTable.GetValue("Code", 0));
                            string ProfitBaseName = Convert.ToString(oDataTable.GetValue("Name", 0));

                            try
                            {
                                SAPbouiCOM.EditText oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("PrBaseE").Specific;
                                oEditText.Value = ProfitBaseCode;
                            }
                            catch { }

                            try
                            {
                                SAPbouiCOM.EditText oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("PrBsDscr").Specific;
                                oEditText.Value = ProfitBaseName;
                            }
                            catch { }
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

        public static void formDataLoad(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            SAPbouiCOM.StaticText oStaticText = null;
            oForm.Freeze(true);

            try
            {
                //fillAmountTaxes(oForm, out errorText);

                setVisibleFormItems(oForm, out errorText);

                int docEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("OIGE").GetValue("DocEntry", 0));

                //-------------------------------------------სასაქონლო ზედნადები----------------------------------->
                string caption = BDOSResources.getTranslate("CreateWaybill");
                int wblDocEntry = 0;
                string wblID = "";
                string wblNum = "";
                string wblSts = "";

                if (docEntry != 0)
                {
                    Dictionary<string, string> wblDocInfo = BDO_Waybills.getWaybillDocumentInfo(docEntry, "60", out errorText);
                    wblDocEntry = Convert.ToInt32(wblDocInfo["DocEntry"]);
                    wblID = wblDocInfo["wblID"];
                    wblNum = wblDocInfo["number"];
                    wblSts = wblDocInfo["status"];

                    if (wblDocEntry != 0)
                    {
                        caption = BDOSResources.getTranslate("Wb") + ": " + wblSts + " " + wblID + (wblNum != "" ? " № " + wblNum : "");
                    }
                }
                else
                {
                    caption = BDOSResources.getTranslate("CreateWaybill");
                    wblDocEntry = 0;
                }

                oForm.DataSources.UserDataSources.Item("BDO_WblDoc").ValueEx = wblDocEntry == 0 ? "" : wblDocEntry.ToString();
                oForm.DataSources.UserDataSources.Item("BDO_WblID").ValueEx = wblID;
                oForm.DataSources.UserDataSources.Item("BDO_WblNum").ValueEx = wblNum;
                oForm.DataSources.UserDataSources.Item("BDO_WblSts").ValueEx = wblSts;
                oStaticText = (SAPbouiCOM.StaticText)oForm.Items.Item("BDO_WblTxt").Specific;
                oStaticText.Caption = caption;
                //<-------------------------------------------სასაქონლო ზედნადები-----------------------------------
            }
            catch (Exception ex)
            {
                oForm.DataSources.UserDataSources.Item("BDO_WblDoc").ValueEx = "";
                oForm.DataSources.UserDataSources.Item("BDO_WblID").ValueEx = "";
                oForm.DataSources.UserDataSources.Item("BDO_WblNum").ValueEx = "";
                oForm.DataSources.UserDataSources.Item("BDO_WblSts").ValueEx = "";
                oStaticText = (SAPbouiCOM.StaticText)oForm.Items.Item("BDO_WblTxt").Specific;
                oStaticText.Caption = BDOSResources.getTranslate("CreateWaybill");

                int errCode;
                string errMsg;

                Program.oCompany.GetLastError(out errCode, out errMsg);
                errorText = BDOSResources.getTranslate("ErrorDescription") + " : " + errMsg + "! " + BDOSResources.getTranslate("Code") + " : " + errCode + "! " + BDOSResources.getTranslate("OtherInfo") + " : " + ex.Message;
            }
            finally
            {
                oForm.Freeze(false);
                GC.Collect();
            }
        }

        public static void setVisibleFormItems(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            try
            {
                oForm.Items.Item("BDO_WblDoc").Enabled = false;

                string docEntry = oForm.DataSources.DBDataSources.Item("OIGE").GetValue("DocEntry", 0).Trim();
                bool docEntryIsEmpty = string.IsNullOrEmpty(docEntry);
                oForm.Items.Item("21").Click(SAPbouiCOM.BoCellClickType.ct_Regular); //მისაწვდომობის შეზღუდვისთვის

                oForm.Items.Item("liableVat").Enabled = (docEntryIsEmpty == true);
                bool LiableVat = oForm.DataSources.DBDataSources.Item("OIGE").GetValue("U_liableVat", 0) == "Y";
                oForm.Items.Item("VatExpAcct").Enabled = (LiableVat && docEntryIsEmpty == true);

                if (ProfitTaxTypeIsSharing)
                {
                    oForm.Items.Item("liablePrTx").Enabled = (docEntryIsEmpty == true);
                    bool LiablePrTx = oForm.DataSources.DBDataSources.Item("OIGE").GetValue("U_liablePrTx", 0) == "Y";
                    oForm.Items.Item("PrBaseE").Enabled = (LiablePrTx && docEntryIsEmpty == true);
                    string uniqueID_lf_ProfitBaseCFL = "CFL_ProfitBase";
                    oForm.Items.Item("PrBaseE").Specific.ChooseFromListUID = uniqueID_lf_ProfitBaseCFL;
                    oForm.Items.Item("PrBaseE").Specific.ChooseFromListAlias = "Code";
                    oForm.Items.Item("AmtPrTxE").Enabled = false;
                    oForm.Items.Item("AmtVatE").Enabled = false;
                }
                else
                {
                    oForm.Items.Item("liablePrTx").Visible = false;
                    oForm.Items.Item("PrBaseS").Visible = false;
                    oForm.Items.Item("PrBaseE").Visible = false;
                    oForm.Items.Item("PrBsDscr").Visible = false;
                    oForm.Items.Item("PrBaseLB").Visible = false;
                    oForm.Items.Item("AmtPrTxS").Visible = false;
                    oForm.Items.Item("AmtPrTxE").Visible = false;
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

        public static void itemPressed(SAPbouiCOM.Form oForm, SAPbouiCOM.ItemEvent pVal, out int newDocEntry, out string bstrUDOObjectType, out string errorText)
        {
            errorText = null;
            newDocEntry = 0;
            bstrUDOObjectType = null;

            int docEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("OIGE").GetValue("DocEntry", 0));
            string cancelled = oForm.DataSources.DBDataSources.Item("OIGE").GetValue("CANCELED", 0).Trim();
            string docType = oForm.DataSources.DBDataSources.Item("OIGE").GetValue("DocType", 0).Trim();
            string objectType = "60";

            if (pVal.ItemUID == "BDO_WblTxt")
            {
                string wblDoc = oForm.DataSources.UserDataSources.Item("BDO_WblDoc").ValueEx;
                bstrUDOObjectType = "UDO_F_BDO_WBLD_D";

                if (docEntry != 0 & (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_VIEW_MODE))
                {
                    if (wblDoc == "" && cancelled == "N" && docType == "I")
                    {
                        BDO_Waybills.createDocument(objectType, docEntry, null, null, null, null, out newDocEntry, out errorText);
                        if (errorText == null & newDocEntry != 0)
                        {
                            oForm.DataSources.UserDataSources.Item("BDO_WblDoc").ValueEx = newDocEntry.ToString();
                            //errorText = "წარმატებით შეიქმნა ზედნადების დოკუმენტი! DocEntry : " + newDocEntry;
                            formDataLoad(oForm, out errorText);
                            return;
                        }
                    }
                    else if (cancelled != "N")
                    {
                        errorText = BDOSResources.getTranslate("DocumentMustNotBeCancelledOrCancellation");
                    }
                    else if (docType != "I")
                    {
                        errorText = BDOSResources.getTranslate("DocumentTypeMustBeItem");
                    }
                }
                else
                {
                    errorText = BDOSResources.getTranslate("ToCreateWaybillWriteDocument");
                }
            }
        }

        public static void uiApp_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(BusinessObjectInfo.FormTypeEx, Program.currentFormCount);

            if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD)
            {
                if (BusinessObjectInfo.BeforeAction == true)
                {
                    SAPbouiCOM.DBDataSource DocDBSource = oForm.DataSources.DBDataSources.Item(0);
                    if (DocDBSource.GetValue("CANCELED", 0) == "N")
                    {
                        //უარყოფითი ნაშთების კონტროლი დოკ.თარიღით
                        bool rejection = false;
                        CommonFunctions.blockNegativeStockByDocDate(oForm, "OIGE", "IGE1", "WhsCode", out rejection);
                        if (rejection)
                        {
                            Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("DocumentCannotBeAdded"));
                            BubbleEvent = false;
                        }

                        //ძირითადი საშუალებების შემოწმება
                        if (BatchNumberSelection.SelectedBatches != null)
                        {
                            bool rejectionAsset = false;
                            CommonFunctions.blockAssetInvoice(oForm, "OIGE", out rejectionAsset);
                            if (rejectionAsset)
                            {
                                Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("DocumentCannotBeAdded"));
                                BubbleEvent = false;
                            }
                        }

                        if (ProfitTaxTypeIsSharing == true)
                        {
                            // მოგების გადასახადი
                            if (DocDBSource.GetValue("U_liablePrTx", 0) == "Y")
                            {
                                bool TaxAccountsIsEmpty = ProfitTax.TaxAccountsIsEmpty();
                                if (TaxAccountsIsEmpty == true || oForm.DataSources.DBDataSources.Item("OIGE").GetValue("U_prBase", 0) == "")
                                {
                                    Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("TaxAccounts") + ", " + BDOSResources.getTranslate("TaxableObject") + " " + BDOSResources.getTranslate("YouCantLeaveEmpty"));
                                    Program.uiApp.MessageBox(BDOSResources.getTranslate("OperationUnsuccesfullSeeLog"));
                                    BubbleEvent = false;
                                }
                            }
                            if (DocDBSource.GetValue("U_liableVat", 0) == "Y")
                            {
                                if (DocDBSource.GetValue("U_VatExpAcct", 0) == "")
                                {
                                    Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("VatExpense") + " " + BDOSResources.getTranslate("Account") + " " + BDOSResources.getTranslate("YouCantLeaveEmpty"));
                                    Program.uiApp.MessageBox(BDOSResources.getTranslate("OperationUnsuccesfullSeeLog"));
                                    BubbleEvent = false;
                                }
                                if (DocDBSource.GetValue("U_BDOSVatCod", 0) == "")
                                {
                                    Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("VatGroup") + " " + BDOSResources.getTranslate("YouCantLeaveEmpty"));
                                    Program.uiApp.MessageBox(BDOSResources.getTranslate("OperationUnsuccesfullSeeLog"));
                                    BubbleEvent = false;
                                }
                            }
                        }
                    }
                }

                if (BusinessObjectInfo.ActionSuccess != BusinessObjectInfo.BeforeAction && BubbleEvent)
                {
                    //დოკუმენტის გატარების დროს გატარდეს ბუღლტრული გატარება
                    SAPbouiCOM.DBDataSource DocDBSourcePAYR = oForm.DataSources.DBDataSources.Item(0);

                    if (DocDBSourcePAYR.GetValue("CANCELED", 0) == "N")
                    {
                        string DocEntry = DocDBSourcePAYR.GetValue("DocEntry", 0);
                        DocEntry = DocEntry == "" ? "0" : DocEntry;

                        string DocCurrency = DocDBSourcePAYR.GetValue("DocCur", 0);
                        decimal DocRate = FormsB1.cleanStringOfNonDigits(DocDBSourcePAYR.GetValue("DocRate", 0));
                        string DocNum = DocDBSourcePAYR.GetValue("DocNum", 0);
                        DateTime DocDate = DateTime.ParseExact(DocDBSourcePAYR.GetValue("DocDate", 0), "yyyyMMdd", CultureInfo.InvariantCulture);

                        CommonFunctions.StartTransaction();

                        Program.JrnLinesGlobal = new DataTable();
                        DataTable reLines = null;
                        DataTable JrnLinesDT = createAdditionalEntries(oForm, null, null, DocCurrency, DocEntry, out reLines, DocRate);

                        JrnEntry(DocEntry, DocNum, DocDate, JrnLinesDT, reLines, out errorText);
                        if (errorText != null)
                        {
                            Program.uiApp.MessageBox(errorText);
                            BubbleEvent = false;
                        }
                        else
                        {
                            if (BusinessObjectInfo.ActionSuccess == false)
                            {
                                Program.JrnLinesGlobal = JrnLinesDT;
                            }
                        }

                        if (Program.oCompany.InTransaction)
                        {
                            //თუ დოკუმენტი გატარდა, მერე ვაკეთებს ბუღალტრულ გატარებას
                            if (BusinessObjectInfo.ActionSuccess & BusinessObjectInfo.BeforeAction == false)
                            {
                                CommonFunctions.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                                BatchNumberSelection.SelectedBatches = null;
                            }
                            else
                            {
                                CommonFunctions.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
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

            if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE & BusinessObjectInfo.BeforeAction == false & BusinessObjectInfo.ActionSuccess == true)
            {
                if (Program.cancellationTrans == true & Program.canceledDocEntry != 0)
                {
                    cancellation(oForm, Program.canceledDocEntry, out errorText);
                    Program.canceledDocEntry = 0;
                }
            }

            if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD & BusinessObjectInfo.BeforeAction == false)
            {
                formDataLoad(oForm, out errorText);
            }
        }

        public static void cancellation(SAPbouiCOM.Form oForm, int docEntry, out string errorText)
        {
            errorText = null;

            try
            {
                JournalEntry.cancellation(oForm, docEntry, "60", out errorText);
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

                if (pVal.ItemUID == "Taxes" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED & pVal.BeforeAction == true)
                {
                    oForm.PaneLevel = 3;
                    setVisibleFormItems(oForm, out errorText);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
                {
                    if ((pVal.ItemUID == "liablePrTx" || pVal.ItemUID == "liableVat") && pVal.BeforeAction == false)
                    {
                        oForm.Freeze(true);
                        LiableTaxes_OnClick(oForm, out errorText);
                        oForm.Freeze(false);
                    }
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                {
                    SAPbouiCOM.ChooseFromListEvent oCFLEvento = null;
                    oCFLEvento = ((SAPbouiCOM.ChooseFromListEvent)(pVal));

                    string sCFL_ID = oCFLEvento.ChooseFromListUID;

                    if (pVal.BeforeAction == false)
                    {
                        chooseFromList(oForm, oCFLEvento, pVal.ItemUID, pVal.BeforeAction, out errorText);
                    }
                    else if (sCFL_ID.Length > 1 && sCFL_ID.Substring(0, sCFL_ID.Length - 2) == "Rule_CFL")
                    {
                        oForm.Freeze(true);
                        string dimensionCode = sCFL_ID.Substring(sCFL_ID.Length - 2, 1);

                        SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(sCFL_ID);
                        SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();

                        string startDateStr = oForm.Items.Item("38").Specific.Value;
                        DateTime DocDate = DateTime.TryParseExact(startDateStr, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DocDate) == false ? DateTime.Now : DocDate;

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

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD && pVal.BeforeAction)
                {
                    createFormItems(oForm, out errorText);
                    formDataLoad(oForm, out errorText);

                    changeFormItems(oForm, out errorText);
                }

                if (pVal.ItemUID == "FillAmtTxs" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && !pVal.BeforeAction && !pVal.InnerEvent)
                {
                    try
                    {
                        fillAmountTaxes(oForm);
                    }
                    catch (Exception ex)
                    {
                        Program.uiApp.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    }
                }

                if ((pVal.ItemUID == "BDO_WblTxt") & pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED & pVal.BeforeAction == false)
                {
                    oForm.Freeze(true);
                    int newDocEntry = 0;
                    string bstrUDOObjectType = null;

                    itemPressed(oForm, pVal, out newDocEntry, out bstrUDOObjectType, out errorText);

                    if (errorText != null)
                    {
                        Program.uiApp.MessageBox(errorText);
                    }

                    oForm.Freeze(false);
                    oForm.Update();

                    if (newDocEntry != 0 && bstrUDOObjectType != null)
                    {
                        Program.uiApp.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_UserDefinedObject, bstrUDOObjectType, newDocEntry.ToString());
                    }
                }

                //არ წაშალოთ!!! (დაგვჭირდება ოდესმე)
                //if (pVal.ItemUID == "13" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_GOT_FOCUS)
                //{
                //    if (!pVal.BeforeAction)
                //    {
                //        if (pVal.ColUID == "1") //Item No.
                //        {
                //            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item(pVal.ItemUID).Specific));
                //            itemCodeOld = oMatrix.GetCellSpecific(pVal.ColUID, pVal.Row).Value;
                //        }
                //        else if (pVal.ColUID == "15") //Whse
                //        {
                //            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item(pVal.ItemUID).Specific));
                //            warehouseOld = oMatrix.GetCellSpecific(pVal.ColUID, pVal.Row).Value;
                //        }
                //    }
                //}

                //if (pVal.ItemUID == "13" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_LOST_FOCUS)
                //{
                //    if (!pVal.BeforeAction)
                //    {
                //        if (pVal.ColUID == "1" || pVal.ColUID == "15") //Item No. || //Whse
                //        {
                //            updateInStockByWarehouseAndDate(oForm, pVal.Row);
                //        }
                //    }
                //}

                if (pVal.ItemUID == "StckByDate" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && !pVal.BeforeAction && !pVal.InnerEvent)
                {
                    updateInStockByWarehouseAndDate(oForm);
                }

                if (pVal.ItemUID == "9" && pVal.ItemChanged && !pVal.BeforeAction)
                {
                    updateInStockByWarehouseAndDate(oForm);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_RESIZE && pVal.BeforeAction == false)
                {
                    oForm.Freeze(true);
                    resizeForm(oForm, out errorText);
                    oForm.Freeze(false);
                }
            }
        }

        public static void changeFormItems(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("13").Specific));

            SAPbouiCOM.Columns oColumns = oMatrix.Columns;
            SAPbouiCOM.Column oColumn = oColumns.Item("U_BDOSInStck");
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("InStockByDate");
            oColumn.Editable = false;
        }

        private static void updateInStockByWarehouseAndDate(SAPbouiCOM.Form oForm, int rowIndex = 0)
        {
            string docDate = oForm.DataSources.DBDataSources.Item("OIGE").GetValue("DocDate", 0);         
            try
            {
                oForm.Freeze(true);

                SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("13").Specific));             
                int rowCount = rowIndex == 0 ? oMatrix.RowCount : rowIndex;
                int i = rowIndex == 0 ? 1 : rowIndex;

                for (; i <= rowCount; i++)
                {
                    string warehouse = oMatrix.GetCellSpecific("15", i).Value;
                    string itemCode = oMatrix.GetCellSpecific("1", i).Value;

                    if ((itemCode != itemCodeOld || warehouse != warehouseOld) || rowIndex == 0)
                    {
                        if (!string.IsNullOrEmpty(itemCode) && !string.IsNullOrEmpty(warehouse) && !string.IsNullOrEmpty(docDate))
                        {
                            SAPbouiCOM.Column oColumn = oMatrix.Columns.Item("U_BDOSInStck");
                            oColumn.Editable = true;
                            oMatrix.Columns.Item("U_BDOSInStck").Cells.Item(i).Specific.Value = FormsB1.ConvertDecimalToStringForEditboxStrings(CommonFunctions.getInStockByWarehouseAndDate(itemCode, warehouse, docDate));
                            oMatrix.Columns.Item("2").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            oColumn.Editable = false;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Program.uiApp.StatusBar.SetSystemMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
                itemCodeOld = null;
                warehouseOld = null;
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

            oItem = oForm.Items.Item("Taxes");
            int top = oItem.Top + 30;
            int height = 14;

            oItem = oForm.Items.Item("VtExpAccE");
            oItem.Top = top;
            oItem = oForm.Items.Item("VatExpAcct");
            oItem.Top = top;
            oItem = oForm.Items.Item("VtExpAccLB");
            oItem.Top = top;
            oItem = oForm.Items.Item("AmtVatS");
            oItem.Top = top;
            oItem = oForm.Items.Item("AmtVatE");
            oItem.Top = top;

            top = top + height + 5;
            oItem = oForm.Items.Item("PrBaseS");
            oItem.Top = top;
            oItem = oForm.Items.Item("PrBaseE");
            oItem.Top = top;
            oItem = oForm.Items.Item("PrBsDscr");
            oItem.Top = top;
            oItem = oForm.Items.Item("PrBaseLB");
            oItem.Top = top;
            oItem = oForm.Items.Item("AmtPrTxS");
            oItem.Top = top;
            oItem = oForm.Items.Item("AmtPrTxE");
            oItem.Top = top;

            top = top + height + 5;
            oItem = oForm.Items.Item("FillAmtTxs");
            oItem.Top = top;

            top = top + height + 5;
            oItem = oForm.Items.Item("VatCodS");
            oItem.Top = top;
            oItem = oForm.Items.Item("VatCodE");
            oItem.Top = top;


            top = top + height + 5;
            oItem = oForm.Items.Item("ProjectS");
            oItem.Top = top;
            oItem = oForm.Items.Item("ProjectE");
            oItem.Top = top;

            for (int i = 1; i <= 5; i++)
            {
                top = top + height + 1;
                try
                {
                    oItem = oForm.Items.Item("DistrRul" + i + "S");
                    oItem.Top = top;
                    oItem = oForm.Items.Item("DistrRul" + i + "E");
                    oItem.Top = top;
                }
                catch { }
            }
        }

        public static void fillAmountTaxes(SAPbouiCOM.Form oForm)
        {          
            try
            {
                oForm.Freeze(true);

                string docDateStr = oForm.DataSources.DBDataSources.Item("OIGE").GetValue("DocDate", 0);
                if (string.IsNullOrEmpty(docDateStr))
                {
                    string errorText = BDOSResources.getTranslate("PostingDate") + " " + BDOSResources.getTranslate("YouCantLeaveEmpty");
                    throw new Exception(errorText);
                }
                DateTime docDate = DateTime.ParseExact(docDateStr, "yyyyMMdd", CultureInfo.InvariantCulture);

                NumberFormatInfo Nfi = new NumberFormatInfo() { NumberDecimalSeparator = "." };

                bool liablePrTx = false;
                bool liableVat = false;
                decimal amtPrTx = 0;
                decimal amtVat = 0;

                if (((SAPbouiCOM.CheckBox)(oForm.Items.Item("liablePrTx").Specific)).Checked)
                    liablePrTx = true;

                if (((SAPbouiCOM.CheckBox)(oForm.Items.Item("liableVat").Specific)).Checked)
                    liableVat = true;

                if (liablePrTx || liableVat)
                {
                    string ListNum = oForm.DataSources.DBDataSources.Item("OIGE").GetValue("GroupNum", 0);

                    SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        
                    string isGrossPrc = "N";
                    string query = @"SELECT ""IsGrossPrc"" FROM ""OPLN"" WHERE ""OPLN"".""ListNum""='" + ListNum.Replace("'", "''") + "'";
                    oRecordSet.DoQuery(query);
                    if (!oRecordSet.EoF)
                    {
                        isGrossPrc = oRecordSet.Fields.Item("IsGrossPrc").Value;
                    }

                    decimal profitTaxRate = ProfitTax.GetProfitTaxRate();

                    SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("13").Specific));
                    for (int i = 1; i < oMatrix.RowCount + 1; i++)
                    {
                        decimal quantity = FormsB1.cleanStringOfNonDigits(oMatrix.Columns.Item("9").Cells.Item(i).Specific.Value);
                        decimal price = FormsB1.cleanStringOfNonDigits(oMatrix.Columns.Item("10").Cells.Item(i).Specific.Value);
                        string itemCode = oMatrix.Columns.Item("1").Cells.Item(i).Specific.Value.ToString();

                        decimal amt = quantity * price;
                        decimal vatRate = CommonFunctions.GetVatGroupRate("", itemCode);
                        decimal currAmtVat;

                        if (isGrossPrc == "Y")
                            currAmtVat = (amt * 100) / (100 + vatRate) * vatRate / 100;
                        else
                            currAmtVat = amt * vatRate / 100;

                        if (liableVat)
                            amtVat = amtVat + currAmtVat;

                        if (liablePrTx)
                        {
                            if (isGrossPrc != "Y" && docDate < new DateTime(2019, 7, 5))
                            {
                                amt = amt + currAmtVat;
                            }
                            amtPrTx = amtPrTx + amt / (100 - profitTaxRate) * profitTaxRate;
                        }
                    }
                }

                SAPbouiCOM.Item oItemAmtPrTx = oForm.Items.Item("AmtPrTxE");
                SAPbouiCOM.Item oItemAmtVat = oForm.Items.Item("AmtVatE");
                oItemAmtPrTx.Enabled = true;
                oItemAmtVat.Enabled = true;

                SAPbouiCOM.EditText oEditAmtPrTx = ((SAPbouiCOM.EditText)(oItemAmtPrTx.Specific));
                oEditAmtPrTx.Value = FormsB1.ConvertDecimalToString(CommonFunctions.roundAmountByGeneralSettings(amtPrTx, "Sum"));

                SAPbouiCOM.EditText oEditAmtVat = ((SAPbouiCOM.EditText)(oItemAmtVat.Specific));
                oEditAmtVat.Value = FormsB1.ConvertDecimalToString(CommonFunctions.roundAmountByGeneralSettings(amtVat, "Sum"));
            }
            catch(Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                oForm.Items.Item("21").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                oForm.Items.Item("AmtPrTxE").Enabled = false;
                oForm.Items.Item("AmtVatE").Enabled = false;
                oForm.Freeze(false);
            }
        }

        public static DataTable createAdditionalEntries(SAPbouiCOM.Form oForm, SAPbobsCOM.GeneralData oGeneralData, DataTable DTSource, string DocCurrency, string DocEntry, out DataTable reLines, decimal DocRate)
        {
            DataTable jeLines = JournalEntry.JournalEntryTable();
            DataRow jeLinesRow = null;

            DataTable AccountTable = CommonFunctions.GetOACTTable();

            reLines = ProfitTax.ProfitTaxTable();
            DataRow reLinesRow = null;

            SAPbouiCOM.DBDataSource DBDataSourceTable = null;
            int JEcount = 0;

            if (oForm == null)
            {
                JEcount = DTSource.Rows.Count;
            }
            else
            {
                DBDataSourceTable = oForm.DataSources.DBDataSources.Item("IGE1");
                JEcount = DBDataSourceTable.Size;
            }

            string ListNum = oForm.DataSources.DBDataSources.Item("OIGE").GetValue("GroupNum", 0);
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string IsGrossPrc = "N";
            string query = @"SELECT ""IsGrossPrc"" FROM ""OPLN"" WHERE ""OPLN"".""ListNum""='" + ListNum.Replace("'", "''") + "'";
            oRecordSet.DoQuery(query);
            if (!oRecordSet.EoF)
            {
                IsGrossPrc = oRecordSet.Fields.Item("IsGrossPrc").Value;
            }

            bool liablePrTx = false;
            bool liableVat = false;

            //მოგების გადასახადის გატარება
            ProfitTaxTypeIsSharing = ProfitTax.ProfitTaxTypeIsSharing();

            if (ProfitTaxTypeIsSharing)
            {
                if (((SAPbouiCOM.CheckBox)(oForm.Items.Item("liablePrTx").Specific)).Checked)
                    liablePrTx = true;

                if (((SAPbouiCOM.CheckBox)(oForm.Items.Item("liableVat").Specific)).Checked)
                    liableVat = true;

                decimal U_BDO_PrTxRt = Convert.ToDecimal(CommonFunctions.getOADM("U_BDO_PrTxRt").ToString());

                if (liablePrTx)
                {
                    string prBase = oForm.DataSources.DBDataSources.Item("OIGE").GetValue("U_prBase", 0).Trim();
                    string DebitAccount = CommonFunctions.getOADM("U_BDO_CapAcc").ToString();
                    string CreditAccount = CommonFunctions.getOADM("U_BDO_TaxAcc").ToString();

                    for (int i = 0; i < JEcount; i++)
                    {
                        //decimal Price = FormsB1.cleanStringOfNonDigits(CommonFunctions.getChildOrDbDataSourceValue(DBDataSourceTable, null, DTSource, "Price", i).ToString());
                        //decimal Quantity = FormsB1.cleanStringOfNonDigits(CommonFunctions.getChildOrDbDataSourceValue(DBDataSourceTable, null, DTSource, "Quantity", i).ToString());
                        //string ItemCode = CommonFunctions.getChildOrDbDataSourceValue(DBDataSourceTable, null, DTSource, "ItemCode", i).ToString();
                        //decimal BaseGross = Price * Quantity;

                        //decimal VatRate = CommonFunctions.GetVatGroupRate("", ItemCode);
                        //if (IsGrossPrc != "Y")
                        //{
                        //    BaseGross = BaseGross + BaseGross * VatRate / 100;
                        //}

                        //decimal TaxAmount = BaseGross * U_BDO_PrTxRt / (100 - U_BDO_PrTxRt);
                        DocCurrency = DocCurrency == CommonFunctions.getLocalCurrency() ? "" : DocCurrency;
                        //decimal TaxAmountFC = DocCurrency == "" ? 0 : TaxAmount / DocRate;

                        var amtPrTx = oForm.DataSources.DBDataSources.Item("OIGE").GetValue("U_amtPrTx", 0);
                        var amtPrTxFC = DocCurrency == "" ? 0 : Convert.ToDecimal(amtPrTx) / DocRate;
                        var amtVat = oForm.DataSources.DBDataSources.Item("OIGE").GetValue("U_amtVat", 0);

                        jeLinesRow = jeLines.Rows.Add();
                        jeLinesRow["AccountCode"] = CreditAccount; //Credit
                        jeLinesRow["ShortName"] = CreditAccount;
                        jeLinesRow["ContraAccount"] = DebitAccount;
                        jeLinesRow["Credit"] = Convert.ToDouble(amtPrTx, CultureInfo.InvariantCulture);
                        jeLinesRow["FCCredit"] = Convert.ToDouble(amtPrTxFC, CultureInfo.InvariantCulture);
                        jeLinesRow["Debit"] = 0;
                        jeLinesRow["FCCurrency"] = DocCurrency;

                        jeLinesRow = jeLines.Rows.Add();
                        jeLinesRow["AccountCode"] = DebitAccount;
                        jeLinesRow["ShortName"] = DebitAccount;
                        jeLinesRow["ContraAccount"] = CreditAccount;
                        jeLinesRow["Credit"] = 0;
                        jeLinesRow["Debit"] = Convert.ToDouble(amtPrTx, CultureInfo.InvariantCulture);
                        jeLinesRow["FCDebit"] = Convert.ToDouble(amtPrTxFC, CultureInfo.InvariantCulture);
                        jeLinesRow["FCCurrency"] = DocCurrency;

                        reLinesRow = reLines.Rows.Add();
                        reLinesRow["debitAccount"] = DebitAccount;
                        reLinesRow["creditAccount"] = CreditAccount;
                        reLinesRow["prBase"] = prBase;
                        reLinesRow["txType"] = "Accrual";
                        reLinesRow["amtTx"] = Convert.ToDouble(amtVat, CultureInfo.InvariantCulture);
                        reLinesRow["amtPrTx"] = Convert.ToDouble(amtPrTx, CultureInfo.InvariantCulture);
                    }
                }
            }

            if (liableVat)
            {
                string DebitAccount = oForm.DataSources.DBDataSources.Item("OIGE").GetValue("U_VatExpAcct", 0).Trim();
                string U_DistrRule1 = oForm.DataSources.DBDataSources.Item("OIGE").GetValue("U_BDOSDstRl1", 0).Trim();
                string U_DistrRule2 = oForm.DataSources.DBDataSources.Item("OIGE").GetValue("U_BDOSDstRl2", 0).Trim();
                string U_DistrRule3 = oForm.DataSources.DBDataSources.Item("OIGE").GetValue("U_BDOSDstRl3", 0).Trim();
                string U_DistrRule4 = oForm.DataSources.DBDataSources.Item("OIGE").GetValue("U_BDOSDstRl4", 0).Trim();
                string U_DistrRule5 = oForm.DataSources.DBDataSources.Item("OIGE").GetValue("U_BDOSDstRl5", 0).Trim();
                string U_PrjCode = oForm.DataSources.DBDataSources.Item("OIGE").GetValue("U_BDOSPrjCod", 0).Trim();
                string U_BDOSVatCod = oForm.DataSources.DBDataSources.Item("OIGE").GetValue("U_BDOSVatCod", 0).Trim();

                for (int i = 0; i < JEcount; i++)
                {
                    string ItemCode = CommonFunctions.getChildOrDbDataSourceValue(DBDataSourceTable, null, DTSource, "ItemCode", i).ToString();
                    decimal VatRate = CommonFunctions.GetVatGroupRate("", ItemCode);

                    SAPbobsCOM.VatGroups oVatCode;
                    oVatCode = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVatGroups);
                    oVatCode.GetByKey(U_BDOSVatCod);
                    string CreditAccount = oVatCode.TaxAccount;

                    decimal Price = FormsB1.cleanStringOfNonDigits(CommonFunctions.getChildOrDbDataSourceValue(DBDataSourceTable, null, DTSource, "Price", i).ToString());
                    decimal Quantity = FormsB1.cleanStringOfNonDigits(CommonFunctions.getChildOrDbDataSourceValue(DBDataSourceTable, null, DTSource, "Quantity", i).ToString());
                    decimal BaseGross = Price * Quantity;
                    decimal TaxAmount = 0;

                    if (IsGrossPrc == "N")
                    {
                        TaxAmount = BaseGross * VatRate / 100;
                    }
                    else
                    {
                        TaxAmount = BaseGross * VatRate / (100 + VatRate);
                    }

                    DocCurrency = DocCurrency == CommonFunctions.getLocalCurrency() ? "" : DocCurrency;
                    decimal TaxAmountFC = DocCurrency == "" ? 0 : TaxAmount / DocRate;

                    JournalEntry.AddJournalEntryRow(AccountTable, jeLines, "OnlyDebit", DebitAccount, CreditAccount, TaxAmount, TaxAmountFC, DocCurrency, "", "", "", "", "", "", "", "");
                    JournalEntry.AddJournalEntryRow(AccountTable, jeLines, "OnlyCredit", DebitAccount, CreditAccount, TaxAmount, TaxAmountFC, DocCurrency, U_DistrRule1, U_DistrRule2, U_DistrRule3, U_DistrRule4, U_DistrRule5, U_PrjCode, U_BDOSVatCod, "");
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
                          VatGroupCode = row.Field<string>("VatGroup")


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

                          row["Credit"] = g.Sum(r => r.Field<double>("Credit"));
                          row["Debit"] = g.Sum(r => r.Field<double>("Debit"));
                          row["FCCredit"] = g.Sum(r => r.Field<double>("FCCredit"));
                          row["FCDebit"] = g.Sum(r => r.Field<double>("FCDebit"));
                          return row;
                      }).CopyToDataTable();
            }

            if (reLines.Rows.Count > 0)
            {
                DataTable reLinesNew = reLines;

                reLinesNew = reLinesNew.AsEnumerable()
                      .GroupBy(row => new
                      {
                          debitAccount = row.Field<string>("debitAccount"),
                          creditAccount = row.Field<string>("creditAccount"),
                          prBase = row.Field<string>("prBase"),
                          txType = row.Field<string>("txType")

                      })

                      .Select(g =>
                      {
                          var row = reLinesNew.NewRow();
                          row["debitAccount"] = g.Key.debitAccount;
                          row["creditAccount"] = g.Key.creditAccount;
                          row["prBase"] = g.Key.prBase;
                          row["txType"] = g.Key.txType;


                          row["amtTx"] = g.Sum(r => r.Field<float>("amtTx"));
                          row["amtPrTx"] = g.Sum(r => r.Field<float>("amtPrTx"));

                          return row;
                      }).CopyToDataTable();

                reLines = reLinesNew;
            }
            return jeLines;
        }

        public static void JrnEntry(string DocEntry, string DocNum, DateTime DocDate, DataTable JrnLinesDT, DataTable reLines, out string errorText)
        {
            errorText = null;

            try
            {
                JournalEntry.JrnEntry(DocEntry, "60", "Goods Issue: " + DocNum, DocDate, JrnLinesDT, out errorText);

                if (errorText != null)
                {
                    return;
                }

                for (int i = 0; i < reLines.Rows.Count; i++)
                {
                    reLines.Rows[i]["DocEntry"] = DocEntry == "" ? 0 : Convert.ToInt32(DocEntry);
                    reLines.Rows[i]["DocNum"] = DocNum.ToString();
                    reLines.Rows[i]["docDate"] = DocDate;
                }

                ProfitTax.AddRecord(reLines, "60", "Goods Issue: " + DocNum, out errorText);
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
        }

        public static void LiableTaxes_OnClick(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            string liableVat = oForm.DataSources.DBDataSources.Item("OIGE").GetValue("U_liableVat", 0).Trim();
            string liablePrTx = oForm.DataSources.DBDataSources.Item("OIGE").GetValue("U_liablePrTx", 0).Trim();

            if (liableVat != "Y")
            {
                SAPbouiCOM.EditText oEdit = oForm.Items.Item("VatExpAcct").Specific;
                oEdit.Value = "";

                SAPbouiCOM.Item oItemAmtVat = oForm.Items.Item("AmtVatE");
                oItemAmtVat.Enabled = true;
                oForm.Items.Item("AmtVatE").Specific.Value = "";
                oForm.Items.Item("21").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                oForm.Items.Item("AmtVatE").Enabled = false;
            }

            if (liablePrTx == "Y")
            {
                SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string query = @"SELECT  
		                                 ""U_BDO_prBsGI"",
		                                 ""U_BDO_prGIDs"" 
                                FROM ""OADM""";

                oRecordSet.DoQuery(query);
                if (!oRecordSet.EoF)
                {
                    try
                    {
                        SAPbouiCOM.EditText oEdit = oForm.Items.Item("PrBaseE").Specific;
                        oEdit.Value = oRecordSet.Fields.Item("U_BDO_prBsGI").Value;
                    }
                    catch { }

                    try
                    {
                        SAPbouiCOM.EditText oEdit = oForm.Items.Item("PrBsDscr").Specific;
                        oEdit.Value = oRecordSet.Fields.Item("U_BDO_prGIDs").Value;
                    }
                    catch { }
                }
            }
            else
            {
                SAPbouiCOM.EditText oEdit = oForm.Items.Item("PrBaseE").Specific;
                oEdit.Value = "";

                oEdit = oForm.Items.Item("PrBsDscr").Specific;
                oEdit.Value = "";

                SAPbouiCOM.Item oItemAmtPrTx = oForm.Items.Item("AmtPrTxE");
                oItemAmtPrTx.Enabled = true;
                oForm.Items.Item("AmtPrTxE").Specific.Value = "";
                oForm.Items.Item("21").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                oForm.Items.Item("AmtPrTxE").Enabled = false;
            }

            setVisibleFormItems(oForm, out errorText);
        }
    }
}
