using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace BDO_Localisation_AddOn
{
    static partial class BusinessPartners
    {
        public static string GetCardCodeByTin(string tin, string cardType, out string cardName)
        {
            cardName = null;

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string query = $"SELECT \"CardCode\", \"CardName\", \"Currency\" FROM \"OCRD\" WHERE \"LicTradNum\" = '{tin}' ";

            if (string.IsNullOrEmpty(cardType) == false)
            {
                query = query + $" AND \"CardType\" = '{cardType}'";
            }

            oRecordSet.DoQuery(query);

            if (!oRecordSet.EoF)
            {
                cardName = oRecordSet.Fields.Item("CardName").Value;
                return oRecordSet.Fields.Item("CardCode").Value;
            }
            else
            {
                return null;
            }
        }

        public static void createUserFields( out string errorText)
        {
            errorText = null;
            Dictionary<string, object> fieldskeysMap;
            List<string> listValidValues;

            fieldskeysMap = new Dictionary<string, object>();
            listValidValues = new List<string>();
            listValidValues.Add(""); //-1
            listValidValues.Add("Individual Enterpreneur"); //1 //ინდივიდუალური მეწარმე
            listValidValues.Add("LTD"); //2 //შეზღუდული პასუხისმგებლობის საზოგადოება
            listValidValues.Add("Physical Entity"); //3 //ფიზიკური პირი
            listValidValues.Add("JLC"); //4 //სოლიდარული პასუხისმგებლობის საზოგადოება
            listValidValues.Add("JSC"); //5 //სააქციო საზოგადოება
            listValidValues.Add("Commandite"); //6 //კომანდიტური საზოგადოება
            listValidValues.Add("Cooperative"); //7 //კოოპერატივი
            listValidValues.Add("Non Commercial Legal Entity"); //8 //არაკომერციული იურიდიული პირი
            listValidValues.Add("Foreign Enterprise Branch"); //9 //უცხოური საწარმოს ფილიალი
            listValidValues.Add("Foreign Enterprise Physical Entity"); //10 //უცხოური საწარმო - ფიზიკური პირი
            listValidValues.Add("Foreign Enterprise Legal Entity"); //11 //უცხოური საწარმო - იურიდიული პირი
            listValidValues.Add("LEPL"); //12 //საჯარო სამართლის იურიდიული პირი
            listValidValues.Add("Government Body"); //13 //სახელმწიფო ორგანო
            listValidValues.Add("Partnership"); //14 //ამხანაგობა

            //UDO.addNewValidValuesUserFieldsMD( "ACRD", "BDO_TaxTyp", "14", "ამხანაგობა", out errorText);

            fieldskeysMap.Add("Name", "BDO_TaxTyp");
            fieldskeysMap.Add("TableName", "OCRD");
            fieldskeysMap.Add("Description", "Tax Type");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 150);
            fieldskeysMap.Add("ValidValues", listValidValues);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            //საქონლის ძიების პარამეტრი
            fieldskeysMap = new Dictionary<string, object>();
            listValidValues = new List<string>();
            listValidValues.Add(""); //-1
            listValidValues.Add("Name"); //1 //დასახელება
            listValidValues.Add("Code"); //2 //კოდი

            fieldskeysMap.Add("Name", "BDO_ItmPrm");
            fieldskeysMap.Add("TableName", "OCRD");
            fieldskeysMap.Add("Description", "Item Search Parameter");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 150);
            fieldskeysMap.Add("ValidValues", listValidValues);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            //შესაბამისობის კოტროლი შესყიდვისას
            fieldskeysMap = new Dictionary<string, object>();
            listValidValues = new List<string>();
            listValidValues.Add(""); //-1
            listValidValues.Add("Not Control"); //1 //არ გაკონტროლდეს
            listValidValues.Add("Control Total Amount"); //2 //მთლიანი თანხის მიხედვით
            listValidValues.Add("Control Positions Amount"); //3 //პოზიციების მიხედვით

            fieldskeysMap.Add("Name", "BDO_MapCnt");
            fieldskeysMap.Add("TableName", "OCRD");
            fieldskeysMap.Add("Description", "Item Mapping Control");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 150);
            fieldskeysMap.Add("ValidValues", listValidValues);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDO_NotInv");
            fieldskeysMap.Add("TableName", "OCRD");
            fieldskeysMap.Add("Description", "Do Not Need Send Waybill");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 1);
            fieldskeysMap.Add("DefaultValue", "N");

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDO_NeedWB");
            fieldskeysMap.Add("TableName", "OCRD");
            fieldskeysMap.Add("Description", "No Posting Purchase Without Waybill");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 1);
            fieldskeysMap.Add("DefaultValue", "N");

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            //მოგების გადასახადი
            fieldskeysMap = new Dictionary<string, object>(); // გათავისუფლებული მოგების გადასახადისგან
            fieldskeysMap.Add("Name", "BDO_PTExem");
            fieldskeysMap.Add("TableName", "OCRD");
            fieldskeysMap.Add("Description", "Profit Tax exempt");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 1);
            fieldskeysMap.Add("DefaultValue", "N");

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // ოფშორულ ქვეყნებში რეგისტრირებული
            fieldskeysMap.Add("Name", "BDO_RIOfsh");
            fieldskeysMap.Add("TableName", "OCRD");
            fieldskeysMap.Add("Description", "Registered in Offshore");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 1);
            fieldskeysMap.Add("DefaultValue", "N");

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            listValidValues = new List<string>();
            listValidValues.Add(""); //-1
            listValidValues.Add("Micro Business Status"); //0 //მიკრო ბიზნესის სტატუსი
            listValidValues.Add("Fixed Paying"); //1 //ფიქსირებული გადამხდელი

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDO_PhysTp");
            fieldskeysMap.Add("TableName", "OCRD");
            fieldskeysMap.Add("Description", "Tax Type");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 150);
            fieldskeysMap.Add("ValidValues", listValidValues);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            GC.Collect();
        }

        public static void createFormItems(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            Dictionary<string, object> formItems;
            string itemName = "";

            formItems = new Dictionary<string, object>();
            itemName = "BDO_TypST";
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", oForm.Items.Item("42").Left);
            formItems.Add("Width", oForm.Items.Item("42").Width);
            formItems.Add("Top", oForm.Items.Item("42").Top + oForm.Items.Item("42").Height + 1);
            formItems.Add("Height", oForm.Items.Item("42").Height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("TaxType"));

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            List<string> listValidValues = new List<string>();
            listValidValues.Add(""); //-1
            listValidValues.Add(BDOSResources.getTranslate("IndividualEnterpreneur")); //1 //ინდივიდუალური მეწარმე
            listValidValues.Add(BDOSResources.getTranslate("LTD")); //2 //შეზღუდული პასუხისმგებლობის საზოგადოება
            listValidValues.Add(BDOSResources.getTranslate("PhysicalEntity")); //3 //ფიზიკური პირი
            listValidValues.Add(BDOSResources.getTranslate("JLC")); //4 //სოლიდარული პასუხისმგებლობის საზოგადოება
            listValidValues.Add(BDOSResources.getTranslate("JSC")); //5 //სააქციო საზოგადოება
            listValidValues.Add(BDOSResources.getTranslate("Commandite")); //6 //კომანდიტური საზოგადოება
            listValidValues.Add(BDOSResources.getTranslate("Cooperative")); //7 //კოოპერატივი //Limited Partnership
            listValidValues.Add(BDOSResources.getTranslate("NonCommercialLegalEntity")); //8 //არაკომერციული იურიდიული პირი
            listValidValues.Add(BDOSResources.getTranslate("ForeignEnterpriseBranch")); //9 //უცხოური საწარმოს ფილიალი
            listValidValues.Add(BDOSResources.getTranslate("ForeignEnterprisePhysicalEntity")); //10 //უცხოური საწარმო - ფიზიკური პირი
            listValidValues.Add(BDOSResources.getTranslate("ForeignEnterpriseLegalEntity")); //11 //უცხოური საწარმო - იურიდიული პირი
            listValidValues.Add(BDOSResources.getTranslate("LEPL")); //12 //საჯარო სამართლის იურიდიული პირი
            listValidValues.Add(BDOSResources.getTranslate("GovernmentBody")); //13 //სახელმწიფო ორგანო //Public Authority
            listValidValues.Add(BDOSResources.getTranslate("Partnership")); //14 //ამხანაგობა

            formItems = new Dictionary<string, object>();
            itemName = "BDO_TaxTyp";
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OCRD");
            formItems.Add("Alias", "U_BDO_TaxTyp");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            formItems.Add("Left", oForm.Items.Item("41").Left);
            formItems.Add("Width", oForm.Items.Item("41").Width);
            formItems.Add("Top", oForm.Items.Item("41").Top + oForm.Items.Item("41").Height + 1);
            formItems.Add("Height", oForm.Items.Item("41").Height);
            formItems.Add("UID", itemName);
            formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
            formItems.Add("DisplayDesc", true);
            formItems.Add("ValidValues", listValidValues);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            //მოგების გადასახადი
            int Left_PR = oForm.Items.Item("BDO_TaxTyp").Left + oForm.Items.Item("BDO_TaxTyp").Width + 10;
            int Width_PR = 200;
            int Top_PR = oForm.Items.Item("BDO_TaxTyp").Top;
            int Height_PR = oForm.Items.Item("BDO_TaxTyp").Height;

            formItems = new Dictionary<string, object>();
            itemName = "BDO_PTExem";
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OCRD");
            formItems.Add("Alias", "U_BDO_PTExem");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            formItems.Add("Left", Left_PR);
            formItems.Add("Width", Width_PR);
            formItems.Add("Top", Top_PR);
            formItems.Add("Height", Height_PR);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("ProfitTaxExempt"));
            formItems.Add("ValOff", "N");
            formItems.Add("ValOn", "Y");
            formItems.Add("DisplayDesc", true);
            formItems.Add("ToPane", 1);
            formItems.Add("FromPane", 1);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "BDO_RIOfsh";
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OCRD");
            formItems.Add("Alias", "U_BDO_RIOfsh");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            formItems.Add("Left", Left_PR);
            formItems.Add("Width", Width_PR);
            formItems.Add("Top", Top_PR);
            formItems.Add("Height", Height_PR);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("RegisteredInOffshore"));
            formItems.Add("ValOff", "N");
            formItems.Add("ValOn", "Y");
            formItems.Add("DisplayDesc", true);
            formItems.Add("ToPane", 1);
            formItems.Add("FromPane", 1);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "BDO_PhsTpt";
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", Left_PR);
            formItems.Add("Width", oForm.Items.Item("242000004").Width);
            formItems.Add("Top", Top_PR);
            formItems.Add("Height", Height_PR);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Status"));

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            List<string> listValidValuesPT = new List<string>();
            listValidValuesPT.Add(""); //-1
            listValidValuesPT.Add(BDOSResources.getTranslate("MicroBusinessStatus")); //0 //მიკრო ბიზნესის სტატუსი
            listValidValuesPT.Add(BDOSResources.getTranslate("FixedPaying")); //1 //ფიქსირებული გადამხდელი

            formItems = new Dictionary<string, object>();
            itemName = "BDO_PhysTp";
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OCRD");
            formItems.Add("Alias", "U_BDO_PhysTp");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            formItems.Add("Left", oForm.Items.Item("242000003").Left);
            formItems.Add("Width", oForm.Items.Item("242000003").Width);
            formItems.Add("Top", Top_PR);
            formItems.Add("Height", Height_PR);
            formItems.Add("UID", itemName);
            formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
            formItems.Add("DisplayDesc", true);
            formItems.Add("ValidValues", listValidValuesPT);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            formItems.Add("Left", oForm.Items.Item("41").Left + oForm.Items.Item("41").Width + 1);
            formItems.Add("Width", 18);
            formItems.Add("Top", oForm.Items.Item("41").Top - 2);
            formItems.Add("Height", 18);
            formItems.Add("Image", "15886_MENU"); //"15886_MENU_CHECKED" //"WS_TOPSEARCH_PICKER"); //WS_TOPSEARCH_PICKER "PNG_1536_MENU" //WS_COCKPIT_SWITCH_UI_MENU_ITEM
            formItems.Add("UID", "BDO_TinBtn"); //

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            //საქონლის ძიების პარამეტრი
            formItems = new Dictionary<string, object>();
            itemName = "BDO_ItmPST";
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", oForm.Items.Item("358").Left);
            formItems.Add("Width", oForm.Items.Item("358").Width);
            formItems.Add("Top", oForm.Items.Item("358").Top + oForm.Items.Item("358").Height + 1);
            formItems.Add("Height", oForm.Items.Item("358").Height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("ItemSearchParameter"));
            formItems.Add("ToPane", 1);
            formItems.Add("FromPane", 1);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            listValidValues = new List<string>();
            listValidValues.Add(""); //-1
            listValidValues.Add(BDOSResources.getTranslate("Name")); //1 //დასახელება
            listValidValues.Add(BDOSResources.getTranslate("Code")); //2 //კოდი

            formItems = new Dictionary<string, object>();
            itemName = "BDO_ItmPrm";
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OCRD");
            formItems.Add("Alias", "U_BDO_ItmPrm");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            formItems.Add("Left", oForm.Items.Item("362").Left);
            formItems.Add("Width", oForm.Items.Item("362").Width);
            formItems.Add("Top", oForm.Items.Item("362").Top + oForm.Items.Item("362").Height + 1);
            formItems.Add("Height", oForm.Items.Item("362").Height);
            formItems.Add("UID", itemName);
            formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
            formItems.Add("DisplayDesc", true);
            formItems.Add("ToPane", 1);
            formItems.Add("FromPane", 1);
            formItems.Add("ValidValues", listValidValues);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            //შესაბამისობის კონტროლი
            formItems = new Dictionary<string, object>();
            itemName = "BDO_MapCST";
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", oForm.Items.Item("BDO_ItmPST").Left);
            formItems.Add("Width", oForm.Items.Item("BDO_ItmPST").Width);
            formItems.Add("Top", oForm.Items.Item("BDO_ItmPST").Top + oForm.Items.Item("BDO_ItmPST").Height + 1);
            formItems.Add("Height", oForm.Items.Item("BDO_ItmPST").Height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("ItemMappingControl"));
            formItems.Add("ToPane", 1);
            formItems.Add("FromPane", 1);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            listValidValues = new List<string>();
            listValidValues.Add(""); //-1
            listValidValues.Add(BDOSResources.getTranslate("NotControl")); //1 //არ გაკონტროლდეს
            listValidValues.Add(BDOSResources.getTranslate("ControlTotalAmount")); //2 //მთლიანი თანხის მიხედვით
            listValidValues.Add(BDOSResources.getTranslate("ControlPositionsAmount")); //3 //პოზიციების მიხედვით

            formItems = new Dictionary<string, object>();
            itemName = "BDO_MapCnt";
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OCRD");
            formItems.Add("Alias", "U_BDO_MapCnt");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            formItems.Add("Left", oForm.Items.Item("BDO_ItmPrm").Left);
            formItems.Add("Width", oForm.Items.Item("BDO_ItmPrm").Width);
            formItems.Add("Top", oForm.Items.Item("BDO_ItmPrm").Top + oForm.Items.Item("BDO_ItmPrm").Height + 1);
            formItems.Add("Height", oForm.Items.Item("BDO_ItmPrm").Height);
            formItems.Add("UID", itemName);
            formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
            formItems.Add("DisplayDesc", true);
            formItems.Add("ToPane", 1);
            formItems.Add("FromPane", 1);
            formItems.Add("ValidValues", listValidValues);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            //ზედნადების გარეშე აკრძალვა
            formItems = new Dictionary<string, object>();
            itemName = "BDO_NeedWB";
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OCRD");
            formItems.Add("Alias", "U_BDO_NeedWB");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            formItems.Add("Left", oForm.Items.Item("BDO_MapCST").Left);
            formItems.Add("Width", oForm.Items.Item("BDO_MapCST").Width);
            formItems.Add("Top", oForm.Items.Item("BDO_MapCST").Top + oForm.Items.Item("149").Height + 1);
            formItems.Add("Height", oForm.Items.Item("BDO_MapCST").Height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("LinkWaybill"));
            formItems.Add("ValOff", "N");
            formItems.Add("ValOn", "Y");
            formItems.Add("DisplayDesc", true);
            formItems.Add("ToPane", 1);
            formItems.Add("FromPane", 1);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "BDO_NotInv";
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OCRD");
            formItems.Add("Alias", "U_BDO_NotInv");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            formItems.Add("Left", oForm.Items.Item("230").Left);
            formItems.Add("Width", oForm.Items.Item("230").Width);
            formItems.Add("Top", oForm.Items.Item("230").Top + oForm.Items.Item("230").Height + 2);
            formItems.Add("Height", oForm.Items.Item("230").Height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("NoNeedTaxInvoice"));
            formItems.Add("ValOff", "N");
            formItems.Add("ValOn", "Y");
            formItems.Add("DisplayDesc", true);
            formItems.Add("Visible", false);
            formItems.Add("ToPane", 10);
            formItems.Add("FromPane", 10);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            GC.Collect();
        }

        public static void comboSelect(SAPbouiCOM.Form oForm, string itemUID, bool before_Action, out string errorText)
        {
            errorText = null;
            try
            {
                if (before_Action == false)
                {
                    if (itemUID == "40")
                    {
                        string cardType = oForm.DataSources.DBDataSources.Item("OCRD").GetValue("CardType", 0).Trim();
                        string BDO_NotInv = oForm.DataSources.DBDataSources.Item("OCRD").GetValue("U_BDO_NotInv", 0).Trim();
                        int panelLevel = oForm.PaneLevel;

                        if (cardType == "C" || cardType == "L") //თუ მყიდველია ან პოტენციური კლიენტი
                        {
                        }
                        else
                        {
                            oForm.PaneLevel = 10;
                            if (BDO_NotInv == "Y")
                            {
                                SAPbouiCOM.CheckBox oBDO_NotInv = ((SAPbouiCOM.CheckBox)(oForm.Items.Item("BDO_NotInv").Specific));
                                oBDO_NotInv.Checked = false;
                            }
                        }
                        oForm.PaneLevel = panelLevel;
                    }

                    if (itemUID == "BDO_TaxTyp")
                    {
                        //თუ მოგების გადასახადის ველები არ ჩანს, უნდა გამოირთოს
                        string BDO_TaxType = oForm.DataSources.DBDataSources.Item("OCRD").GetValue("U_BDO_TaxTyp", 0).Trim();
                        if ((BDO_TaxType == "" || BDO_TaxType == "1" || BDO_TaxType == "3" || BDO_TaxType == "10" || BDO_TaxType == "11") == true)
                        {
                            SAPbouiCOM.CheckBox oBDO_PTExem = ((SAPbouiCOM.CheckBox)(oForm.Items.Item("BDO_PTExem").Specific));
                            oBDO_PTExem.Checked = false;
                        }
                        if ((BDO_TaxType == "10" || BDO_TaxType == "11") == false)
                        {
                            SAPbouiCOM.CheckBox oBDO_RIOfsh = ((SAPbouiCOM.CheckBox)(oForm.Items.Item("BDO_RIOfsh").Specific));
                            oBDO_RIOfsh.Checked = false;
                        }
                        if (BDO_TaxType != "3")
                        {
                            SAPbouiCOM.ComboBox oBDO_PhysTp = ((SAPbouiCOM.ComboBox)(oForm.Items.Item("BDO_PhysTp").Specific));
                            oBDO_PhysTp.Select("-1", SAPbouiCOM.BoSearchKey.psk_ByValue);
                        }
                    }

                    if (itemUID == "178")
                    {

                        int paneLevel = oForm.PaneLevel;

                        SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("178").Specific));
                        
                        for (int j = 1; j <= oMatrix.RowCount; )
                        {
                            
                            string countryCode = oMatrix.Columns.Item("8").Cells.Item(j).Specific.Value;

                            if (countryCode != "GE")
                            {                                
                                string cardType = oForm.DataSources.DBDataSources.Item("OCRD").GetValue("CardType", 0).Trim();
                
                                if (cardType == "C" || cardType == "L") //თუ მყიდველია ან პოტენციური კლიენტი
                                {
                                    oForm.PaneLevel = 10;
                                    SAPbouiCOM.CheckBox oBDO_NotInv = ((SAPbouiCOM.CheckBox)(oForm.Items.Item("BDO_NotInv").Specific));
                                    oBDO_NotInv.Checked = true;
                                    
                                }

                            }
                            goto exitLoop;
                        }

                    exitLoop:
                        oForm.PaneLevel = paneLevel;
                        return;
                    }

                }
                else if (before_Action == true)
                {
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

        public static void setVisibleFormItems( SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            try
            {
                string cardType = oForm.DataSources.DBDataSources.Item("OCRD").GetValue("CardType", 0).Trim();
                string BDO_NotInv = oForm.DataSources.DBDataSources.Item("OCRD").GetValue("U_BDO_NotInv", 0).Trim();
                string BDO_TaxType = oForm.DataSources.DBDataSources.Item("OCRD").GetValue("U_BDO_TaxTyp", 0).Trim();
                int panelLevel = oForm.PaneLevel;

                if (cardType == "C" || cardType == "L") //თუ მყიდველია ან პოტენციური კლიენტი
                {
                    //oForm.PaneLevel = 10;
                    //oForm.Items.Item("BDO_NotInv").Visible = true;

                    //მომწოდებლის რეკვიზიტები
                    oForm.PaneLevel = 1;
                    oForm.Items.Item("BDO_NeedWB").Visible = false;
                    //მომწოდებლის რეკვიზიტები
                }
                else
                {
                    //oForm.PaneLevel = 10;
                    //oForm.Items.Item("BDO_NotInv").Visible = false;

                    //მომწოდებლის რეკვიზიტები
                    oForm.PaneLevel = 1;
                    oForm.Items.Item("BDO_NeedWB").Visible = true;
                    //მომწოდებლის რეკვიზიტები
                }

                oForm.PaneLevel = panelLevel;

                oForm.Items.Item("128").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                oForm.Items.Item("BDO_PTExem").Visible = ((BDO_TaxType == "" || BDO_TaxType == "1" || BDO_TaxType == "3" || BDO_TaxType == "10" || BDO_TaxType == "11") == false);
                oForm.Items.Item("BDO_RIOfsh").Visible = (BDO_TaxType == "10" || BDO_TaxType == "11");
                oForm.Items.Item("BDO_PhysTp").Visible = (BDO_TaxType == "3");
                oForm.Items.Item("BDO_PhsTpt").Visible = (BDO_TaxType == "3");
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

        public static void BDO_TinBtn_OnClick(  SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            string tin = oForm.DataSources.DBDataSources.Item("OCRD").GetValue("LicTradNum", 0).Trim();
            if (tin == "")
            {
                errorText = BDOSResources.getTranslate("TINNotFill");
                return;
            }

            Dictionary<string, string> rsSettings = CompanyDetails.getRSSettings( out errorText);
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

            bool isVatPayer = oWayBill.is_vat_payer_tin(tin, out errorText);
            if (errorText != null)
            {
                return;
            }

            try
            {
                int panelLevel = oForm.PaneLevel;
                string vatStatus = oForm.DataSources.DBDataSources.Item("OCRD").GetValue("VatStatus", 0).Trim();
                string cardType = oForm.DataSources.DBDataSources.Item("OCRD").GetValue("CardType", 0).Trim();
                string BDO_NotInv = oForm.DataSources.DBDataSources.Item("OCRD").GetValue("U_BDO_NotInv", 0).Trim();

                if (cardType == "S") //თუ მომწოდებელია
                {
                    if (isVatPayer == true & vatStatus != "Y") //Liable
                    {
                        oForm.PaneLevel = 10;
                        SAPbouiCOM.ComboBox oVatStatus = ((SAPbouiCOM.ComboBox)(oForm.Items.Item("76").Specific));
                        oVatStatus.Select("Liable", SAPbouiCOM.BoSearchKey.psk_ByDescription);
                    }
                    else if (isVatPayer == false & vatStatus != "N") //Exempt
                    {
                        oForm.PaneLevel = 10;
                        SAPbouiCOM.ComboBox oVatStatus = ((SAPbouiCOM.ComboBox)(oForm.Items.Item("76").Specific));
                        oVatStatus.Select("Exempt", SAPbouiCOM.BoSearchKey.psk_ByDescription);
                    }
                }
                //if (cardType == "C" || cardType == "L") //თუ მყიდველია ან პოტენციური კლიენტი
                //{
                    if (isVatPayer == true & BDO_NotInv != "N")
                    {
                        oForm.PaneLevel = 10;
                        SAPbouiCOM.CheckBox oBDO_NotInv = ((SAPbouiCOM.CheckBox)(oForm.Items.Item("BDO_NotInv").Specific));
                        oBDO_NotInv.Checked = false;
                    }
                    else if (isVatPayer == false & BDO_NotInv != "Y")
                    {
                        oForm.PaneLevel = 10;
                        SAPbouiCOM.CheckBox oBDO_NotInv = ((SAPbouiCOM.CheckBox)(oForm.Items.Item("BDO_NotInv").Specific));
                        oBDO_NotInv.Checked = true;
                    }
                //}

                oForm.PaneLevel = panelLevel;
            }
            catch (Exception ex)
            {
                int errCode;
                string errMsg;

                Program.oCompany.GetLastError(out errCode, out errMsg);
                errorText = BDOSResources.getTranslate("ErrorDescription") + " : " + errMsg + "! " + BDOSResources.getTranslate("Code") + " : " + errCode + "! " + BDOSResources.getTranslate("OtherInfo") + " : " + ex.Message;
            }

            TaxService oTaxService = new TaxService();
            Dictionary<string, object> getTPInfo_result = oTaxService.GetTPInfo(tin, out errorText);
            if (errorText != null)
            {
                return;
            }
            bool getInitFromTIN = true;
            if (getTPInfo_result != null)
            {
                string name = getTPInfo_result["Name"].ToString();
                if (name == string.Empty)
                {
                    //errorText = BDOSResources.getTranslate("NotRecognizeObjectByTINInEnreg");
                    //return;
                }
                else
                {

                    getInitFromTIN = false;

                    string cardName = oForm.DataSources.DBDataSources.Item("OCRD").GetValue("CardName", 0).Trim();

                    if (cardName != name)
                    {
                        int changeCardName = Program.uiApp.MessageBox(BDOSResources.getTranslate("NameMismatchFromEnregDoYouWantEdit"), 1, BDOSResources.getTranslate("Yes"), BDOSResources.getTranslate("No"), "");
                        if (changeCardName == 1)
                        {
                            SAPbouiCOM.EditText oCardName = ((SAPbouiCOM.EditText)(oForm.Items.Item("7").Specific));
                            oCardName.Value = name;
                        }
                    }

                    //სამართებლივი ფორმები --->
                    string OrganizationType = getTPInfo_result["OrganizationTypeShort"].ToString();
                    SAPbouiCOM.ComboBox oBDO_TaxTyp = ((SAPbouiCOM.ComboBox)(oForm.Items.Item("BDO_TaxTyp").Specific));

                    switch (OrganizationType)
                    {
                        case "ი/მ": if (oBDO_TaxTyp.Value.Trim() != "1") //ინდივიდუალური მეწარმე
                            {
                                oBDO_TaxTyp.Select("1", SAPbouiCOM.BoSearchKey.psk_ByValue);
                            };
                            break;
                        case "შპს": if (oBDO_TaxTyp.Value.Trim() != "2") //შეზღუდული პასუხისმგებლობის საზოგადოება
                            {
                                oBDO_TaxTyp.Select("2", SAPbouiCOM.BoSearchKey.psk_ByValue);
                            };
                            break;
                        case "ფ/პ": if (oBDO_TaxTyp.Value.Trim() != "3") //ფიზიკური პირი
                            {
                                oBDO_TaxTyp.Select("3", SAPbouiCOM.BoSearchKey.psk_ByValue);
                            };
                            break;
                        case "სპს": if (oBDO_TaxTyp.Value.Trim() != "4") //სოლიდარული პასუხისმგებლობის საზოგადოება
                            {
                                oBDO_TaxTyp.Select("4", SAPbouiCOM.BoSearchKey.psk_ByValue);
                            };
                            break;
                        case "სს": if (oBDO_TaxTyp.Value.Trim() != "5") //სააქციო საზოგადოება
                            {
                                oBDO_TaxTyp.Select("5", SAPbouiCOM.BoSearchKey.psk_ByValue);
                            };
                            break;
                        case "კს": if (oBDO_TaxTyp.Value.Trim() != "6") //კომანდიტური საზოგადოება
                            {
                                oBDO_TaxTyp.Select("6", SAPbouiCOM.BoSearchKey.psk_ByValue);
                            };
                            break;
                        case "კოოპ.": if (oBDO_TaxTyp.Value.Trim() != "7") //კოოპერატივი
                            {
                                oBDO_TaxTyp.Select("7", SAPbouiCOM.BoSearchKey.psk_ByValue);
                            };
                            break;
                        case "ა(ა)იპ": if (oBDO_TaxTyp.Value.Trim() != "8") //არაკომერციული იურიდიული პირი
                            {
                                oBDO_TaxTyp.Select("8", SAPbouiCOM.BoSearchKey.psk_ByValue);
                            };
                            break;
                        case "უცხ.საწ.ფილ.": if (oBDO_TaxTyp.Value.Trim() != "9") //უცხოური საწარმოს ფილიალი
                            {
                                oBDO_TaxTyp.Select("9", SAPbouiCOM.BoSearchKey.psk_ByValue);
                            };
                            break;
                        case "სსიპ": if (oBDO_TaxTyp.Value.Trim() != "12") //საჯარო სამართლის იურიდიული პირი
                            {
                                oBDO_TaxTyp.Select("12", SAPbouiCOM.BoSearchKey.psk_ByValue);
                            };
                            break;
                        case "ამხანაგობა": if (oBDO_TaxTyp.Value.Trim() != "14") //ამხანაგობა
                            {
                                oBDO_TaxTyp.Select("14", SAPbouiCOM.BoSearchKey.psk_ByValue);
                            };
                            break;
                        case "უცხ.აიპ ფილ.":
                            //if (oBDO_TaxTyp.Value.Trim() != "")
                            //{
                            //    //oBDO_TaxTyp.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue);
                            //};
                            break;
                        default:
                            break;
                    }
                    //<--- სამართებლივი ფორმები 

                    //მისამართის ჩაწერა --->
                    try
                    {
                        string Address = getTPInfo_result["Address"].ToString();
                        SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("178").Specific));
                        SAPbouiCOM.Matrix oMatrixBillTo = ((SAPbouiCOM.Matrix)(oForm.Items.Item("69").Specific));

                        SAPbouiCOM.EditText oEdit;
                        SAPbouiCOM.EditText oEditBillTo;
                        int paneLevel = oForm.PaneLevel;

                        for (int i = 1; i <= oMatrixBillTo.RowCount; i++)
                        {
                            oEditBillTo = oMatrixBillTo.Columns.Item("20").Cells.Item(i).Specific;
                            if (oEditBillTo.Value == "Bill To" || oEditBillTo.Value == "Bill to")
                            {
                                for (int k = i + 1; k <= oMatrixBillTo.RowCount; k++)
                                {
                                    oForm.PaneLevel = 7;
                                    oMatrixBillTo.Columns.Item("20").Cells.Item(k).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                    oEditBillTo = oMatrixBillTo.Columns.Item("20").Cells.Item(k).Specific;
                                    if (oEditBillTo.Value == "იურ. მისამართი")
                                    {
                                        for (int j = 1; j <= oMatrix.RowCount; )
                                        {
                                            oEdit = oMatrix.Columns.Item("2").Cells.Item(j).Specific;
                                            if (oEdit.Value != Address)
                                            {
                                                oEdit.Value = Address;
                                            }
                                            goto exitLoop;
                                        }
                                    }
                                    else if (oEditBillTo.Value == "Define New")
                                    {
                                        for (int j = 1; j <= oMatrix.RowCount; )
                                        {
                                            oEdit = oMatrix.Columns.Item("1").Cells.Item(j).Specific;
                                            oEdit.Value = "იურ. მისამართი";
                                            oEdit = oMatrix.Columns.Item("2").Cells.Item(j).Specific;
                                            oEdit.Value = Address;
                                            goto exitLoop;
                                        }
                                    }
                                }
                            }
                        }
                        oForm.PaneLevel = paneLevel;

                    exitLoop:
                        oForm.PaneLevel = paneLevel;
                        return;
                    }
                    catch (Exception ex)
                    {
                        int errCode;
                        string errMsg;

                        Program.oCompany.GetLastError(out errCode, out errMsg);
                        errorText = BDOSResources.getTranslate("ErrorDescription") + " : " + errMsg + "! " + BDOSResources.getTranslate("Code") + " : " + errCode + "! " + BDOSResources.getTranslate("OtherInfo") + " : " + ex.Message;
                        return;
                    }
                    //<--- მისამართის ჩაწერა                               
                }
            }
        
            if (getInitFromTIN == true)
            {                
                string name = BDO_Waybills.getInitFromTIN( tin, out errorText);

                if (name != "")
                {      
                    string cardName = oForm.DataSources.DBDataSources.Item("OCRD").GetValue("CardName", 0).Trim();

                    int changeCardName = 1;

                    if (cardName != "")
                    {
                        changeCardName = Program.uiApp.MessageBox(BDOSResources.getTranslate("NameMismatchFromEnregDoYouWantEditDr"), 1, BDOSResources.getTranslate("Yes"), BDOSResources.getTranslate("No"), "");
        }

                    if (changeCardName == 1)
                    {
                        SAPbouiCOM.EditText oCardName = ((SAPbouiCOM.EditText)(oForm.Items.Item("7").Specific));
                        oCardName.Value = name;
                    }
                }
                else
                {
                    errorText = BDOSResources.getTranslate("NotRecognizeObjectByTINInEnreg");
                    return;
                }
            }

        }

        
        public static void uiApp_FormDataEvent(  ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(BusinessObjectInfo.FormTypeEx, Program.currentFormCount);

            if (oForm.TypeEx == "134")
            {
                if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD & BusinessObjectInfo.BeforeAction == false)
                {
                    BusinessPartners.setVisibleFormItems( oForm, out errorText);
                }

                //გსნ-თი და ბპ ტიპით უნიკალურობის კონტროლი
                if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD || BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE)
                {
                    if (BusinessObjectInfo.BeforeAction == true)
                    {
                        string CardType = oForm.DataSources.DBDataSources.Item(0).GetValue("CardType", 0).Trim();

                        string CardTypeDescription = "";
                        string CardTypeFilter = "";

                        if (CardType == "S")
                        {
                            CardTypeFilter = " ('S') ";
                            CardTypeDescription = BDOSResources.getTranslate("Supplier");
                        }
                        else
                        {
                            CardTypeFilter = " ('C','L') ";
                            CardTypeDescription = BDOSResources.getTranslate("Customer");
                        }

                        SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        string query = @"SELECT * FROM ""OCRD"" " +
                                        @"WHERE ""LicTradNum"" = '" + oForm.DataSources.DBDataSources.Item(0).GetValue("LicTradNum", 0).Trim() + "' " +
                                        @" AND ""CardType"" IN " + CardTypeFilter + "  " +
                                        @" AND ""CardCode"" <> '" + oForm.DataSources.DBDataSources.Item(0).GetValue("CardCode", 0).Trim() + "'";

                        oRecordSet.DoQuery(query);

                        if (!oRecordSet.EoF)
                        {
                            int answer = 0;
                            answer = Program.uiApp.MessageBox(BDOSResources.getTranslate("DuplicateLictTradNumBPofType") + " " + CardTypeDescription + ", " + BDOSResources.getTranslate("Continue") + "?", 1, BDOSResources.getTranslate("Yes"), BDOSResources.getTranslate("No"), "");

                            if (answer != 1)
                            {
                                BubbleEvent = false;
                            }
                        }
                    }
                }
            }
        }

        public static void uiApp_ItemEvent(  string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD & pVal.BeforeAction == true)
                {
                    BusinessPartners.createFormItems(oForm, out errorText);
                    BusinessPartners.setVisibleFormItems( oForm, out errorText);
                }

                if (pVal.ItemUID == "BDO_TinBtn")
                {
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK & pVal.BeforeAction == false)
                    {
                        SAPbouiCOM.Button oButton = (SAPbouiCOM.Button)oForm.Items.Item("BDO_TinBtn").Specific;
                        oButton.Image = "15886_MENU_CHECKED";
                        oForm.Freeze(true);
                        BusinessPartners.BDO_TinBtn_OnClick( oForm, out errorText);
                        oForm.Freeze(false);
                        oButton.Image = "15886_MENU";
                        if (errorText != null)
                        {
                            Program.uiApp.SetStatusBarMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                            return;
                        }
                    }
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_COMBO_SELECT && pVal.BeforeAction==false)
                {
                    oForm.Freeze(true);
                    comboSelect(oForm, pVal.ItemUID, pVal.BeforeAction, out errorText);
                    setVisibleFormItems(oForm, out errorText);
                    oForm.Freeze(false);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED & pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "262")
                    {
                        oForm.Freeze(true);
                        BusinessPartners.setVisibleFormItems( oForm, out errorText);
                        oForm.Freeze(false);
                    }
                }
            }
        }
    }
}
