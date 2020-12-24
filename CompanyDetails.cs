using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using SAPbobsCOM;
using SAPbouiCOM;
using static BDO_Localisation_AddOn.Program;

namespace BDO_Localisation_AddOn
{
    static partial class CompanyDetails
    {
        static int FirstItemHeight = 37;
        static int ItemDistanceVertical = 20;

        //static SAPbouiCOM.UserDataSource BdgCfNameUDS;

        public static void createUserFields(out string errorText)
        {
            errorText = null;
            Dictionary<string, object> fieldskeysMap;
            List<string> listValidValues;

            fieldskeysMap = new Dictionary<string, object>();
            listValidValues = new List<string>();
            listValidValues.Add("Program Code"); //0 // პროგრამის კოდი
            listValidValues.Add("Article"); //1 // არტიკული
            listValidValues.Add("Main Barcode"); //2 // ძირითადი შტრიხკოდი

            fieldskeysMap.Add("Name", "BDO_ItmCod");
            fieldskeysMap.Add("TableName", "OADM");
            fieldskeysMap.Add("Description", "Item Code For Waybill");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);
            fieldskeysMap.Add("ValidValues", listValidValues);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            listValidValues = new List<string>();
            listValidValues.Add("By Organization"); //0 // ორგანიზაციის მიხედვით
            listValidValues.Add("By User"); //1 // მომხმარებლის მიხედვით

            fieldskeysMap.Add("Name", "BDO_UsrTyp");
            fieldskeysMap.Add("TableName", "OADM");
            fieldskeysMap.Add("Description", "User Type For RS.GE");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);
            fieldskeysMap.Add("ValidValues", listValidValues);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            listValidValues = new List<string>();
            listValidValues.Add("HTTP"); //0
            listValidValues.Add("HTTPS"); //1

            fieldskeysMap.Add("Name", "BDO_PrtTyp");
            fieldskeysMap.Add("TableName", "OADM");
            fieldskeysMap.Add("Description", "Protocol Type For RS.GE");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);
            fieldskeysMap.Add("ValidValues", listValidValues);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            listValidValues = new List<string>();
            listValidValues.Add("With Transport"); //0 //ტრანსპორტირებით
            listValidValues.Add("Without Transport"); //1 //ტრანსპორტირების გარეშე

            fieldskeysMap.Add("Name", "BDO_WblTyp");
            fieldskeysMap.Add("TableName", "OADM");
            fieldskeysMap.Add("Description", "Waybill Type");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);
            fieldskeysMap.Add("ValidValues", listValidValues);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            /////////////////
            fieldskeysMap = new Dictionary<string, object>();
            listValidValues = new List<string>();
            listValidValues.Add("No Authorization"); 
            listValidValues.Add("Read Only");
            listValidValues.Add("Full");

            fieldskeysMap.Add("Name", "BDOSWblAut");
            fieldskeysMap.Add("TableName", "OADM");
            fieldskeysMap.Add("Description", "Waybill Authorization");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);
            fieldskeysMap.Add("ValidValues", listValidValues);
            fieldskeysMap.Add("DefaultValue", "2");

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            listValidValues = new List<string>();
            listValidValues.Add("No Authorization");
            listValidValues.Add("Read Only");
            listValidValues.Add("Approval ");

            fieldskeysMap.Add("Name", "BDOSTaxAut");
            fieldskeysMap.Add("TableName", "OADM");
            fieldskeysMap.Add("Description", "Tax Authorization");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);
            fieldskeysMap.Add("ValidValues", listValidValues);
            fieldskeysMap.Add("DefaultValue", "2");

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSDecAtt"); 
            fieldskeysMap.Add("TableName", "OADM");
            fieldskeysMap.Add("Description", "Attach Tax Invoice Declaration");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 1);
            fieldskeysMap.Add("DefaultValue", "Y");

            UDO.addUserTableFields(fieldskeysMap, out errorText);
            //////////////////////

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDO_SU");
            fieldskeysMap.Add("TableName", "OADM");
            fieldskeysMap.Add("Description", "User Name For RS.GE");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDO_SP");
            fieldskeysMap.Add("TableName", "OADM");
            fieldskeysMap.Add("Description", "Password For RS.GE");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSAllDev");
            fieldskeysMap.Add("TableName", "OADM");
            fieldskeysMap.Add("Description", "Allowable Deviation");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            //VAT TYPES
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDO_RSVAT0");
            fieldskeysMap.Add("TableName", "OADM");
            fieldskeysMap.Add("Description", "Vat Type Normal");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDO_RSVAT1");
            fieldskeysMap.Add("TableName", "OADM");
            fieldskeysMap.Add("Description", "Vat Type Zero");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDO_RSVAT2");
            fieldskeysMap.Add("TableName", "OADM");
            fieldskeysMap.Add("Description", "Vat Type None");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            // მოგების გადასახადის დაბეგვრის სისტემა
            fieldskeysMap = new Dictionary<string, object>();
            listValidValues = new List<string>();
            listValidValues.Add(BDOSResources.getTranslate("DifferenceBetweenTheGrossIncomeAndTheDeductions")); //0
            listValidValues.Add(BDOSResources.getTranslate("ProfitSharingTax")); //1

            fieldskeysMap.Add("Name", "BDO_TaxTyp");
            fieldskeysMap.Add("TableName", "OADM");
            fieldskeysMap.Add("Description", "მოგების გადასახადის დაბეგვრის სისტემა");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);
            fieldskeysMap.Add("ValidValues", listValidValues);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDO_PrTxRt");
            fieldskeysMap.Add("TableName", "OADM");
            fieldskeysMap.Add("Description", "განაკვეთი %");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Rate);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDO_CapAcc");
            fieldskeysMap.Add("TableName", "OADM");
            fieldskeysMap.Add("Description", "•	კაპიტალის ანგარიში");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDO_TaxAcc");
            fieldskeysMap.Add("TableName", "OADM");
            fieldskeysMap.Add("Description", "მოგების გადასახადის ბიუჯეტთან ანგარიშსწორების ანგარიში");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //დაბეგვრის ობიექტი Down Payment Request
            fieldskeysMap.Add("Name", "BDO_prBsDR");
            fieldskeysMap.Add("TableName", "OADM");
            fieldskeysMap.Add("Description", "დაბეგვრის ობიექტი Down Payment Request - სთვის");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //დაბეგვრის ობიექტი Goods Issue
            fieldskeysMap.Add("Name", "BDO_prBsGI");
            fieldskeysMap.Add("TableName", "OADM");
            fieldskeysMap.Add("Description", "დაბეგვრის ობიექტი Goods Issue - სთვის");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // დაბეგვრის ობიექტის Down Payment Request სახელი
            fieldskeysMap.Add("Name", "BDO_prDRDs");
            fieldskeysMap.Add("TableName", "OADM");
            fieldskeysMap.Add("Description", "Profit Base Description");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 100);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // დაბეგვრის ობიექტის Goods Issue სახელი
            fieldskeysMap.Add("Name", "BDO_prGIDs");
            fieldskeysMap.Add("TableName", "OADM");
            fieldskeysMap.Add("Description", "Profit Base Description");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 100);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // Chief Accountant
            fieldskeysMap.Add("Name", "BDOSChfAct");
            fieldskeysMap.Add("TableName", "OADM");
            fieldskeysMap.Add("Description", "Chief Accountant");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 100);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // Chief Accountant
            fieldskeysMap.Add("Name", "BDOSIBKW");
            fieldskeysMap.Add("TableName", "OADM");
            fieldskeysMap.Add("Description", "Intbank payroll keywords");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 254);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSResSrv"); //რეზერვის ინვოისიდან ფაქტურის გამოწერა როგორც მომსახურებაზე
            fieldskeysMap.Add("TableName", "OADM");
            fieldskeysMap.Add("Description", "Create Service type Tax invoice from AR Reserve Invoice");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 1);
            fieldskeysMap.Add("DefaultValue", "N");

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSInvoiceRestr"); //ზედნადების გარეშე ანგარიშ-ფაქტურის გამოწერის შეზღუდვა
            fieldskeysMap.Add("TableName", "OADM");
            fieldskeysMap.Add("Description", "Restrict tax invoice creation without waybill");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 1);
            fieldskeysMap.Add("DefaultValue", "N");

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            //საპენსიო
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSPnCoP");
            fieldskeysMap.Add("TableName", "OADM");
            fieldskeysMap.Add("Description", "Pension for Company");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSPnPh");
            fieldskeysMap.Add("TableName", "OADM");
            fieldskeysMap.Add("Description", "Pension for Physical Entity");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            //სამშენებლო და დეველოპერული კომპანია     
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSDevCmp");
            fieldskeysMap.Add("TableName", "OADM");
            fieldskeysMap.Add("Description", "Building or Development Company");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 1);
            fieldskeysMap.Add("DefaultValue", "N");

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            //კომპანიის საპენსიო ხარჯის გამიჯვნა    
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSPnAcc");
            fieldskeysMap.Add("TableName", "OADM");
            fieldskeysMap.Add("Description", "Separation ოf Company Pension Expenses");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 1);
            fieldskeysMap.Add("DefaultValue", "N");

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //Budget Cash Flow
            fieldskeysMap.Add("Name", "BDOSDefCf");
            fieldskeysMap.Add("TableName", "OADM");
            fieldskeysMap.Add("Description", "Default Budget Cash Flow");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSEnbFlM"); //საწვავის მართვის მოდულის გამოყენება
            fieldskeysMap.Add("TableName", "OADM");
            fieldskeysMap.Add("Description", "Enable Using Fuel Management Module");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 1);
            fieldskeysMap.Add("DefaultValue", "N");

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //საწვავის აღრიცხვის საწყობი
            fieldskeysMap.Add("Name", "BDOSFlWhs");
            fieldskeysMap.Add("TableName", "OADM");
            fieldskeysMap.Add("Description", "Fuel Warehouse Code");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 8);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "Discount");
            fieldskeysMap.Add("TableName", "OADM");
            fieldskeysMap.Add("Description", "Using Discount");
            fieldskeysMap.Add("Type", BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 1);
            fieldskeysMap.Add("DefaultValue", "N");

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            GC.Collect();
        }

        public static void createFormItems(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            //oForm.Width = oForm.Width * 2;

            Dictionary<string, object> formItems;
            string itemName = "";
            SAPbouiCOM.Item oFolder;

            //RS.GE (ჩანართი)
            oFolder = oForm.Items.Item("36");
            formItems = new Dictionary<string, object>();
            itemName = "BDO_RSGE";
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_FOLDER);
            formItems.Add("Left", oFolder.Left + oFolder.Width);
            formItems.Add("Width", 80);
            formItems.Add("Top", oFolder.Top);
            formItems.Add("Height", oFolder.Height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", "RS.GE");
            formItems.Add("Pane", 200);
            formItems.Add("ValOn", "0");
            formItems.Add("ValOff", itemName);
            formItems.Add("GroupWith", "36");
            formItems.Add("AffectsFormMode", false);
            formItems.Add("Description", BDOSResources.getTranslate("DataForRS"));

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            //Internet Banking(ჩანართი)
            oFolder = oForm.Items.Item("36");
            formItems = new Dictionary<string, object>();
            itemName = "BDO_INBNK";
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_FOLDER);
            formItems.Add("Left", oFolder.Left + oFolder.Width);
            formItems.Add("Width", 50);
            formItems.Add("Top", oFolder.Top);
            formItems.Add("Height", oFolder.Height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("IntBanking"));
            formItems.Add("Pane", 16);
            formItems.Add("ValOn", "0");
            formItems.Add("ValOff", itemName);
            formItems.Add("GroupWith", "BDO_RSGE");
            formItems.Add("AffectsFormMode", false);
            formItems.Add("Description", BDOSResources.getTranslate("IntBanking"));

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            int top = 37;

            //პროტოკოლის რეჟიმი
            formItems = new Dictionary<string, object>();
            itemName = "PrtTyp";
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", 13);
            formItems.Add("Width", 207);
            formItems.Add("Top", top);
            formItems.Add("Height", 15);
            formItems.Add("Caption", BDOSResources.getTranslate("ProtocolMode"));
            formItems.Add("UID", itemName);
            formItems.Add("FromPane", 12);
            formItems.Add("ToPane", 12);
            formItems.Add("LinkTo", "BDO_PrtTyp");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "BDO_PrtTyp";
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OADM");
            formItems.Add("Alias", "U_BDO_PrtTyp");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            formItems.Add("Left", 230);
            formItems.Add("Width", 250);
            formItems.Add("Top", top + 1);
            formItems.Add("Height", 14);
            formItems.Add("UID", itemName);
            formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
            formItems.Add("DisplayDesc", true);
            formItems.Add("FromPane", 12);
            formItems.Add("ToPane", 12);
            formItems.Add("Description", BDOSResources.getTranslate("ProtocolMode"));

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + 15;

            //სერვისის მომხმარებელი
            formItems = new Dictionary<string, object>();
            itemName = "UsrTyp";
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", 13);
            formItems.Add("Width", 207);
            formItems.Add("Top", top);
            formItems.Add("Height", 15);
            formItems.Add("Caption", BDOSResources.getTranslate("ServiceUser"));
            formItems.Add("UID", itemName);
            formItems.Add("FromPane", 12);
            formItems.Add("ToPane", 12);
            formItems.Add("LinkTo", "BDO_UsrTyp");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            List<string> listValidValues = new List<string>();
            listValidValues.Add(BDOSResources.getTranslate("ByOrganization")); //0 // ორგანიზაციის მიხედვით
            listValidValues.Add(BDOSResources.getTranslate("ByUser")); //1 // მომხმარებლის მიხედვით

            formItems = new Dictionary<string, object>();
            itemName = "BDO_UsrTyp";
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OADM");
            formItems.Add("Alias", "U_BDO_UsrTyp");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            formItems.Add("Left", 230);
            formItems.Add("Width", 250);
            formItems.Add("Top", top + 1);
            formItems.Add("Height", 14);
            formItems.Add("UID", itemName);
            formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
            formItems.Add("DisplayDesc", true);
            formItems.Add("FromPane", 12);
            formItems.Add("ToPane", 12);
            formItems.Add("Description", BDOSResources.getTranslate("ServiceUser"));
            formItems.Add("ValidValues", listValidValues);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + 15;

            //ზედნადების ტიპი (მინიშნების გარეშე)
            formItems = new Dictionary<string, object>();
            itemName = "WblTyp";
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", 13);
            formItems.Add("Width", 207);
            formItems.Add("Top", top);
            formItems.Add("Height", 15);
            formItems.Add("Caption", BDOSResources.getTranslate("WaybillTypeDefault"));
            formItems.Add("UID", itemName);
            formItems.Add("FromPane", 12);
            formItems.Add("ToPane", 12);
            formItems.Add("LinkTo", "BDO_WblTyp");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            listValidValues = new List<string>();
            listValidValues.Add(BDOSResources.getTranslate("WithTransport")); //0 //ტრანსპორტირებით
            listValidValues.Add(BDOSResources.getTranslate("WithoutTransport")); //1 //ტრანსპორტირების გარეშე

            formItems = new Dictionary<string, object>();
            itemName = "BDO_WblTyp";
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OADM");
            formItems.Add("Alias", "U_BDO_WblTyp");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            formItems.Add("Left", 230);
            formItems.Add("Width", 250);
            formItems.Add("Top", top + 1);
            formItems.Add("Height", 14);
            formItems.Add("UID", itemName);
            formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
            formItems.Add("DisplayDesc", true);
            formItems.Add("FromPane", 12);
            formItems.Add("ToPane", 12);
            formItems.Add("Description", BDOSResources.getTranslate("WaybillTypeDefault"));
            formItems.Add("ValidValues", listValidValues);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + 15;

            //ნომენკლატურის კოდი
            formItems = new Dictionary<string, object>();
            itemName = "ItmID";
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", 13);
            formItems.Add("Width", 207);
            formItems.Add("Top", top);
            formItems.Add("Height", 15);
            formItems.Add("Caption", BDOSResources.getTranslate("ItemCode"));
            formItems.Add("UID", itemName);
            formItems.Add("FromPane", 12);
            formItems.Add("ToPane", 12);
            formItems.Add("LinkTo", "BDO_ItmCod");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            listValidValues = new List<string>();
            listValidValues.Add(BDOSResources.getTranslate("ProgramCode")); //0 // პროგრამის კოდი
            listValidValues.Add(BDOSResources.getTranslate("Article")); //1 // არტიკული
            listValidValues.Add(BDOSResources.getTranslate("MainBarcode")); //2 // ძირითადი შტრიხკოდი

            formItems = new Dictionary<string, object>();
            itemName = "BDO_ItmCod";
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OADM");
            formItems.Add("Alias", "U_BDO_ItmCod");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            formItems.Add("Left", 230);
            formItems.Add("Width", 250);
            formItems.Add("Top", top + 1);
            formItems.Add("Height", 14);
            formItems.Add("UID", itemName);
            formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
            formItems.Add("DisplayDesc", true);
            formItems.Add("FromPane", 12);
            formItems.Add("ToPane", 12);
            formItems.Add("Description", BDOSResources.getTranslate("ItemCodeForWaybill"));
            formItems.Add("ValidValues", listValidValues);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            bool multiSelection = false;
            string objectType = "5"; //Warehouse
            string uniqueID_lf_VATCFL;

            //VAT NORMAL - "0"
            top = top + 15;

            uniqueID_lf_VATCFL = "V_CFL0";
            FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_lf_VATCFL);

            //პირობის დადება ინფუთით
            SAPbouiCOM.ChooseFromList oCFL_VAT = oForm.ChooseFromLists.Item(uniqueID_lf_VATCFL);
            SAPbouiCOM.Conditions oCons_VAT = oCFL_VAT.GetConditions();
            SAPbouiCOM.Condition oCon_VAT = oCons_VAT.Add();
            oCon_VAT.Alias = "Category";
            oCon_VAT.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCon_VAT.CondVal = "I"; //Active Account
            oCFL_VAT.SetConditions(oCons_VAT);

            formItems = new Dictionary<string, object>();
            itemName = "RSVAT0";
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", 13);
            formItems.Add("Width", 207);
            formItems.Add("Top", top);
            formItems.Add("Height", 15);
            formItems.Add("Caption", BDOSResources.getTranslate("VATTypeNormal"));
            formItems.Add("UID", itemName);
            formItems.Add("FromPane", 12);
            formItems.Add("ToPane", 12);
            formItems.Add("LinkTo", "BDO_RSVAT0");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "BDO_RSVAT0";
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OADM");
            formItems.Add("Alias", "U_BDO_RSVAT0");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", 230);
            formItems.Add("Width", 250);
            formItems.Add("Top", top + 1);
            formItems.Add("Height", 14);
            formItems.Add("UID", itemName);
            formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
            formItems.Add("DisplayDesc", true);
            formItems.Add("FromPane", 12);
            formItems.Add("ToPane", 12);
            formItems.Add("Description", BDOSResources.getTranslate("VATTypeNormal"));
            formItems.Add("ChooseFromListUID", uniqueID_lf_VATCFL);
            formItems.Add("ChooseFromListAlias", "Code");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }
            //VAT NORMAL - 0

            //VAT ნულოვანი - "1"
            top = top + 15;

            uniqueID_lf_VATCFL = "V_CFL1";
            FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_lf_VATCFL);

            //პირობის დადება ინფუთით
            SAPbouiCOM.ChooseFromList oCFL_VAT1 = oForm.ChooseFromLists.Item(uniqueID_lf_VATCFL);
            SAPbouiCOM.Conditions oCons_VAT1 = oCFL_VAT1.GetConditions();
            SAPbouiCOM.Condition oCon_VAT1 = oCons_VAT1.Add();
            oCon_VAT1.Alias = "Category";
            oCon_VAT1.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCon_VAT1.CondVal = "I"; //Active Account
            oCFL_VAT1.SetConditions(oCons_VAT1);

            formItems = new Dictionary<string, object>();
            itemName = "RSVAT1";
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", 13);
            formItems.Add("Width", 207);
            formItems.Add("Top", top);
            formItems.Add("Height", 15);
            formItems.Add("Caption", BDOSResources.getTranslate("VATType0"));
            formItems.Add("UID", itemName);
            formItems.Add("FromPane", 12);
            formItems.Add("ToPane", 12);
            formItems.Add("LinkTo", "BDO_RSVAT1");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "BDO_RSVAT1";
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OADM");
            formItems.Add("Alias", "U_BDO_RSVAT1");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", 230);
            formItems.Add("Width", 250);
            formItems.Add("Top", top + 1);
            formItems.Add("Height", 14);
            formItems.Add("UID", itemName);
            formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
            formItems.Add("DisplayDesc", true);
            formItems.Add("FromPane", 12);
            formItems.Add("ToPane", 12);
            formItems.Add("Description", BDOSResources.getTranslate("VATType0"));
            formItems.Add("ChooseFromListUID", uniqueID_lf_VATCFL);
            formItems.Add("ChooseFromListAlias", "Code");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }
            //VAT ნულოვანი - 1

            //VAT არ იბეგრება "2"
            top = top + 15;

            uniqueID_lf_VATCFL = "V_CFL2";
            FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_lf_VATCFL);

            //პირობის დადება ინფუთით
            SAPbouiCOM.ChooseFromList oCFL_VAT2 = oForm.ChooseFromLists.Item(uniqueID_lf_VATCFL);
            SAPbouiCOM.Conditions oCons_VAT2 = oCFL_VAT2.GetConditions();
            SAPbouiCOM.Condition oCon_VAT2 = oCons_VAT2.Add();
            oCon_VAT2.Alias = "Category";
            oCon_VAT2.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCon_VAT2.CondVal = "I"; //Active Account
            oCFL_VAT2.SetConditions(oCons_VAT2);

            formItems = new Dictionary<string, object>();
            itemName = "RSVAT2";
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", 13);
            formItems.Add("Width", 207);
            formItems.Add("Top", top);
            formItems.Add("Height", 15);
            formItems.Add("Caption", BDOSResources.getTranslate("VATTypeNone"));
            formItems.Add("UID", itemName);
            formItems.Add("FromPane", 12);
            formItems.Add("ToPane", 12);
            formItems.Add("LinkTo", "BDO_RSVAT2");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "BDO_RSVAT2";
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OADM");
            formItems.Add("Alias", "U_BDO_RSVAT2");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", 230);
            formItems.Add("Width", 250);
            formItems.Add("Top", top + 1);
            formItems.Add("Height", 14);
            formItems.Add("UID", itemName);
            formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
            formItems.Add("DisplayDesc", true);
            formItems.Add("FromPane", 12);
            formItems.Add("ToPane", 12);
            formItems.Add("Description", BDOSResources.getTranslate("VATTypeNone"));
            formItems.Add("ChooseFromListUID", uniqueID_lf_VATCFL);
            formItems.Add("ChooseFromListAlias", "Code");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            ////    top = top + 30;
            //VAT არ იბეგრება "2"

            //რეზერვის ინვოისიდან ფაქტურის გამოწერა როგორც მომსახურებაზე
            top = top + 25;

            formItems = new Dictionary<string, object>();
            itemName = "BDOSResSrv";
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OADM");
            formItems.Add("Alias", "U_BDOSResSrv");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
            formItems.Add("Left", 13);
            formItems.Add("Width", 300);
            formItems.Add("Top", top);
            formItems.Add("Height", 14);
            formItems.Add("UID", itemName);
            formItems.Add("FromPane", 12);
            formItems.Add("ToPane", 12);
            formItems.Add("Description", BDOSResources.getTranslate("ReserveInvoiceAsService"));
            formItems.Add("Caption", BDOSResources.getTranslate("ReserveInvoiceAsService"));
            formItems.Add("ValOff", "N");
            formItems.Add("ValOn", "Y");
            formItems.Add("DisplayDesc", true);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }
            //რეზერვის ინვოისიდან ფაქტურის გამოწერა როგორც მომსახურებაზე

            //ზედნადების გარეშე ანგარიშ-ფაქტურის გამოწერის შეზღუდვა
            top = top + 25;

            formItems = new Dictionary<string, object>();
            itemName = "BDOSInRstr";
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OADM");
            formItems.Add("Alias", "U_BDOSInvoiceRestr");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
            formItems.Add("Left", 13);
            formItems.Add("Width", 400);
            formItems.Add("Top", top);
            formItems.Add("Height", 14);
            formItems.Add("UID", itemName);
            formItems.Add("FromPane", 12);
            formItems.Add("ToPane", 12);
            formItems.Add("Description", BDOSResources.getTranslate("RestrictTaxWithoutWaybill"));
            formItems.Add("Caption", BDOSResources.getTranslate("RestrictTaxWithoutWaybill"));
            formItems.Add("ValOff", "N");
            formItems.Add("ValOn", "Y");
            formItems.Add("DisplayDesc", true);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }
            //ზედნადების გარეშე ანგარიშ-ფაქტურის გამოწერის შეზღუდვა

            top = top + 25;

            formItems = new Dictionary<string, object>();
            itemName = "BDOSAllDeS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", 13);
            formItems.Add("Width", 207);
            formItems.Add("Top", top);
            formItems.Add("Height", 15);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("AllowableDeviation"));
            formItems.Add("FromPane", 12);
            formItems.Add("ToPane", 12);
            formItems.Add("LinkTo", "BDOSAllDeE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "BDOSAllDeE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OADM");
            formItems.Add("Alias", "U_BDOSAllDev");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", 230);
            formItems.Add("Width", 250);
            formItems.Add("Top", top + 1);
            formItems.Add("Height", 14);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("FromPane", 12);
            formItems.Add("ToPane", 12);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + 25;

            

            //სერვის მომხმარებლის სახელი (RS.GE)
            formItems = new Dictionary<string, object>();
            itemName = "SU";
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", 13);
            formItems.Add("Width", 207);
            formItems.Add("Top", top);
            formItems.Add("Height", 15);
            formItems.Add("Caption", BDOSResources.getTranslate("ServiceUser"));
            formItems.Add("UID", itemName);
            formItems.Add("FromPane", 12);
            formItems.Add("ToPane", 12);
            formItems.Add("LinkTo", "BDO_SU");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "BDO_SU";
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OADM");
            formItems.Add("Alias", "U_BDO_SU");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", 230);
            formItems.Add("Width", 250);
            formItems.Add("Top", top + 1);
            formItems.Add("Height", 14);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("FromPane", 12);
            formItems.Add("ToPane", 12);
            formItems.Add("Description", BDOSResources.getTranslate("ServiceUser"));

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + 15;

            //სერვის მომხმარებლის პაროლი (RS.GE) ორგანიზაციის მიხედვით
            formItems = new Dictionary<string, object>();
            itemName = "SP";
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", 13);
            formItems.Add("Width", 207);
            formItems.Add("Top", top);
            formItems.Add("Height", 15);
            formItems.Add("Caption", BDOSResources.getTranslate("ServiceUserPassword"));
            formItems.Add("UID", itemName);
            formItems.Add("FromPane", 12);
            formItems.Add("ToPane", 12);
            formItems.Add("LinkTo", "BDO_SP");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "BDO_SP";
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OADM");
            formItems.Add("Alias", "U_BDO_SP");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", 230);
            formItems.Add("Width", 250);
            formItems.Add("Top", top + 1);
            formItems.Add("Height", 14);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("FromPane", 12);
            formItems.Add("ToPane", 12);
            formItems.Add("IsPassword", true);
            formItems.Add("Description", BDOSResources.getTranslate("ServiceUserPassword"));

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            /////////////////
            

            top = top + 15;

            //ზედნადების ავტორიზაცია
            formItems = new Dictionary<string, object>();
            itemName = "WblAut";
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", 13);
            formItems.Add("Width", 207);
            formItems.Add("Top", top);
            formItems.Add("Height", 15);
            formItems.Add("Caption", BDOSResources.getTranslate("WaybillAuthorization"));
            formItems.Add("UID", itemName);
            formItems.Add("FromPane", 12);
            formItems.Add("ToPane", 12);
            formItems.Add("LinkTo", "BDOSWblAut");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            listValidValues = new List<string>();
            listValidValues.Add("No Authorization");
            listValidValues.Add("Read Only");
            listValidValues.Add("Full");

            formItems = new Dictionary<string, object>();
            itemName = "BDOSWblAut";
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OADM");
            formItems.Add("Alias", "U_BDOSWblAut");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            formItems.Add("Left", 230);
            formItems.Add("Width", 250);
            formItems.Add("Top", top + 1);
            formItems.Add("Height", 14);
            formItems.Add("UID", itemName);
            formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
            formItems.Add("DisplayDesc", true);
            formItems.Add("FromPane", 12);
            formItems.Add("ToPane", 12);
            formItems.Add("Description", BDOSResources.getTranslate("WaybillAuthorization"));
            formItems.Add("ValidValues", listValidValues);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + 15;

            //ზედნადების ავტორიზაცია
            formItems = new Dictionary<string, object>();
            itemName = "TaxAut";
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", 13);
            formItems.Add("Width", 207);
            formItems.Add("Top", top);
            formItems.Add("Height", 15);
            formItems.Add("Caption", BDOSResources.getTranslate("TaxAuthorization"));
            formItems.Add("UID", itemName);
            formItems.Add("FromPane", 12);
            formItems.Add("ToPane", 12);
            formItems.Add("LinkTo", "BDOSTaxAut");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            listValidValues = new List<string>();
            listValidValues.Add("No Authorization");
            listValidValues.Add("Read Only");
            listValidValues.Add("Approval");

            formItems = new Dictionary<string, object>();
            itemName = "BDOSTaxAut";
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OADM");
            formItems.Add("Alias", "U_BDOSTaxAut");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            formItems.Add("Left", 230);
            formItems.Add("Width", 250);
            formItems.Add("Top", top + 1);
            formItems.Add("Height", 14);
            formItems.Add("UID", itemName);
            formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
            formItems.Add("DisplayDesc", true);
            formItems.Add("FromPane", 12);
            formItems.Add("ToPane", 12);
            formItems.Add("Description", BDOSResources.getTranslate("TaxAuthorization"));
            formItems.Add("ValidValues", listValidValues);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + 15;

            formItems = new Dictionary<string, object>();
            itemName = "BDOSDecAtt";
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OADM");
            formItems.Add("Alias", "U_BDOSDecAtt");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
            formItems.Add("Left", 13);
            formItems.Add("Width", 400);
            formItems.Add("Top", top);
            formItems.Add("Height", 14);
            formItems.Add("UID", itemName);
            formItems.Add("FromPane", 12);
            formItems.Add("ToPane", 12);
            formItems.Add("Description", BDOSResources.getTranslate("AttachTaxInvoiceOnDeclaration"));
            formItems.Add("Caption", BDOSResources.getTranslate("AttachTaxInvoiceOnDeclaration"));
            formItems.Add("ValOff", "N");
            formItems.Add("ValOn", "Y");
            formItems.Add("DisplayDesc", true);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + 15;
            /////////////////




            SAPbouiCOM.Item oItemOK = oForm.Items.Item("1");
            formItems = new Dictionary<string, object>();
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            formItems.Add("Left", 13);
            formItems.Add("Width", oItemOK.Width * 2);
            formItems.Add("Top", top + 3);
            formItems.Add("Height", oItemOK.Height);
            formItems.Add("Caption", BDOSResources.getTranslate("SavePassword"));
            formItems.Add("UID", "BDO_SetPas");
            formItems.Add("FromPane", 12);
            formItems.Add("ToPane", 12);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }


            #region Use Discount

            var topD = oForm.Items.Item("230").Top;
            var leftD = oForm.Items.Item("162").Left;
            var heightD = oForm.Items.Item("230").Height;

            formItems = new Dictionary<string, object>();
            itemName = "Discount";
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OADM");
            formItems.Add("Alias", "U_Discount");
            formItems.Add("Bound", true);
            formItems.Add("Type", BoFormItemTypes.it_CHECK_BOX);
            formItems.Add("DataType", BoDataType.dt_SHORT_TEXT);
            formItems.Add("Left", leftD);
            formItems.Add("Width", 150);
            formItems.Add("Top", topD);
            formItems.Add("Height", heightD);
            formItems.Add("UID", itemName);
            formItems.Add("FromPane", 11);
            formItems.Add("ToPane", 11);
            formItems.Add("Description", BDOSResources.getTranslate("DiscountUse"));
            formItems.Add("Caption", BDOSResources.getTranslate("DiscountUse"));
            formItems.Add("ValOff", "N");
            formItems.Add("ValOn", "Y");
            formItems.Add("DisplayDesc", true);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            #endregion

            top = top + oItemOK.Height;

            //მომხმარებლების ცხრილი
            itemName = "BDO_UsrMtx";
            formItems = new Dictionary<string, object>();
            formItems.Add("isDataSource", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
            formItems.Add("Left", 13);
            formItems.Add("Width", oForm.ClientWidth);
            formItems.Add("Top", top + 1);
            formItems.Add("Height", oForm.ClientHeight * 3 / 4-40);
            formItems.Add("UID", itemName);
            formItems.Add("FromPane", 12);
            formItems.Add("ToPane", 12);
            formItems.Add("DisplayDesc", true);
            formItems.Add("AffectsFormMode", true);
            formItems.Add("Description", BDOSResources.getTranslate("UserTable"));

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("BDO_UsrMtx").Specific));
            SAPbouiCOM.Columns oColumns = oMatrix.Columns;
            SAPbouiCOM.Column oColumn;

            oColumn = oColumns.Add("DSUserID", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oColumn.TitleObject.Caption = "ID      ";
            oColumn.Width = 40;
            oColumn.Editable = false;
            oColumn.TitleObject.Sortable = true;
            oColumn.AffectsFormMode = false;
            SAPbouiCOM.LinkedButton oLink = oColumn.ExtendedObject;
            oLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_User;

            oColumn = oColumns.Add("DSUserCode", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("UserCode");
            oColumn.Width = 40;
            oColumn.Editable = false;
            oColumn.Visible = false;
            oColumn.AffectsFormMode = false;

            oColumn = oColumns.Add("DSUserName", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("User");
            oColumn.Width = 40;
            oColumn.Editable = false;
            oColumn.AffectsFormMode = false;
            oColumn.TitleObject.Sortable = true;

            oColumn = oColumns.Add("DSSU", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("ServiceUser");
            oColumn.Width = 40;
            oColumn.Editable = true;
            oColumn.AffectsFormMode = false;
            oColumn.TitleObject.Sortable = true;


            oColumn = oColumns.Add("DSSP", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Password");
            oColumn.Width = 40;
            oColumn.Editable = true;
            oColumn.Visible = false;
            oColumn.AffectsFormMode = false;

            oColumn = oColumns.Add("WBAUT", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("WaybillAuthorization");
            oColumn.Width = 40;
            oColumn.Editable = true;
            oColumn.AffectsFormMode = true;
            oColumn.TitleObject.Sortable = true;
            oColumn.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
            oColumn.ValidValues.Add("0", BDOSResources.getTranslate("NoAuthorization"));
            oColumn.ValidValues.Add("1", BDOSResources.getTranslate("ReadOnly"));
            oColumn.ValidValues.Add("2", BDOSResources.getTranslate("Full"));
            oColumn.DisplayDesc = true;

            oColumn = oColumns.Add("TXAUT", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("TaxAuthorization");
            oColumn.Width = 40;
            oColumn.Editable = true;
            oColumn.AffectsFormMode = true;
            oColumn.TitleObject.Sortable = true;
            oColumn.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
            oColumn.ValidValues.Add("0", BDOSResources.getTranslate("NoAuthorization"));
            oColumn.ValidValues.Add("1", BDOSResources.getTranslate("ReadOnly"));
            oColumn.ValidValues.Add("2", BDOSResources.getTranslate("Approval"));
            oColumn.DisplayDesc = true;

            oColumn = oColumns.Add("DCAT", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("AttachTaxInvoiceOnDeclaration");
            oColumn.Width = 40;
            oColumn.Editable = true;
            oColumn.AffectsFormMode = true;
            oColumn.TitleObject.Sortable = true;
            oColumn.TitleObject.Sortable = true;
            oColumn.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
            oColumn.ValidValues.Add("N", BDOSResources.getTranslate("No"));
            oColumn.ValidValues.Add("Y", BDOSResources.getTranslate("Yes"));
            oColumn.DisplayDesc = true;



            SAPbouiCOM.DBDataSource oDBDataSource;
            oDBDataSource = oForm.DataSources.DBDataSources.Add("OUSR");

            oColumn = oColumns.Item("DSUserID");
            oColumn.DataBind.SetBound(true, "OUSR", "USERID");
            oColumn = oColumns.Item("DSUserCode");
            oColumn.DataBind.SetBound(true, "OUSR", "USER_CODE");
            oColumn = oColumns.Item("DSUserName");
            oColumn.DataBind.SetBound(true, "OUSR", "U_NAME");
            oColumn = oColumns.Item("DSSU");
            oColumn.DataBind.SetBound(true, "OUSR", "U_BDO_SU");
            oColumn = oColumns.Item("DSSP");
            oColumn.DataBind.SetBound(true, "OUSR", "U_BDO_SP");

            oColumn = oColumns.Item("WBAUT");
            oColumn.DataBind.SetBound(true, "OUSR", "U_BDOSWblAut");
            oColumn = oColumns.Item("TXAUT");
            oColumn.DataBind.SetBound(true, "OUSR", "U_BDOSTaxAut");
            oColumn = oColumns.Item("DCAT");
            oColumn.DataBind.SetBound(true, "OUSR", "U_BDOSDecAtt");




            oMatrix.Clear();
            oDBDataSource.Query();
            oMatrix.LoadFromDataSource();
            oMatrix.AutoResizeColumns();

            oColumn = oMatrix.Columns.Item("DSUserName");
            oColumn.TitleObject.Sortable = true;
            oColumn.TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending);
            oColumn.AffectsFormMode = false;

            int wsdlMTRWidth = oForm.ClientWidth;
            top = oForm.Items.Item("1470002220").Top;
            int left = oForm.Items.Item("54").Left;
            //int height = 15;
            top = top + 35;


            formItems = new Dictionary<string, object>();
            itemName = "BDOSIBKs";
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left);
            formItems.Add("Width", 150);
            formItems.Add("Top", FirstItemHeight);
            formItems.Add("Height", 15);
            formItems.Add("Caption", BDOSResources.getTranslate("PayrolKeywords"));
            formItems.Add("UID", itemName);
            formItems.Add("FromPane", 200);
            formItems.Add("ToPane", 200);
            formItems.Add("LinkTo", "BDOSIBKW");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "BDOSIBKW";
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OADM");
            formItems.Add("Alias", "U_BDOSIBKW");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left + 150 + 5);
            formItems.Add("Width", 280);
            formItems.Add("Top", FirstItemHeight);
            formItems.Add("Height", 14);
            formItems.Add("UID", itemName);
            formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
            formItems.Add("DisplayDesc", true);
            formItems.Add("FromPane", 200);
            formItems.Add("ToPane", 200);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            if (UDO.UserDefinedFieldExists("OVPM", "BDOSBdgCf"))
            {
                formItems = new Dictionary<string, object>();
                itemName = "BDOSDevCmp";
                formItems.Add("isDataSource", true);
                formItems.Add("DataSource", "DBDataSources");
                formItems.Add("TableName", "OADM");
                formItems.Add("Alias", "U_BDOSDevCmp");
                formItems.Add("Bound", true);
                formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
                formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                formItems.Add("Left", left);
                formItems.Add("Width", 300);
                formItems.Add("Top", FirstItemHeight + ItemDistanceVertical * 1);
                formItems.Add("Height", 14);
                formItems.Add("UID", itemName);
                formItems.Add("FromPane", 200);
                formItems.Add("ToPane", 200);
                formItems.Add("Description", BDOSResources.getTranslate("IsDevelopmentCompany"));
                formItems.Add("Caption", BDOSResources.getTranslate("IsDevelopmentCompany"));
                formItems.Add("ValOff", "N");
                formItems.Add("ValOn", "Y");
                formItems.Add("DisplayDesc", true);

                FormsB1.createFormItem(oForm, formItems, out errorText);
                if (errorText != null)
                {
                    return;
                }

                /*formItems = new Dictionary<string, object>();
                itemName = "BDOSDefCfS"; //10 characters
                formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                formItems.Add("Left", left);
                formItems.Add("Width", 150);
                formItems.Add("Top", FirstItemHeight + ItemDistanceVertical * 2);
                formItems.Add("Height", 14);
                formItems.Add("FromPane", 200);
                formItems.Add("ToPane", 200);
                formItems.Add("UID", itemName);
                formItems.Add("Caption", BDOSResources.getTranslate("BudgetCashFlow"));
                formItems.Add("LinkTo", "BDOSBdgCfE");

                FormsB1.createFormItem(oForm, formItems, out errorText);
                if (errorText != null)
                {
                    return;
                }

                multiSelection = false;
                objectType = "UDO_F_BDOSBUCFW_D";
                string uniqueID_lf_Budg_CFL = "Budg_CFL";
                FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_lf_Budg_CFL);

                formItems = new Dictionary<string, object>();
                itemName = "BDOSDefCfE"; //10 characters
                formItems.Add("isDataSource", true);
                formItems.Add("DataSource", "DBDataSources");
                formItems.Add("TableName", "OADM");
                formItems.Add("Alias", "U_BDOSDefCf");
                formItems.Add("Bound", true);
                formItems.Add("FromPane", 200);
                formItems.Add("ToPane", 200);
                formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                formItems.Add("Left", left + 150 + 5);
                formItems.Add("Width", 40);
                formItems.Add("Top", FirstItemHeight + ItemDistanceVertical * 2);
                formItems.Add("Height", 14);
                formItems.Add("UID", itemName);
                formItems.Add("DisplayDesc", true);
                formItems.Add("ChooseFromListUID", uniqueID_lf_Budg_CFL);
                formItems.Add("ChooseFromListAlias", "Code");

                FormsB1.createFormItem(oForm, formItems, out errorText);
                if (errorText != null)
                {
                    return;
                }
            
                //BdgCfNameUDS = oForm.DataSources.UserDataSources.Add("BDOSDefCfN", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
            
                formItems = new Dictionary<string, object>();
                itemName = "BDOSDefCfN"; //10 characters
                formItems.Add("isDataSource", true);
                formItems.Add("DataSource", "UserDataSources");
                formItems.Add("TableName", "");
                formItems.Add("Length", 200);
                formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                formItems.Add("Alias", "BDOSDefCfN");
                formItems.Add("Bound", true);
                formItems.Add("FromPane", 200);
                formItems.Add("ToPane", 200);
                formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                formItems.Add("Left", left + 150 + 5 + 20 + 5 + 20);
                formItems.Add("Width", 240);
                formItems.Add("Top", FirstItemHeight + ItemDistanceVertical * 2);
                formItems.Add("Height", 14);
                formItems.Add("UID", itemName);
                formItems.Add("DisplayDesc", true);
                formItems.Add("ChooseFromListUID", uniqueID_lf_Budg_CFL);
                formItems.Add("ChooseFromListAlias", "Code");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
                }*/

            }

            top = top + 20;

            formItems = new Dictionary<string, object>();
            itemName = "wsdlMTR"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
            formItems.Add("Left", left);
            formItems.Add("Width", wsdlMTRWidth);
            formItems.Add("Top", FirstItemHeight + 3 * ItemDistanceVertical);
            formItems.Add("Height", 100);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("AffectsFormMode", true);
            formItems.Add("FromPane", 200);
            formItems.Add("ToPane", 200);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("wsdlMTR").Specific));
            oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
            oColumns = oMatrix.Columns;

            oColumn = oColumns.Add("Code", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "  ";
            oColumn.Editable = false;
            oColumn.Width = 20;

            wsdlMTRWidth = wsdlMTRWidth - 20;

            oColumn = oColumns.Add("U_program", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Program");
            oColumn.Width = wsdlMTRWidth / 2;
            oColumn.Editable = true;
            oColumn.Visible = true;

            oColumn = oColumns.Add("U_mode", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Mode");
            oColumn.Width = wsdlMTRWidth / 2;
            oColumn.Editable = true;
            oColumn.Visible = true;

            oDBDataSource = oForm.DataSources.DBDataSources.Add("@BDO_INTB");

            oColumn = oColumns.Item("Code");
            oColumn.DataBind.SetBound(true, "@BDO_INTB", "Code");

            oColumn = oColumns.Item("U_program");
            oColumn.DataBind.SetBound(true, "@BDO_INTB", "U_program");
            oColumn.DisplayDesc = true;
            oColumn.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;

            oColumn = oColumns.Item("U_mode");
            oColumn.DataBind.SetBound(true, "@BDO_INTB", "U_mode");
            oColumn.DisplayDesc = true;
            oColumn.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
            SAPbobsCOM.UserTable oUserTable = Program.oCompany.UserTables.Item("BDO_INTB");
            SAPbobsCOM.ValidValues validValues = oUserTable.UserFields.Fields.Item("U_mode").ValidValues;
            for (int i = 0; i < validValues.Count; i++)
            {
                string value = validValues.Item(i).Value;
                oColumn.ValidValues.Add(value, BDOSResources.getTranslate(value));
            }

            oMatrix.Clear();
            oDBDataSource.Query();
            oMatrix.LoadFromDataSource();

            //მოგების გადასახადი (ჩანართი)
            int pane = 13;

            oFolder = oForm.Items.Item("BDO_RSGE");
            formItems = new Dictionary<string, object>();
            itemName = "BDO_TAXG";
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_FOLDER);
            formItems.Add("Left", oFolder.Left + oFolder.Width);
            formItems.Add("Width", 50);
            formItems.Add("Top", oFolder.Top);
            formItems.Add("Height", oFolder.Height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Taxes"));
            formItems.Add("Pane", pane);
            formItems.Add("ValOn", "0");
            formItems.Add("ValOff", itemName);
            formItems.Add("GroupWith", "BDO_RSGE");
            formItems.Add("AffectsFormMode", false);
            formItems.Add("Description", BDOSResources.getTranslate("Taxes"));

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            int topTAX = 37;
            int left_s = 13;
            int left_e = 230;
            int width_s = 207;
            int width_e = 250;

            //მოგების გადასახადის დაბეგვრის სისტემა
            formItems = new Dictionary<string, object>();
            itemName = "TaxTyp";
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", topTAX);
            formItems.Add("Height", 15);
            formItems.Add("Caption", BDOSResources.getTranslate("ProfitTaxSystem"));
            formItems.Add("UID", itemName);
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);
            formItems.Add("LinkTo", "BDO_TaxTyp");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "BDO_TaxTyp";
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OADM");
            formItems.Add("Alias", "U_BDO_TaxTyp");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", topTAX + 1);
            formItems.Add("Height", 14);
            formItems.Add("UID", itemName);
            formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
            formItems.Add("DisplayDesc", true);
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);
            formItems.Add("Description", BDOSResources.getTranslate("ProfitTaxSystem"));

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            topTAX = topTAX + 15;

            formItems = new Dictionary<string, object>();
            itemName = "TaxRate";
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", topTAX);
            formItems.Add("Height", 15);
            formItems.Add("Caption", BDOSResources.getTranslate("Rate") + " %");
            formItems.Add("UID", itemName);
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);
            formItems.Add("LinkTo", "BDO_PrTxRt");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "BDO_PrTxRt";
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OADM");
            formItems.Add("Alias", "U_BDO_PrTxRt");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", topTAX + 1);
            formItems.Add("Height", 14);
            formItems.Add("UID", itemName);
            formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
            formItems.Add("DisplayDesc", true);
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);
            formItems.Add("Description", BDOSResources.getTranslate("Rate"));

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            topTAX = topTAX + 15;

            string objectTypeAcct = "1";
            string uniqueID_lf_AcctCFL_Cap;
            string uniqueID_lf_AcctCFL_Tax;

            formItems = new Dictionary<string, object>();
            itemName = "CapAcc";
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", topTAX);
            formItems.Add("Height", 15);
            formItems.Add("Caption", BDOSResources.getTranslate("Capital") + " " + BDOSResources.getTranslate("Account"));
            formItems.Add("UID", itemName);
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);
            formItems.Add("LinkTo", "BDO_CapAcc");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            uniqueID_lf_AcctCFL_Cap = "Acct_CFL_Cap";
            FormsB1.addChooseFromList(oForm, multiSelection, objectTypeAcct, uniqueID_lf_AcctCFL_Cap);

            //პირობის დადება ანგარიშის არჩევის სიაზე
            SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(uniqueID_lf_AcctCFL_Cap);
            SAPbouiCOM.Conditions oCons = oCFL.GetConditions();
            SAPbouiCOM.Condition oCon = oCons.Add();
            oCon.Alias = "Postable";
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCon.CondVal = "Y"; //Active Account
            oCFL.SetConditions(oCons);

            formItems = new Dictionary<string, object>();
            itemName = "BDO_CapAcc";
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OADM");
            formItems.Add("Alias", "U_BDO_CapAcc");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", topTAX + 1);
            formItems.Add("Height", 14);
            formItems.Add("UID", itemName);
            formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
            formItems.Add("DisplayDesc", true);
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);
            formItems.Add("Description", BDOSResources.getTranslate("Capital") + " " + BDOSResources.getTranslate("Account"));
            formItems.Add("ChooseFromListUID", uniqueID_lf_AcctCFL_Cap);
            formItems.Add("ChooseFromListAlias", "AcctCode");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            //golden errow
            formItems = new Dictionary<string, object>();
            itemName = "CapAccLB"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            formItems.Add("Left", left_e - 20);
            formItems.Add("Top", topTAX + 1);
            formItems.Add("Height", 14);
            formItems.Add("UID", itemName);
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);
            formItems.Add("LinkTo", "BDO_CapAcc");
            formItems.Add("LinkedObjectType", objectTypeAcct);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            topTAX = topTAX + 15;

            formItems = new Dictionary<string, object>();
            itemName = "TaxAcc";
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", topTAX);
            formItems.Add("Height", 15);
            formItems.Add("Caption", BDOSResources.getTranslate("ProfitTaxPayable") + " " + BDOSResources.getTranslate("Account"));
            formItems.Add("UID", itemName);
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);
            formItems.Add("LinkTo", "BDO_TaxAcc");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            uniqueID_lf_AcctCFL_Tax = "Acct_CFL_Tax";
            FormsB1.addChooseFromList(oForm, multiSelection, objectTypeAcct, uniqueID_lf_AcctCFL_Tax);
            //პირობის დადება ანგარიშის არჩევის სიაზე

            SAPbouiCOM.ChooseFromList oCFL_Tax = oForm.ChooseFromLists.Item(uniqueID_lf_AcctCFL_Tax);
            SAPbouiCOM.Conditions oCons_Tax = oCFL_Tax.GetConditions();
            SAPbouiCOM.Condition oCon_Tax = oCons_Tax.Add();
            oCon_Tax.Alias = "Postable";
            oCon_Tax.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCon_Tax.CondVal = "Y"; //Active Account
            oCFL_Tax.SetConditions(oCons_Tax);

            formItems = new Dictionary<string, object>();
            itemName = "BDO_TaxAcc";
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OADM");
            formItems.Add("Alias", "U_BDO_TaxAcc");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", topTAX);
            formItems.Add("Height", 14);
            formItems.Add("UID", itemName);
            formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
            formItems.Add("DisplayDesc", true);
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);
            formItems.Add("Description", BDOSResources.getTranslate("ProfitTaxPayable") + " " + BDOSResources.getTranslate("Account"));
            formItems.Add("ChooseFromListUID", uniqueID_lf_AcctCFL_Tax);
            formItems.Add("ChooseFromListAlias", "AcctCode");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            //golden errow
            formItems = new Dictionary<string, object>();
            itemName = "TaxAccLB"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            formItems.Add("Left", left_e - 20);
            formItems.Add("Top", topTAX);
            formItems.Add("Height", 14);
            formItems.Add("UID", itemName);
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);
            formItems.Add("LinkTo", "BDO_TaxAcc");
            formItems.Add("LinkedObjectType", objectTypeAcct);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            topTAX = topTAX + 25;

            formItems = new Dictionary<string, object>();
            itemName = "DfaultPrBs"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s - 5);
            formItems.Add("Width", width_s + 50);
            formItems.Add("Top", topTAX);
            formItems.Add("Height", 14);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("DefaultTaxableObjects"));
            formItems.Add("TextStyle", 4);
            formItems.Add("FontSize", 12);
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            topTAX = topTAX + 25;

            //PrBase Down Payment Request
            formItems = new Dictionary<string, object>();
            itemName = "PrBaseSDR"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_e);
            formItems.Add("Top", topTAX);
            formItems.Add("Height", 14);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("DefaulfTaxableObjectDownPaymentToProfitTaxExemptEntities"));
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);
            formItems.Add("LinkTo", "PrBaseEDR");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            objectType = "UDO_F_BDO_PTBS_D";
            string uniqueID_CFL_DR = "CFL_ProfitBase_DR";
            FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_CFL_DR);

            formItems = new Dictionary<string, object>();
            itemName = "PrBaseEDR"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OADM");
            formItems.Add("Alias", "U_BDO_prBsDR");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", 30);
            formItems.Add("Top", topTAX);
            formItems.Add("Height", 14);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);
            formItems.Add("ChooseFromListUID", uniqueID_CFL_DR);
            formItems.Add("ChooseFromListAlias", "Code");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "PrBsDRDscr"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OADM");
            formItems.Add("Alias", "U_BDO_prDRDs");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e + 32);
            formItems.Add("Width", width_e - 32);
            formItems.Add("Top", topTAX);
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
            itemName = "PrBaseDRLB"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            formItems.Add("Left", left_e - 20);
            formItems.Add("Top", topTAX);
            formItems.Add("Height", 14);
            formItems.Add("UID", itemName);
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);
            formItems.Add("LinkTo", "PrBaseEDR");
            formItems.Add("LinkedObjectType", objectType);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            topTAX = topTAX + 15;
            //PrBase Goods Issue
            formItems = new Dictionary<string, object>();
            itemName = "PrBaseSGI"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_e);
            formItems.Add("Top", topTAX);
            formItems.Add("Height", 14);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("DefaulfTaxableObjectGoodsAndServicesDeliveredFreeOfCharge"));
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);
            formItems.Add("LinkTo", "PrBaseEGI");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            string uniqueID_CFL_GI = "CFL_ProfitBase_GI";
            FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_CFL_GI);

            formItems = new Dictionary<string, object>();
            itemName = "PrBaseEGI"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OADM");
            formItems.Add("Alias", "U_BDO_prBsGI");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", 30);
            formItems.Add("Top", topTAX);
            formItems.Add("Height", 14);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);
            formItems.Add("ChooseFromListUID", uniqueID_CFL_GI);
            formItems.Add("ChooseFromListAlias", "Code");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "PrBsGIDscr"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OADM");
            formItems.Add("Alias", "U_BDO_prGIDs");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e + 32);
            formItems.Add("Width", width_e - 32);
            formItems.Add("Top", topTAX);
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
            itemName = "PrBaseGILB"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            formItems.Add("Left", left_e - 20);
            formItems.Add("Top", topTAX);
            formItems.Add("Height", 14);
            formItems.Add("UID", itemName);
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);
            formItems.Add("LinkTo", "PrBaseEGI");
            formItems.Add("LinkedObjectType", objectType);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            //საპენსიო
            int height = 14;
            objectType = "178"; //SAPbouiCOM.BoLinkedObject.lf_GLAccounts, Business Partner object 
            string uniqueID_lf_WTCodeCFLCO = "WTax_CFLCO";
            FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_lf_WTCodeCFLCO);

            objectType = "178"; //SAPbouiCOM.BoLinkedObject.lf_GLAccounts, Business Partner object 
            string uniqueID_lf_WTCodeCFLPH = "WTax_CFLPH";
            FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_lf_WTCodeCFLPH);

            topTAX = topTAX + height + 5;

            formItems = new Dictionary<string, object>();
            itemName = "PnCoPS";
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", topTAX);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("CompanyPension"));
            formItems.Add("Description", BDOSResources.getTranslate("CompanyPension"));
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);
            formItems.Add("LinkTo", "BDOSPnCoP");
            formItems.Add("Visible", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "BDOSPnCoP";
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OADM");
            formItems.Add("Alias", "U_BDOSPnCoP");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", topTAX);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Visible", false);
            formItems.Add("ChooseFromListUID", uniqueID_lf_WTCodeCFLCO);
            formItems.Add("ChooseFromListAlias", "WTCode");
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            topTAX = topTAX + height + 5;

            formItems = new Dictionary<string, object>();
            itemName = "PnPhS";
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", topTAX);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("PhysicalEntityPension"));
            formItems.Add("Description", BDOSResources.getTranslate("PhysicalEntityPension"));
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);
            formItems.Add("LinkTo", "BDOSPnPh");
            formItems.Add("Visible", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "BDOSPnPh";
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OADM");
            formItems.Add("Alias", "U_BDOSPnPh");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", topTAX);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Visible", false);
            formItems.Add("ChooseFromListUID", uniqueID_lf_WTCodeCFLPH);
            formItems.Add("ChooseFromListAlias", "WTCode");
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "BDOSPnAcc";
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OADM");
            formItems.Add("Alias", "U_BDOSPnAcc");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
            formItems.Add("Left", left);
            formItems.Add("Width", 300);
            formItems.Add("Top", topTAX+15);
            formItems.Add("Height", 14);
            formItems.Add("UID", itemName);
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);
            formItems.Add("Description", BDOSResources.getTranslate("SeparationOfCompanyPensionExpenses"));
            formItems.Add("Caption", BDOSResources.getTranslate("SeparationOfCompanyPensionExpenses"));
            formItems.Add("ValOff", "N");
            formItems.Add("ValOn", "Y");
            formItems.Add("DisplayDesc", true);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            //Chief Accountant
            pane = 2;
            SAPbouiCOM.Item oItemS = oForm.Items.Item("10");
            SAPbouiCOM.Item oItemE = oForm.Items.Item("62");

            top = oItemS.Top + oItemS.Height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "ChiefAcct";
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", oItemS.Left);
            formItems.Add("Width", oItemS.Width);
            formItems.Add("Top", top);
            formItems.Add("Height", oItemS.Height);
            formItems.Add("Caption", BDOSResources.getTranslate("ChiefAccountant"));
            formItems.Add("UID", itemName);
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);
            formItems.Add("LinkTo", "BDO_TaxTyp");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "ChiefAcctS";
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OADM");
            formItems.Add("Alias", "U_BDOSChfAct");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", oItemE.Left);
            formItems.Add("Width", oItemE.Width);
            formItems.Add("Top", top);
            formItems.Add("Height", oItemE.Height);
            formItems.Add("UID", itemName);
            formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
            formItems.Add("DisplayDesc", true);
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);
            formItems.Add("Description", BDOSResources.getTranslate("ChiefAccountant"));

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            // Basic Initialisation
            pane = 11;
            SAPbouiCOM.Item oitem = oForm.Items.Item("234000008");
            top = oitem.Top + oitem.Height + 5;
            height = oitem.Height;
            left_s = oitem.Left;
            width_s = oitem.Width;
            width_e = oForm.Items.Item("120").Width;
            left_e = oForm.Items.Item("120").Left;

            formItems = new Dictionary<string, object>();
            itemName = "BDOSEnbFlM";
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OADM");
            formItems.Add("Alias", "U_BDOSEnbFlM");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);
            formItems.Add("Description", BDOSResources.getTranslate("EnableUsingFuelManagementModule"));
            formItems.Add("Caption", BDOSResources.getTranslate("EnableUsingFuelManagementModule"));
            formItems.Add("ValOff", "N");
            formItems.Add("ValOn", "Y");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "BDOSFlWhsS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s + 20);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("FuelWarehouse"));
            formItems.Add("LinkTo", "BDOSFlWhsE");
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            string objectTypeWhs = "64";
            string uniqueID_lf_FromLoc = "FlWhs_CFL";
            FormsB1.addChooseFromList(oForm, multiSelection, objectTypeWhs, uniqueID_lf_FromLoc);
           
            formItems = new Dictionary<string, object>();
            itemName = "BDOSFlWhsE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OADM");
            formItems.Add("Alias", "U_BDOSFlWhs");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("ChooseFromListUID", uniqueID_lf_FromLoc);
            formItems.Add("ChooseFromListAlias", "WhsCode");
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "BDOSWhsLB"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            formItems.Add("Left", left_e - 20);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("LinkTo", "BDOSFlWhsE");
            formItems.Add("LinkedObjectType", objectTypeWhs);
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            GC.Collect();
        }

        public static void usrTyp_OnClick(SAPbouiCOM.Form oForm)
        {
            setVisibleFormItems(oForm);
        }

        public static void taxTyp_OnClick(SAPbouiCOM.Form oForm)
        {
            setVisibleFormItems(oForm);
        }

        public static Dictionary<string, string> getRSSettings(out string errorText)
        {
            errorText = null;

            Dictionary<string, string> rsSettingsFromDB = new Dictionary<string, string>();
            rsSettingsFromDB.Add("SU", "");
            rsSettingsFromDB.Add("SP", "");
            rsSettingsFromDB.Add("ItemCode", "");
            rsSettingsFromDB.Add("UserType", "");
            rsSettingsFromDB.Add("ProtocolType", "");
            rsSettingsFromDB.Add("WaybillType", "");
            rsSettingsFromDB.Add("WBAUT", "");
            rsSettingsFromDB.Add("TXAUT", "");
            rsSettingsFromDB.Add("DCAUT", "");
            


            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string query = @"SELECT ""U_BDO_SU"", ""U_BDO_SP"", ""U_BDO_ItmCod"", ""U_BDO_UsrTyp"", ""U_BDO_PrtTyp"", ""U_BDO_WblTyp"", ""U_BDOSWblAut"", ""U_BDOSTaxAut"", ""U_BDOSDecAtt"" FROM ""OADM""";

            try
            {
                oRecordSet.DoQuery(query);

                while (!oRecordSet.EoF)
                {
                    rsSettingsFromDB["SU"] = oRecordSet.Fields.Item("U_BDO_SU").Value.ToString();
                    rsSettingsFromDB["SP"] = oRecordSet.Fields.Item("U_BDO_SP").Value.ToString();
                    rsSettingsFromDB["WBAUT"] = oRecordSet.Fields.Item("U_BDOSWblAut").Value.ToString();
                    rsSettingsFromDB["TXAUT"] = oRecordSet.Fields.Item("U_BDOSTaxAut").Value.ToString();
                    rsSettingsFromDB["DCAUT"] = oRecordSet.Fields.Item("U_BDOSDecAtt").Value.ToString();
                    
                    rsSettingsFromDB["ItemCode"] = oRecordSet.Fields.Item("U_BDO_ItmCod").Value.ToString();
                    rsSettingsFromDB["UserType"] = oRecordSet.Fields.Item("U_BDO_UsrTyp").Value.ToString();
                    rsSettingsFromDB["ProtocolType"] = oRecordSet.Fields.Item("U_BDO_PrtTyp").Value.ToString() == "0" ? "HTTP" : "HTTPS";
                    rsSettingsFromDB["WaybillType"] = oRecordSet.Fields.Item("U_BDO_WblTyp").Value.ToString();
                    oRecordSet.MoveNext();
                    break;
                }

                if (rsSettingsFromDB["UserType"] == "1") //მომხმარებლის მიხედვით
                {
                    query = @"SELECT ""U_BDO_SU"", ""U_BDO_SP"", ""U_BDOSWblAut"", ""U_BDOSTaxAut"", ""U_BDOSDecAtt"" FROM ""OUSR"" WHERE ""USER_CODE"" = '" + Program.oCompany.UserName + "'";
                    oRecordSet.DoQuery(query);

                    while (!oRecordSet.EoF)
                    {
                        rsSettingsFromDB["SU"] = oRecordSet.Fields.Item("U_BDO_SU").Value.ToString();
                        rsSettingsFromDB["SP"] = oRecordSet.Fields.Item("U_BDO_SP").Value.ToString();
                        rsSettingsFromDB["WBAUT"] = oRecordSet.Fields.Item("U_BDOSWblAut").Value.ToString();
                        rsSettingsFromDB["TXAUT"] = oRecordSet.Fields.Item("U_BDOSTaxAut").Value.ToString();
                        rsSettingsFromDB["WBAUT"] = oRecordSet.Fields.Item("U_BDOSWblAut").Value.ToString();
                        rsSettingsFromDB["TXAUT"] = oRecordSet.Fields.Item("U_BDOSTaxAut").Value.ToString();
                        rsSettingsFromDB["DCAUT"] = oRecordSet.Fields.Item("U_BDOSDecAtt").Value.ToString();

                        oRecordSet.MoveNext();
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                int errCode;
                string errMsg;

                Program.oCompany.GetLastError(out errCode, out errMsg);
                errorText = BDOSResources.getTranslate("ErrorOfRSSettings") + " " + BDOSResources.getTranslate("ErrorDescription") + " " + errMsg + "! " + BDOSResources.getTranslate("Code") + " : " + errCode + "" + BDOSResources.getTranslate("OtherInfo") + ": " + ex.Message;
                return rsSettingsFromDB;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                GC.Collect();
            }

            return rsSettingsFromDB;
        }

        public static void setVisibleFormItems(SAPbouiCOM.Form oForm)
        {
            try
            {
                if (oForm.PaneLevel == 12)
                {
                    string typeUser = oForm.DataSources.DBDataSources.Item("OADM").GetValue("U_BDO_UsrTyp", 0).Trim();

                    if (typeUser == "0") //ორგანიზაციის მიხედვით
                    {
                        oForm.Items.Item("SU").Visible = true;
                        oForm.Items.Item("BDO_SU").Visible = true;
                        oForm.Items.Item("SP").Visible = true;
                        oForm.Items.Item("BDO_SP").Visible = true;

                        oForm.Items.Item("WblAut").Visible = true;
                        oForm.Items.Item("BDOSWblAut").Visible = true;
                        oForm.Items.Item("TaxAut").Visible = true;
                        oForm.Items.Item("BDOSTaxAut").Visible = true;
                        oForm.Items.Item("BDOSDecAtt").Visible = true;
                        


                        oForm.Items.Item("BDO_UsrMtx").Visible = false;
                        oForm.Items.Item("BDO_SetPas").Visible = false;
                    }
                    else if (typeUser == "1")  //მომხმარებლის მიხედვით
                    {
                        oForm.Items.Item("SU").Visible = false;
                        oForm.Items.Item("BDO_SU").Visible = false;
                        oForm.Items.Item("SP").Visible = false;
                        oForm.Items.Item("BDO_SP").Visible = false;
                        oForm.Items.Item("WblAut").Visible = false;
                        oForm.Items.Item("BDOSWblAut").Visible = false;
                        oForm.Items.Item("BDOSDecAtt").Visible = false;
                        oForm.Items.Item("TaxAut").Visible = false;
                        oForm.Items.Item("BDOSTaxAut").Visible = false;
                        oForm.Items.Item("BDO_UsrMtx").Visible = true;
                        oForm.Items.Item("BDO_SetPas").Visible = true;
                    }
                    else
                    {
                        oForm.Items.Item("SU").Visible = false;
                        oForm.Items.Item("BDO_SU").Visible = false;
                        oForm.Items.Item("SP").Visible = false;
                        oForm.Items.Item("BDO_SP").Visible = false;
                        oForm.Items.Item("WblAut").Visible = false;
                        oForm.Items.Item("BDOSWblAut").Visible = false;
                        oForm.Items.Item("BDOSDecAtt").Visible = false;
                        oForm.Items.Item("TaxAut").Visible = false;
                        oForm.Items.Item("BDOSTaxAut").Visible = false;
                        oForm.Items.Item("BDO_UsrMtx").Visible = false;
                        oForm.Items.Item("BDO_SetPas").Visible = false;
                    }
                }

                else if (oForm.PaneLevel == 200)
                {
                    /* bool isDevelopment = oForm.DataSources.DBDataSources.Item("OADM").GetValue("U_BDOSDevCmp", 0).Trim() == "Y";

                     if (oForm.Items.Item("BDOSDefCfN").Visible)
                     {
                         oForm.Items.Item("BDOSDefCfN").Specific.Active = false;
                     }
                     oForm.Items.Item("BDOSDefCfN").Visible = isDevelopment;

                     if (oForm.Items.Item("BDOSDefCfE").Visible)
                     {
                         oForm.Items.Item("BDOSDefCfE").Specific.Active = false;
                     }
                     oForm.Items.Item("BDOSDefCfE").Visible = isDevelopment;

                     oForm.Items.Item("BDOSDefCfS").Visible = isDevelopment;

                     string bName = UDO.GetUDOFieldValueByParam("UDO_F_BDOSBUCFW_D", "Code", oForm.Items.Item("BDOSDefCfE").Specific.Value, "Name");

                     oForm.DataSources.UserDataSources.Item("BDOSDefCfN").ValueEx = bName;*/

                }

                else if (oForm.PaneLevel == 13)
                {
                    string typeTax = oForm.DataSources.DBDataSources.Item("OADM").GetValue("U_BDO_TaxTyp", 0).Trim();
                    if (typeTax == "0") // ერთობლივი შემოსავალი
                    {
                        oForm.Items.Item("CapAcc").Visible = false;
                        oForm.Items.Item("BDO_CapAcc").Visible = false;
                        oForm.Items.Item("CapAccLB").Visible = false;
                        oForm.Items.Item("TaxAcc").Visible = false;
                        oForm.Items.Item("BDO_TaxAcc").Visible = false;
                        oForm.Items.Item("TaxAccLB").Visible = false;
                        oForm.Items.Item("DfaultPrBs").Visible = false;
                        oForm.Items.Item("PrBaseSDR").Visible = false;
                        oForm.Items.Item("PrBaseDRLB").Visible = false;
                        oForm.Items.Item("PrBaseEDR").Visible = false;
                        oForm.Items.Item("PrBsDRDscr").Visible = false;
                        oForm.Items.Item("PrBaseSGI").Visible = false;
                        oForm.Items.Item("PrBaseGILB").Visible = false;
                        oForm.Items.Item("PrBaseEGI").Visible = false;
                        oForm.Items.Item("PrBsGIDscr").Visible = false;
                    }
                    else
                    {
                        oForm.Items.Item("CapAcc").Visible = true;
                        oForm.Items.Item("BDO_CapAcc").Visible = true;
                        oForm.Items.Item("CapAccLB").Visible = true;
                        oForm.Items.Item("TaxAcc").Visible = true;
                        oForm.Items.Item("BDO_TaxAcc").Visible = true;
                        oForm.Items.Item("TaxAccLB").Visible = true;
                        oForm.Items.Item("DfaultPrBs").Visible = true;
                        oForm.Items.Item("PrBaseSDR").Visible = true;
                        oForm.Items.Item("PrBaseDRLB").Visible = true;
                        oForm.Items.Item("PrBaseEDR").Visible = true;
                        oForm.Items.Item("PrBsDRDscr").Visible = true;
                        oForm.Items.Item("PrBaseSGI").Visible = true;
                        oForm.Items.Item("PrBaseGILB").Visible = true;
                        oForm.Items.Item("PrBaseEGI").Visible = true;
                        oForm.Items.Item("PrBsGIDscr").Visible = true;
                    }
                }

                else if (oForm.PaneLevel == 11)
                {
                    bool enableFuelManagment = oForm.DataSources.DBDataSources.Item("OADM").GetValue("U_BDOSEnbFlM", 0).Trim() == "Y";
                    oForm.Items.Item("BDOSFlWhsS").Visible = enableFuelManagment;
                    oForm.Items.Item("BDOSWhsLB").Visible = enableFuelManagment;
                    oForm.Items.Item("BDOSFlWhsE").Visible = enableFuelManagment;
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

        public static void setPas_OnClick(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            //ფორმის აუცილებელი თვისებები
            Dictionary<string, object> formProperties = new Dictionary<string, object>();
            formProperties.Add("UniqueID", "BDO_SetPasForm");
            formProperties.Add("BorderStyle", SAPbouiCOM.BoFormBorderStyle.fbs_Fixed);
            formProperties.Add("Title", BDOSResources.getTranslate("SaveEditPassword"));
            formProperties.Add("Left", oForm.Left + 10);
            formProperties.Add("ClientWidth", 250);
            formProperties.Add("Top", oForm.Top + 50);
            formProperties.Add("ClientHeight", 15);

            SAPbouiCOM.Form oSetPasForm;
            bool newForm;
            bool formExist = FormsB1.createForm(formProperties, out oSetPasForm, out newForm, out errorText);

            if (errorText != null)
            {
                return;
            }
            if (formExist == true)
            {
                if (newForm == true)
                {
                    Dictionary<string, object> formItems;
                    string itemName;
                    int top = 1;
                    //სერვის მომხმარებლის პაროლი (RS.GE) 
                    formItems = new Dictionary<string, object>();
                    itemName = "SP";
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", 7);
                    formItems.Add("Width", 125);
                    formItems.Add("Top", top);
                    formItems.Add("Height", 14);
                    formItems.Add("Caption", BDOSResources.getTranslate("ServiceUserPassword"));
                    formItems.Add("UID", itemName);
                    formItems.Add("LinkTo", "BDO_SP");
                    formItems.Add("RightJustified", false);

                    FormsB1.createFormItem(oSetPasForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "BDO_SP";
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("Left", 130);
                    formItems.Add("Width", 163);
                    formItems.Add("Top", top + 1);
                    formItems.Add("Height", 14);
                    formItems.Add("UID", itemName);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("IsPassword", true);
                    formItems.Add("Description", BDOSResources.getTranslate("ServiceUserPassword"));
                    formItems.Add("RightJustified", false);

                    FormsB1.createFormItem(oSetPasForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    top = top + 35;

                    itemName = "3";
                    formItems = new Dictionary<string, object>();
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", 7);
                    formItems.Add("Width", 65);
                    formItems.Add("Top", top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", "OK");

                    FormsB1.createFormItem(oSetPasForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }
                }

                oSetPasForm.Visible = true;
                oSetPasForm.Select();
            }

            GC.Collect();
        }

        public static void setPasForm_oK_OnClick(SAPbouiCOM.Form oForm, int USERID, string BDO_SU, string BDO_SP, out string errorText)
        {
            errorText = null;
            Users.updateUsersRS_Info(USERID, BDO_SU, BDO_SP, out errorText);
            oForm.Close();
        }

        public static void chooseFromList(SAPbouiCOM.Form oForm, SAPbouiCOM.ItemEvent pVal, SAPbouiCOM.IChooseFromListEvent oCFLEvento)
        {
            try
            {
                oForm.Freeze(true);

                if (!pVal.BeforeAction)
                {
                    SAPbouiCOM.DataTable oDataTable = oCFLEvento.SelectedObjects;
                    if (oDataTable != null)
                    {
                        if (oCFLEvento.ChooseFromListUID == "V_CFL1" || oCFLEvento.ChooseFromListUID == "V_CFL2" || oCFLEvento.ChooseFromListUID == "V_CFL0")
                        {
                            string VatCode = oDataTable.GetValue("Code", 0);
                            string VatFieldIndex = oCFLEvento.ChooseFromListUID.Substring(5);

                            LanguageUtils.IgnoreErrors<string>(() => oForm.Items.Item("BDO_RSVAT" + VatFieldIndex).Specific.Value = VatCode);
                        }
                        else if (oCFLEvento.ChooseFromListUID == "Acct_CFL_Cap" || oCFLEvento.ChooseFromListUID == "Acct_CFL_Tax")
                        {
                            string AcctCode = oDataTable.GetValue("AcctCode", 0);
                            string AcctFieldIndex = oCFLEvento.ChooseFromListUID.Substring(9);

                            LanguageUtils.IgnoreErrors<string>(() => oForm.Items.Item("BDO_" + AcctFieldIndex + "Acc").Specific.Value = AcctCode);
                        }
                        else if (oCFLEvento.ChooseFromListUID == "CFL_ProfitBase_DR" || oCFLEvento.ChooseFromListUID == "CFL_ProfitBase_GI")
                        {
                            string PrBsCode = oDataTable.GetValue("Code", 0);
                            string PrBsName = oDataTable.GetValue("Name", 0);

                            string PrBsFieldIndex = oCFLEvento.ChooseFromListUID.Substring(15);

                            LanguageUtils.IgnoreErrors<string>(() => oForm.Items.Item("PrBaseE").Specific.Value = PrBsCode);
                            LanguageUtils.IgnoreErrors<string>(() => oForm.Items.Item("PrBs" + PrBsFieldIndex + "Dscr").Specific.Value = PrBsName);
                        }
                        else if (oCFLEvento.ChooseFromListUID == "WTax_CFLCO" || oCFLEvento.ChooseFromListUID == "WTax_CFLPH")
                        {
                            string WTCode = oDataTable.GetValue("WTCode", 0);

                            if (oCFLEvento.ChooseFromListUID == "WTax_CFLCO")
                                LanguageUtils.IgnoreErrors<string>(() => oForm.Items.Item("BDOSPnCoP").Specific.Value = WTCode);
                            else
                                LanguageUtils.IgnoreErrors<string>(() => oForm.Items.Item("BDOSPnPh").Specific.Value = WTCode);
                        }
                        else if (oCFLEvento.ChooseFromListUID == "FlWhs_CFL")
                        {
                            string whsCode = oDataTable.GetValue("WhsCode", 0);
                            LanguageUtils.IgnoreErrors<string>(() => oForm.Items.Item("BDOSFlWhsE").Specific.Value = whsCode);
                        }
                        if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                        }
                    }
                }
                else
                {
                    if (oCFLEvento.ChooseFromListUID == "FlWhs_CFL")
                    {
                        SAPbouiCOM.ChooseFromList oCFLWhs = oForm.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID);
                        SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();

                        SAPbouiCOM.Condition oCon = oCons.Add();
                        oCon.Alias = "U_BDOSWhType";
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = "Fuel";
                        oCFLWhs.SetConditions(oCons);
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

        public static void update(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            oForm.Freeze(true);
            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("wsdlMTR").Specific));
            try
            {
                //oMatrix.FlushToDataSource();

                ///SAPbouiCOM.DBDataSource oDBDataSource = oForm.DataSources.DBDataSources.Item("@BDO_INTB");
                for (int i = 1; i <= oMatrix.RowCount; i++)
                {
                    string code = oMatrix.Columns.Item("Code").Cells.Item(i).Specific.Value;
                    string program = oMatrix.Columns.Item("U_program").Cells.Item(i).Specific.Value;
                    string mode = oMatrix.Columns.Item("U_mode").Cells.Item(i).Specific.Value;
                    string wsdl = null;
                    string id = "";

                    string url = "";
                    int port = 0;
                    int returnCode;

                    if (program == "TBC" || program == "BOG")
                    {
                        if (mode == "real")
                        {
                            if (program == "TBC")
                            {
                                wsdl = "https://dbi.tbconline.ge/dbi/dbiService";
                            }
                            else
                            {
                                wsdl = "https://api.businessonline.ge/api/";
                                url = "https://businessonline.ge";
                                port = 0;
                                id = "d7313ff8-52b6-450f-bf5b-2fd9d98702ca";
                            }
                        }
                        else if (mode == "test")
                        {
                            if (program == "TBC")
                            {
                                wsdl = "https://test.tbconline.ge/dbi/dbiService"; //"test.tbconline.ge";
                            }
                            else
                            {
                                wsdl = "https://cib2-web-dev.bog.ge/api/"; //91.209.131.231
                                url = "https://cib2-web-dev.bog.ge"; //91.209.131.231
                                port = 8090;
                                id = "cbdab9e8-b834-474c-8b82-c56856fc3baf";
                            }
                        }
                        else if (mode == "realNew" && program == "BOG")
                        {
                            wsdl = "https://api.businessonline.ge/api/";
                            url = "https://account.bog.ge";
                            port = 0;
                            id = "d7313ff8-52b6-450f-bf5b-2fd9d98702ca";
                        }
                        else if(mode == "testNew" && program == "BOG")
                        {
                            wsdl = "https://cib-api-staging.bog.ge/api/";
                            url = "https://account-test.bog.ge";
                            port = 0;
                            id = "b2f8b285-ea48-40a7-b64a-443f7104a0ec";
                        }
                        else
                        {
                            wsdl = null;
                            errorText = BDOSResources.getTranslate("ModeIsNotSelected") + "! " + program;
                            oForm.Freeze(false);
                            return;
                        }

                        SAPbobsCOM.UserTable oUserTable = Program.oCompany.UserTables.Item("BDO_INTB");

                        if (oUserTable.GetByKey(code) == false)
                        {
                            oUserTable.UserFields.Fields.Item("U_program").Value = program;
                            oUserTable.UserFields.Fields.Item("U_mode").Value = mode;
                            oUserTable.UserFields.Fields.Item("U_WSDL").Value = wsdl;
                            oUserTable.UserFields.Fields.Item("U_ID").Value = id;
                            oUserTable.UserFields.Fields.Item("U_URL").Value = url;
                            oUserTable.UserFields.Fields.Item("U_port").Value = port;
                            returnCode = oUserTable.Add();
                        }
                        else
                        {
                            oUserTable.UserFields.Fields.Item("U_program").Value = program;
                            oUserTable.UserFields.Fields.Item("U_mode").Value = mode;
                            oUserTable.UserFields.Fields.Item("U_WSDL").Value = wsdl;
                            oUserTable.UserFields.Fields.Item("U_ID").Value = id;
                            oUserTable.UserFields.Fields.Item("U_URL").Value = url;
                            oUserTable.UserFields.Fields.Item("U_port").Value = port;
                            returnCode = oUserTable.Update();
                        }

                        if (returnCode != 0)
                        {
                            int errCode;
                            string errMsg;

                            Program.oCompany.GetLastError(out errCode, out errMsg);
                            errorText = "Error description : " + errMsg + "! Code : " + errCode;
                            oForm.Freeze(false);
                            return;
                        }
                    }
                }

                //oMatrix.LoadFromDataSource();
                oForm.Freeze(false);
            }
            catch (Exception ex)
            {
                oForm.Freeze(false);
                errorText = ex.Message;
                return;
            }
        }

        public static void updateUsers(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            oForm.Freeze(true);
            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("BDO_UsrMtx").Specific));
            CommonFunctions.StartTransaction();

            try
            {
                SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                
                for (int i = 1; i <= oMatrix.RowCount; i++)
                {
                    string WBAUT = oMatrix.Columns.Item("WBAUT").Cells.Item(i).Specific.Value;
                    string TAXAUT = oMatrix.Columns.Item("TXAUT").Cells.Item(i).Specific.Value;
                    string DECAUT = oMatrix.Columns.Item("DCAT").Cells.Item(i).Specific.Value;
                    string USERID = oMatrix.Columns.Item("DSUserID").Cells.Item(i).Specific.Value;
    
                    string updateQuery = @"UPDATE ""OUSR""
                                            SET ""U_BDOSWblAut"" = N'" + WBAUT + @"',
                                            ""U_BDOSTaxAut"" = N'" + TAXAUT + @"',
                                            ""U_BDOSDecAtt"" = N'" + DECAUT + @"'
                                        WHERE ""OUSR"".""USERID"" = N'" + USERID + "'";

                    oRecordSet.DoQuery(updateQuery);
                }

                CommonFunctions.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);

                oForm.Freeze(false);
            }
            catch (Exception ex)
            {
                CommonFunctions.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                oForm.Freeze(false);
                errorText = ex.Message;
                return;
            }
        }

        public static void uiApp_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(BusinessObjectInfo.FormTypeEx, Program.currentFormCount);

            if (oForm.TypeEx == "136")
            {
                if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD & BusinessObjectInfo.BeforeAction == true)
                {
                    return;
                }

                if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE & BusinessObjectInfo.BeforeAction == true)
                {
                    update(oForm, out errorText);
                    if (errorText != null)
                    {
                        Program.uiApp.MessageBox(errorText);
                    }

                    string typeUser = oForm.DataSources.DBDataSources.Item("OADM").GetValue("U_BDO_UsrTyp", 0).Trim();
                    if (typeUser == "1") //თანამშრომლების მიხედვით
                    {
                        updateUsers(oForm, out errorText);
                        if (errorText != null)
                        {
                            Program.uiApp.MessageBox(errorText);
                        }
                    }

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


                //----------------------------->Company details-ზე პაროლის დანიშვნის ფორმა<-----------------------------
                if (pVal.FormUID == "BDO_SetPasForm")
                {
                    if (pVal.ItemUID == "3" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK & pVal.BeforeAction == false)
                    {
                        string BDO_SP = oForm.Items.Item("BDO_SP").Specific.value;
                        setPasForm_oK_OnClick(oForm, Program.USERID, Program.BDO_SU, BDO_SP, out errorText);
                        if (errorText == null)
                        {
                            Program.uiApp.StatusBar.SetText(BDOSResources.getTranslate("OperationCompletedSuccessfully"), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                        }
                        else
                        {
                            Program.uiApp.StatusBar.SetText(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        }
                        Program.USERID = 0;
                        Program.BDO_SU = null;
                    }
                }//----------------------------->Company details-ზე პაროლის დანიშვნის ფორმა<-----------------------------

                else if (pVal.FormTypeEx == "136")
                {
                    oForm.Freeze(true);

                    if (pVal.ItemUID != "2")
                    {
                        if ((pVal.ItemUID == "BDO_RSVAT0" || pVal.ItemUID == "BDO_RSVAT1" || pVal.ItemUID == "BDO_RSVAT2") & pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                        {
                            SAPbouiCOM.ChooseFromListEvent oCFLEvento = (SAPbouiCOM.ChooseFromListEvent)pVal;
                            chooseFromList(oForm, pVal, oCFLEvento);
                        }

                        if (pVal.ItemUID == "BDOSDevCmp" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED & pVal.BeforeAction == true)
                        {
                            setVisibleFormItems(oForm);
                        }

                        if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD & pVal.BeforeAction == true)
                        {
                            createFormItems(oForm, out errorText);
                            setVisibleFormItems(oForm);
                        }

                        if (pVal.ItemUID == "BDO_INBNK" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED & pVal.BeforeAction == true)
                        {
                            oForm.PaneLevel = 200;
                            setVisibleFormItems(oForm);
                        }

                        if (pVal.ItemUID == "BDO_RSGE" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED & pVal.BeforeAction == true)
                        {
                            oForm.PaneLevel = 12;
                            setVisibleFormItems(oForm);
                        }

                        if (pVal.ItemUID == "BDO_UsrTyp" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_COMBO_SELECT & pVal.BeforeAction == false)
                        {
                            usrTyp_OnClick(oForm);
                        }

                        if (pVal.ItemUID == "BDO_SetPas" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK & pVal.BeforeAction == false)
                        {
                            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("BDO_UsrMtx").Specific));
                            SAPbouiCOM.CellPosition cellPos = oMatrix.GetCellFocus();
                            if (cellPos == null)
                            {
                                Program.uiApp.MessageBox(BDOSResources.getTranslate("ForPasswordSetChooseServiceUser"));
                                oForm.Freeze(false);
                                return;
                            }
                            SAPbouiCOM.EditText oEdit = oMatrix.Columns.Item(cellPos.ColumnIndex).Cells.Item(cellPos.rowIndex).Specific;
                            Program.BDO_SU = oEdit.Value;
                            Program.USERID = Convert.ToInt32(oMatrix.Columns.Item("DSUserID").Cells.Item(cellPos.rowIndex).Specific.value);
                            setPas_OnClick(oForm, out errorText);
                        }
                        if (pVal.ItemUID == "BDO_TAXG" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED & pVal.BeforeAction)
                        {
                            oForm.PaneLevel = 13;
                            setVisibleFormItems(oForm);
                        }

                        if ((pVal.ItemUID == "36" || pVal.ItemUID == "BDOSEnbFlM") && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.BeforeAction)
                        {
                            setVisibleFormItems(oForm);
                        }

                        if ((pVal.ItemUID == "BDO_CapAcc" || pVal.ItemUID == "BDO_TaxAcc") & pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                        {
                            SAPbouiCOM.ChooseFromListEvent oCFLEvento = (SAPbouiCOM.ChooseFromListEvent)pVal;
                            chooseFromList(oForm, pVal, oCFLEvento);
                        }
                        if ((pVal.ItemUID == "PrBaseEDR" || pVal.ItemUID == "PrBaseEGI" || pVal.ItemUID == "BDOSPnCoP" || pVal.ItemUID == "BDOSPnPh" || pVal.ItemUID == "BDOSFlWhsE") & pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                        {
                            SAPbouiCOM.ChooseFromListEvent oCFLEvento = (SAPbouiCOM.ChooseFromListEvent)pVal;
                            chooseFromList(oForm, pVal, oCFLEvento);
                        }

                        if (pVal.ItemUID == "BDO_TaxTyp" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_COMBO_SELECT & pVal.BeforeAction == false)
                        {
                            taxTyp_OnClick(oForm);
                        }
                    }

                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD && pVal.BeforeAction)
                    {
                        FORM_LOAD_FOR_ACTIVATE = true;
                    }

                    else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE && !pVal.BeforeAction)
                    {
                        if (FORM_LOAD_FOR_ACTIVATE)
                        {
                            oForm.Freeze(true);
                            try
                            {
                                if (Convert.ToInt32(oCompany.language) == 100007 || Convert.ToInt32(oCompany.language) == 3)
                                {
                                    Folder folder = oForm.Items.Item("1320002089").Specific;
                                    folder.Caption = "ელ. დღგ ანგ.";

                                    folder = oForm.Items.Item("34").Specific;
                                    folder.Caption = "აღრც. მონაც";

                                    folder = oForm.Items.Item("36").Specific;
                                    folder.Caption = "საწყ. ინიც";
                                }

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

                    oForm.Freeze(false);
                }
            }
        }

        public static bool IsDiscountUsed()
        {
            var result = false;
            Recordset oRecordSet = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            var query = new StringBuilder();
            query.Append("SELECT \"U_Discount\" \n");
            query.Append("FROM \"OADM\"");

            oRecordSet.DoQuery(query.ToString());

            if (!oRecordSet.EoF)
            {
                result = oRecordSet.Fields.Item("U_Discount").Value == "Y";
            }

            return result;
        }
    }
}
