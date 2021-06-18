using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Resources;
using System.Globalization;
using System.Data;
using BDO_Localisation_AddOn.TBC_Integration_Services;
using System.Data.SqlClient;
using BDO_Localisation_AddOn.BOG_Integration_Services;
using System.Net.Http;
using BDO_Localisation_AddOn.BOG_Integration_Services.Model;
using System.Threading.Tasks;

namespace BDO_Localisation_AddOn
{
    static partial class OutgoingPayment
    {
        public static bool ProfitTaxTypeIsSharing = false;
        public static SAPbouiCOM.Form CurrentForm;
        public static decimal oldAmountLC;
        public static decimal oldAmountFC;
        public static decimal oldDocRate;

        public static void createUserFields(out string errorText)
        {
            errorText = null;
            Dictionary<string, object> fieldskeysMap;
            Dictionary<string, string> listValidValuesDict;

            fieldskeysMap = new Dictionary<string, object>();
            listValidValuesDict = new Dictionary<string, string>();
            listValidValuesDict.Add("empty", ""); //ცარიელი
            listValidValuesDict.Add("initialState", "Initial State"); //შექმნილი
            listValidValuesDict.Add("draft", "Draft"); //დროებით შენახული
            listValidValuesDict.Add("registered", "Registered"); //დამატებული
            listValidValuesDict.Add("deleted", "Deleted"); //წაშლილი
            listValidValuesDict.Add("waitingForCertification", "Waiting For Certification"); //ავტორიზაციის მოლოდინში
            listValidValuesDict.Add("inProgress", "In Progress"); //დამუშავების პროცესში
            listValidValuesDict.Add("finished", "Finished"); //დასრულებული
            listValidValuesDict.Add("failed", "Failed"); //უარყოფილი
            listValidValuesDict.Add("cancelled", "Cancelled"); //გაუქმებული
            listValidValuesDict.Add("forSigning", "For Signing"); //ხელმოსაწერი
            listValidValuesDict.Add("signed", "Signed"); //ხელმოწერილი
            listValidValuesDict.Add("finishedWithErrors", "Finished With Errors"); //დასრულებულია შეცდომებით
            listValidValuesDict.Add("readyToLoad", "Ready To Load"); //მომზადებულია გადასატვირთად
            listValidValuesDict.Add("notToUpload", "Not To Upload"); //არ გადაიტვირთოს
            listValidValuesDict.Add("resend", "Resend"); //ხელახლა გადაიტვირთოს
            listValidValuesDict.Add("downloadedFromTheBank", "Downloaded From The Bank"); //ჩამოტვირთულია ბანკიდან

            fieldskeysMap.Add("Name", "status"); //სტატუსი
            fieldskeysMap.Add("TableName", "OVPM");
            fieldskeysMap.Add("Description", "Status");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);
            fieldskeysMap.Add("ValidValues", listValidValuesDict);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            bool result = UDO.addNewValidValuesUserFieldsMD("OVPM", "status", "signed", "Signed", out errorText);
            result = UDO.addNewValidValuesUserFieldsMD("OVPM", "status", "notToUpload", "Not To Upload", out errorText);
            result = UDO.addNewValidValuesUserFieldsMD("OVPM", "status", "downloadedFromTheBank", "Downloaded From The Bank", out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "bStatus"); //პაკეტური ტრანზაქციის სტატუსი
            fieldskeysMap.Add("TableName", "OVPM");
            fieldskeysMap.Add("Description", "Batch Status");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);
            fieldskeysMap.Add("ValidValues", listValidValuesDict);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            result = UDO.addNewValidValuesUserFieldsMD("OVPM", "bStatus", "signed", "Signed", out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "opType"); //ოპერაციის ტიპი
            fieldskeysMap.Add("TableName", "OVPM");
            fieldskeysMap.Add("Description", "Operation Type");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //სახაზინო კოდი
            fieldskeysMap.Add("Name", "tresrCode");
            fieldskeysMap.Add("TableName", "OVPM");
            fieldskeysMap.Add("Description", "Treasury Code");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //მიმღების ანგარიში
            fieldskeysMap.Add("Name", "creditAcct");
            fieldskeysMap.Add("TableName", "OVPM");
            fieldskeysMap.Add("Description", "Credit Account");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //მიმღების ანგარიშის ვალუტა
            fieldskeysMap.Add("Name", "crdtActCur");
            fieldskeysMap.Add("TableName", "OVPM");
            fieldskeysMap.Add("Description", "Credit Account Currency");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //ტრანზაქციის ID
            fieldskeysMap.Add("Name", "paymentID");
            fieldskeysMap.Add("TableName", "OVPM");
            fieldskeysMap.Add("Description", "Payment ID");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //პაკეტური ტრანზაქციის ID
            fieldskeysMap.Add("Name", "bPaymentID");
            fieldskeysMap.Add("TableName", "OVPM");
            fieldskeysMap.Add("Description", "Batch Payment ID");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //პაკეტური ტრანზაქციაში დოკუმენტის პოზიცია
            fieldskeysMap.Add("Name", "posBPaymnt");
            fieldskeysMap.Add("TableName", "OVPM");
            fieldskeysMap.Add("Description", "Position Batch Payment");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Numeric);
            fieldskeysMap.Add("EditSize", 11);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //UniqueID (BOG) - ის დროს ივსება
            fieldskeysMap.Add("Name", "uniqueID");
            fieldskeysMap.Add("TableName", "OVPM");
            fieldskeysMap.Add("Description", "UniqueID");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            listValidValuesDict = new Dictionary<string, string>();
            listValidValuesDict.Add("SHA", "SHA"); //მიმღები მიიღებს შუამავალი ბანკის საკომისიოთი ნაკლებ თანხას (SHA)
            listValidValuesDict.Add("OUR", "OUR"); //მიმღები მიიღებს სრულ თანხას, გადარიცხვის საკომისიოს დაემატება 20USD/30EUR (OUR)

            fieldskeysMap.Add("Name", "chrgDtls"); //ხარჯი
            fieldskeysMap.Add("TableName", "OVPM");
            fieldskeysMap.Add("Description", "Charge Details");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);
            fieldskeysMap.Add("ValidValues", listValidValuesDict);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            listValidValuesDict = new Dictionary<string, string>();
            listValidValuesDict.Add("BULK", "BULK"); //BULK - სტანდარტული გადარიცხვა
            listValidValuesDict.Add("MT103", "MT103"); //MT103 ინდივიდუალური გადარიცხვა

            fieldskeysMap.Add("Name", "dsptchType"); //გადარიცხვის მეთოდი
            fieldskeysMap.Add("TableName", "OVPM");
            fieldskeysMap.Add("Description", "Dispatch Type");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);
            fieldskeysMap.Add("ValidValues", listValidValuesDict);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "rprtCode"); //რეპორტის კოდი
            fieldskeysMap.Add("TableName", "OVPM");
            fieldskeysMap.Add("Description", "Reporting Code");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //დანიშნულება
            fieldskeysMap.Add("Name", "descrpt");
            fieldskeysMap.Add("TableName", "OVPM");
            fieldskeysMap.Add("Description", "Description");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 254);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //დამატებითი დანიშნულება
            fieldskeysMap.Add("Name", "addDescrpt");
            fieldskeysMap.Add("TableName", "OVPM");
            fieldskeysMap.Add("Description", "Additional Description");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 254);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //დოკუმენტის ნომერი ინტ. ბანკში
            fieldskeysMap.Add("Name", "docNumber");
            fieldskeysMap.Add("TableName", "OVPM");
            fieldskeysMap.Add("Description", "Document Number");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //ინტ. ბანკის ოპერაციის კოდი
            fieldskeysMap.Add("Name", "transCode");
            fieldskeysMap.Add("TableName", "OVPM");
            fieldskeysMap.Add("Description", "Transaction Code");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //ინტ. ბანკის ოპერაციის კოდი 2
            fieldskeysMap.Add("Name", "opCode");
            fieldskeysMap.Add("TableName", "OVPM");
            fieldskeysMap.Add("Description", "Operation Code");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //ტრანზაქციის ID 2
            fieldskeysMap.Add("Name", "ePaymentID");
            fieldskeysMap.Add("TableName", "OVPM");
            fieldskeysMap.Add("Description", "External Payment ID");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            //მოგების გადასახადი
            fieldskeysMap = new Dictionary<string, object>(); //ბეგრება განაწილებული მოგებით
            fieldskeysMap.Add("Name", "liablePrTx");
            fieldskeysMap.Add("TableName", "OVPM");
            fieldskeysMap.Add("Description", "Liable to Profit Tax");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 1);
            fieldskeysMap.Add("DefaultValue", "N");

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // დაბეგვრის ობიექტი
            fieldskeysMap.Add("Name", "prBase");
            fieldskeysMap.Add("TableName", "OVPM");
            fieldskeysMap.Add("Description", "Profit Base");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 15);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // დაბეგვრის ობიექტის სახელი
            fieldskeysMap.Add("Name", "prBsDscr");
            fieldskeysMap.Add("TableName", "OVPM");
            fieldskeysMap.Add("Description", "Profit Base DEscription");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 100);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "amtPrTx"); //მოგების გადასახადი
            fieldskeysMap.Add("TableName", "OVPM");
            fieldskeysMap.Add("Description", "Profit Tax Amount");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //Outgoing DocEntry //OVPM
            fieldskeysMap.Add("Name", "outDoc");
            fieldskeysMap.Add("TableName", "OVPM");
            fieldskeysMap.Add("Description", "Outgoing DocEntry");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 11);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            //საშემოსავლო
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSWhtAmt");
            fieldskeysMap.Add("TableName", "OVPM");
            fieldskeysMap.Add("Description", "Withholding Tax");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            //დასაქმებულის საპენსიო
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSPnPhAm");
            fieldskeysMap.Add("TableName", "OVPM");
            fieldskeysMap.Add("Description", "Physical Entity Pens. Amount");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            //დამსაქმებლის საპენსიო
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSPnCoAm");
            fieldskeysMap.Add("TableName", "OVPM");
            fieldskeysMap.Add("Description", "Company Pens. Amount");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            //დამსაქმებლის საპენსიოების განსხვავების თანხა AP(AP Reserve) Invoice-სა და Payment-ს შორის
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSPnCoDiffAm");
            fieldskeysMap.Add("TableName", "OVPM");
            fieldskeysMap.Add("Description", "Company Pens. Difference Amount");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            // use blanket agreement rate ranges
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "UseBlaAgRt");
            fieldskeysMap.Add("TableName", "OVPM");
            fieldskeysMap.Add("Description", "Use Blanket Agreement Rates");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 1);
            fieldskeysMap.Add("DefaultValue", "N");

            UDO.addUserTableFields(fieldskeysMap, out errorText);
        }

        public static void createFormItems(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            ProfitTaxTypeIsSharing = ProfitTax.ProfitTaxTypeIsSharing();

            Dictionary<string, object> formItems;
            Dictionary<string, string> listValidValuesDict = null;
            string itemName = "";

            SAPbouiCOM.Item oItem = oForm.Items.Item("53");
            int left_s = oItem.Left;
            int left_e = oForm.Items.Item("52").Left;
            int height = oItem.Height; //15
            int top = oItem.Top;
            int width_s = oItem.Width;
            int width_e = oForm.Items.Item("52").Width;

            SAPbobsCOM.Payments oVendorPayments = null;
            SAPbobsCOM.ValidValues oValidValues = null;
            SAPbobsCOM.Fields oFields = null;

            oVendorPayments = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments);
            oFields = oVendorPayments.UserFields.Fields;

            oValidValues = oFields.Item("U_status").ValidValues;

            listValidValuesDict = new Dictionary<string, string>();
            for (int i = 0; i < oValidValues.Count; i++)
            {
                string value = oValidValues.Item(i).Value;
                listValidValuesDict.Add(value, BDOSResources.getTranslate(value));
            }

            formItems = new Dictionary<string, object>();
            itemName = "statusS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top + 7 * height + 1);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("BankStatus"));
            formItems.Add("LinkTo", "statusCB");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "statusCB"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OVPM");
            formItems.Add("Alias", "U_status");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top + 7 * height + 1);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
            formItems.Add("DisplayDesc", true);
            formItems.Add("ValidValues", listValidValuesDict);
            formItems.Add("Enabled", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            oItem = oForm.Items.Item("73");
            left_s = oItem.Left;
            left_e = oForm.Items.Item("74").Left;
            height = oItem.Height;
            top = oItem.Top;
            width_s = oItem.Width;
            width_e = oForm.Items.Item("95").Width;

            top = top + height + 1;

            try
            {
                SAPbouiCOM.Item oItem1 = oForm.Items.Item("opTypeCB");
                SAPbouiCOM.ComboBox oComboBox = ((SAPbouiCOM.ComboBox)(oItem1.Specific));

                foreach (SAPbouiCOM.ValidValue oValidValue in oComboBox.ValidValues)
                {
                    if (oValidValue.Value == "transferToOwnAccount")
                        oComboBox.ValidValues.Remove("transferToOwnAccount", SAPbouiCOM.BoSearchKey.psk_ByValue);
                    if (oValidValue.Value == "currencyExchange")
                        oComboBox.ValidValues.Remove("currencyExchange", SAPbouiCOM.BoSearchKey.psk_ByValue);
                    if (oValidValue.Value == "treasuryTransfer")
                        oComboBox.ValidValues.Remove("treasuryTransfer", SAPbouiCOM.BoSearchKey.psk_ByValue);
                    if (oValidValue.Value == "other")
                        oComboBox.ValidValues.Remove("other", SAPbouiCOM.BoSearchKey.psk_ByValue);
                }
                oComboBox.ValidValues.Add("transferToOwnAccount", BDOSResources.getTranslate("transferToOwnAccount")); //გადარიცხვა პირად ანგარიშზე
                oComboBox.ValidValues.Add("currencyExchange", BDOSResources.getTranslate("currencyExchange")); //კონვერტაცია
                oComboBox.ValidValues.Add("treasuryTransfer", BDOSResources.getTranslate("treasuryTransfer")); //სახაზინო გადარიცხვა
                oComboBox.ValidValues.Add("other", BDOSResources.getTranslate("other")); //სხვა
            }
            catch
            {

                formItems = new Dictionary<string, object>();
                itemName = "opTypeS"; //10 characters
                formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                formItems.Add("Left", left_s);
                formItems.Add("Width", width_s);
                formItems.Add("Top", top);
                formItems.Add("Height", height);
                formItems.Add("UID", itemName);
                formItems.Add("Caption", BDOSResources.getTranslate("OperationType"));
                formItems.Add("LinkTo", "opTypeCB");
                formItems.Add("Visible", false);

                FormsB1.createFormItem(oForm, formItems, out errorText);
                if (errorText != null)
                {
                    return;
                }

                oValidValues = oFields.Item("U_opType").ValidValues;

                listValidValuesDict = new Dictionary<string, string>();
                listValidValuesDict.Add("transferToOwnAccount", BDOSResources.getTranslate("transferToOwnAccount")); //გადარიცხვა პირად ანგარიშზე
                listValidValuesDict.Add("currencyExchange", BDOSResources.getTranslate("currencyExchange")); //კონვერტაცია
                listValidValuesDict.Add("treasuryTransfer", BDOSResources.getTranslate("treasuryTransfer")); //სახაზინო გადარიცხვა
                listValidValuesDict.Add("other", BDOSResources.getTranslate("other")); //სხვა

                formItems = new Dictionary<string, object>();
                itemName = "opTypeCB"; //10 characters
                formItems.Add("isDataSource", true);
                formItems.Add("DataSource", "DBDataSources");
                formItems.Add("TableName", "OVPM");
                formItems.Add("Alias", "U_opType");
                formItems.Add("Bound", true);
                formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                formItems.Add("Left", left_e);
                formItems.Add("Width", width_e);
                formItems.Add("Top", top);
                formItems.Add("Height", height);
                formItems.Add("UID", itemName);
                formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
                formItems.Add("DisplayDesc", true);
                formItems.Add("ValidValues", listValidValuesDict);
                formItems.Add("Visible", false);

                FormsB1.createFormItem(oForm, formItems, out errorText);
                if (errorText != null)
                {
                    return;
                }
            }

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "tresrCodeS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("TreasuryCode"));
            formItems.Add("LinkTo", "tresrCodeE");
            formItems.Add("Visible", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "tresrCodeE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OVPM");
            formItems.Add("Alias", "U_tresrCode");
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

            bool multiSelection = false;
            string objectType = "231"; // HouseBankAccounts object
            string uniqueID_lf_HouseBankAccountCFL = "HouseBankAccount_CFL";
            FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_lf_HouseBankAccountCFL);

            formItems = new Dictionary<string, object>();
            itemName = "creditActS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("CorrespondingAccount"));
            formItems.Add("LinkTo", "creditActE");
            formItems.Add("Visible", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "creditActE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OVPM");
            formItems.Add("Alias", "U_creditAcct");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("ChooseFromListUID", uniqueID_lf_HouseBankAccountCFL);
            formItems.Add("ChooseFromListAlias", "Account");
            formItems.Add("Visible", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            //listValidValuesDict = new Dictionary<string, string>();
            listValidValuesDict = CommonFunctions.getCurrencyListForValidValues();

            formItems = new Dictionary<string, object>();
            itemName = "crdActCuCB"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OVPM");
            formItems.Add("Alias", "U_crdtActCur");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            formItems.Add("Left", left_e + width_e + 5);
            formItems.Add("Width", width_s / 2);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("ValidValues", listValidValuesDict);
            formItems.Add("Description", BDOSResources.getTranslate("CurrencyForExchange"));
            formItems.Add("Visible", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            oItem = oForm.Items.Item("53");
            left_s = oItem.Left;
            left_e = oForm.Items.Item("52").Left;
            height = oItem.Height;
            top = oForm.Items.Item("statusCB").Top;
            width_s = oItem.Width;
            width_e = oForm.Items.Item("52").Width;

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "paymentIDS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("PaymentID"));
            formItems.Add("LinkTo", "paymentIDE");
            formItems.Add("Visible", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "paymentIDE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OVPM");
            formItems.Add("Alias", "U_paymentID");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", false);
            formItems.Add("Visible", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "chrgDtlsS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("ChargeDetails"));
            formItems.Add("LinkTo", "chrgDtlsCB");
            formItems.Add("Visible", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            oValidValues = oFields.Item("U_chrgDtls").ValidValues;

            listValidValuesDict = new Dictionary<string, string>();
            for (int i = 0; i < oValidValues.Count; i++)
            {
                string value = oValidValues.Item(i).Value;
                listValidValuesDict.Add(value, BDOSResources.getTranslate(value));
            }

            formItems = new Dictionary<string, object>();
            itemName = "chrgDtlsCB"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OVPM");
            formItems.Add("Alias", "U_chrgDtls");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
            formItems.Add("DisplayDesc", true);
            formItems.Add("ValidValues", listValidValuesDict);
            formItems.Add("Visible", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "dsptTypeS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("DispatchType"));
            formItems.Add("LinkTo", "dsptTypeCB");
            formItems.Add("Visible", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            oValidValues = oFields.Item("U_dsptchType").ValidValues;

            listValidValuesDict = new Dictionary<string, string>();
            for (int i = 0; i < oValidValues.Count; i++)
            {
                string value = oValidValues.Item(i).Value;
                listValidValuesDict.Add(value, BDOSResources.getTranslate(value));
            }

            formItems = new Dictionary<string, object>();
            itemName = "dsptTypeCB"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OVPM");
            formItems.Add("Alias", "U_dsptchType");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
            formItems.Add("DisplayDesc", true);
            formItems.Add("ValidValues", listValidValuesDict);
            formItems.Add("Visible", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }


            //-------------------------------------------


            formItems = new Dictionary<string, object>();
            itemName = "rprtCodeS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top + height + 1);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("ReportingCode"));
            formItems.Add("LinkTo", "rprtCodeCB");
            //formItems.Add("Visible", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }



            listValidValuesDict = new Dictionary<string, string>();
            listValidValuesDict.Add("GDS", "GDS");
            listValidValuesDict.Add("ACM", "AGENCY COMMISSION");
            listValidValuesDict.Add("DCM", "TRADE COMMISION");
            listValidValuesDict.Add("AKA", "BONUS");

            formItems = new Dictionary<string, object>();
            itemName = "rprtCodeCB"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OVPM");
            formItems.Add("Alias", "U_rprtCode");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top + height + 1);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
            formItems.Add("DisplayDesc", true);
            formItems.Add("ValidValues", listValidValuesDict);
            //formItems.Add("Visible", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }


            //-------------------------------------------

            top = oForm.Items.Item("20").Top + oForm.Items.Item("20").Height;
            left_s = oForm.Items.Item("27").Left;
            left_e = oForm.Items.Item("26").Left;
            width_s = oForm.Items.Item("27").Width;
            width_e = oForm.Items.Item("26").Width;

            formItems = new Dictionary<string, object>();
            itemName = "descrptS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Description"));
            formItems.Add("LinkTo", "descrptE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "descrptE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OVPM");
            formItems.Add("Alias", "U_descrpt");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EXTEDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("ScrollBars", SAPbouiCOM.BoScrollBars.sb_Vertical);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "addDescrpS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("AdditionalDescription"));
            formItems.Add("LinkTo", "addDescrpE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "addDescrpE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OVPM");
            formItems.Add("Alias", "U_addDescrpt");
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

            //ღილაკები

            oItem = oForm.Items.Item("14");
            left_s = oItem.Left;
            height = oItem.Height; //15
            top = oItem.Top;
            width_s = oItem.Width;

            listValidValuesDict = new Dictionary<string, string>();
            listValidValuesDict.Add("updateStatus", BDOSResources.getTranslate("UpdateStatus"));
            listValidValuesDict.Add("import", BDOSResources.getTranslate("Export"));

            formItems = new Dictionary<string, object>();
            itemName = "operationB";
            formItems.Add("Caption", BDOSResources.getTranslate("Operations"));
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON_COMBO);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top - height - 2);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("ValidValues", listValidValuesDict);
            formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
            formItems.Add("AffectsFormMode", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }


            listValidValuesDict = new Dictionary<string, string>();
            listValidValuesDict.Add("readyToLoad", BDOSResources.getTranslate("readyToLoad")); //მომზადებულია გადასატვირთად
            listValidValuesDict.Add("resend", BDOSResources.getTranslate("resend")); //ხელახლა გადაიტვირთოს
            listValidValuesDict.Add("notToUpload", BDOSResources.getTranslate("notToUpload")); //არ გადაიტვირთოს

            formItems = new Dictionary<string, object>();
            itemName = "setStatusB";
            formItems.Add("Caption", BDOSResources.getTranslate("SetStatus"));
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON_COMBO);
            formItems.Add("Left", left_s - width_s + 16);
            formItems.Add("Width", width_s - 20);
            formItems.Add("Top", top - height - 2);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("ValidValues", listValidValuesDict);
            formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
            formItems.Add("AffectsFormMode", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            //მოგების გადასახადი
            SAPbouiCOM.Item oItemS = oForm.Items.Item("addDescrpS");
            SAPbouiCOM.Item oItemE = oForm.Items.Item("addDescrpE");

            top = oItemS.Top + oItemS.Height;
            left_s = oItemS.Left;
            left_e = oItemE.Left;
            width_s = oItemS.Width;
            width_e = oItemE.Width;

            formItems = new Dictionary<string, object>();
            itemName = "liablePrTx";
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OVPM");
            formItems.Add("Alias", "U_liablePrTx");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
            formItems.Add("Length", 1);
            formItems.Add("Left", left_s);
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

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "PrBaseS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", 14);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("TaxableObject"));
            formItems.Add("LinkTo", "PrBaseE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            multiSelection = false;
            objectType = "UDO_F_BDO_PTBS_D";
            string uniqueID_lf_ProfitBaseCFL = "CFL_ProfitBase";
            FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_lf_ProfitBaseCFL);

            formItems = new Dictionary<string, object>();
            itemName = "PrBaseE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OVPM");
            formItems.Add("Alias", "U_prBase");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", 30);
            formItems.Add("Top", top);
            formItems.Add("Height", 14);
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
            formItems.Add("TableName", "OVPM");
            formItems.Add("Alias", "U_PrBsDscr");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e + 32);
            formItems.Add("Width", width_e - 32);
            formItems.Add("Top", top);
            formItems.Add("Height", 14);
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
            itemName = "PrBaseLB"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            formItems.Add("Left", left_e - 20);
            formItems.Add("Top", top);
            formItems.Add("Height", 14);
            formItems.Add("UID", itemName);
            formItems.Add("LinkTo", "PrBaseE");
            formItems.Add("LinkedObjectType", objectType);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            oItem = oForm.Items.Item("37");
            formItems = new Dictionary<string, object>();
            itemName = "AmtPrTxS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", oItem.Left);
            formItems.Add("Width", oItem.Width);
            formItems.Add("Top", oItem.Top + oItem.Height + 1);
            formItems.Add("Height", oItem.Height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("ProfitTaxAmount"));
            formItems.Add("LinkTo", "AmtPrTxE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            oItem = oForm.Items.Item("13");
            formItems = new Dictionary<string, object>();
            itemName = "AmtPrTxE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "UserDataSources");
            formItems.Add("TableName", "");
            formItems.Add("Alias", itemName);
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SUM);
            formItems.Add("Length", 11);
            formItems.Add("Left", oItem.Left);
            formItems.Add("Width", oItem.Width);
            formItems.Add("Top", oItem.Top + oItem.Height + 1);
            formItems.Add("Height", oItem.Height);
            formItems.Add("UID", itemName);
            formItems.Add("AffectsFormMode", false);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", false);
            formItems.Add("RightJustified", true);


            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            oItem = oForm.Items.Item("234000001");
            formItems = new Dictionary<string, object>();
            itemName = "FillAmtTxs";
            formItems.Add("Caption", BDOSResources.getTranslate("Fill"));
            formItems.Add("Size", 8);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            formItems.Add("Left", oItem.Left);
            formItems.Add("Width", oItem.Width);
            formItems.Add("Top", oItem.Top + oItem.Height + 1);
            formItems.Add("Height", oItem.Height);
            formItems.Add("UID", itemName);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            //////////////////////////////////////

            oItem = oForm.Items.Item("10001005");
            formItems = new Dictionary<string, object>();
            itemName = "ChngDcDt";
            formItems.Add("Caption", BDOSResources.getTranslate("ChangeDocDate"));
            formItems.Add("Size", 8);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            formItems.Add("Left", oItem.Left - oItem.Width - 1);
            formItems.Add("Width", oItem.Width);
            formItems.Add("Top", oItem.Top);
            formItems.Add("Height", oItem.Height);
            formItems.Add("UID", itemName);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }
            //////////////////////////////////////

            //---------------------- საპენსიო
            oItem = oForm.Items.Item("234000004");
            left_s = oItem.Left;
            height = oItem.Height;
            width_s = oItem.Width;
            top = oItem.Top;

            oItem = oForm.Items.Item("234000005");
            left_e = oItem.Left;
            width_e = oItem.Width;

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "BDOSWhtS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("WithhTax") + " (LC)");
            formItems.Add("LinkTo", "BDOSWhtAmt");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "BDOSWhtAmt"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OVPM");
            formItems.Add("Alias", "U_BDOSWhtAmt");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            //Calculate
            formItems = new Dictionary<string, object>();
            itemName = "CalcPhcEnt";
            formItems.Add("Caption", BDOSResources.getTranslate("Calculate")); //Calculate Physical Entity Taxes
            formItems.Add("Size", 8);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            formItems.Add("Left", left_e + width_e + 5);
            formItems.Add("Width", 100);
            formItems.Add("Top", top);
            formItems.Add("Height", oItem.Height);
            formItems.Add("UID", itemName);
            //formItems.Add("Image", "WS_ANALYTICS_COLLAPSE_BTN_ITEM");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "BDOSPnPhS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("PhysEntityPension") + " (LC)");
            formItems.Add("LinkTo", "BDOSPnPhAm");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "BDOSPnPhAm"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OVPM");
            formItems.Add("Alias", "U_BDOSPnPhAm");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "BDOSPnCoS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("CompPension") + " (LC)");
            formItems.Add("LinkTo", "BDOSPnCoAm");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "BDOSPnCoAm"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OVPM");
            formItems.Add("Alias", "U_BDOSPnCoAm");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "PnCoDiffAm"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OVPM");
            formItems.Add("Alias", "U_BDOSPnCoDiffAm");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e + width_e + 5);
            formItems.Add("Width", 50);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            //---------------------- საპენსიო


            // -------------------- Use blanket agreement rates-----------------

            height = oForm.Items.Item("234000005").Height;
            top = oForm.Items.Item("234000005").Top;
            //int left = left_e + width_e + 5;
            int left = oForm.Items.Item("CalcPhcEnt").Left - 3;

            formItems = new Dictionary<string, object>();
            itemName = "UsBlaAgRtS"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OVPM");
            formItems.Add("Alias", "U_UseBlaAgRt");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
            formItems.Add("Length", 1);
            formItems.Add("Left", left);
            formItems.Add("Width", 100);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("UseBlAgrRt"));
            formItems.Add("ValOff", "N");
            formItems.Add("ValOn", "Y");
            formItems.Add("FromPane", 2);
            formItems.Add("ToPane", 3);
            formItems.Add("Enabled", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }
        }

        public static void CheckAccounts(SAPbouiCOM.Form oForm, out string errorText)
        {
            bool isError = false;
            errorText = null;

            string taxType = oForm.DataSources.DBDataSources.Item("OCRD").GetValue("U_BDO_TaxTyp", 0);

            if (taxType.Trim() != "12")
                return;

            string cardCode = oForm.Items.Item("5").Specific.Value.ToString();

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string query = "SELECT \"OCRD\".\"ECVatGroup\" " +
                            "FROM \"OCRD\" " +
                            "WHERE \"OCRD\".\"CardCode\" = '" + cardCode + "'";

            oRecordSet.DoQuery(query);

            if (!oRecordSet.EoF)
            {
                if (oRecordSet.Fields.Item("ECVatGroup").Value == "")
                {
                    isError = true;
                }
                else
                {
                    string vatGrp = oRecordSet.Fields.Item("ECVatGroup").Value;
                    oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    query = "SELECT \"OVTG\".\"U_BDOSAccF\", " +
                                    "\"OVTG\".\"Account\" " +
                                    "FROM \"OVTG\" " +
                                    "WHERE \"OVTG\".\"Code\" = '" + vatGrp + "'";

                    oRecordSet.DoQuery(query);

                    if (!oRecordSet.EoF)
                    {
                        if (oRecordSet.Fields.Item("U_BDOSAccF").Value == "" || oRecordSet.Fields.Item("Account").Value == "")
                            errorText = BDOSResources.getTranslate("CheckVatGroupAccounts");
                    }
                }
            }

            if (isError)
                errorText = BDOSResources.getTranslate("CheckVatGroupForBP");
        }

        public static void taxes_OnClick(SAPbouiCOM.Form oForm)
        {
            bool isWTLiable = oForm.DataSources.DBDataSources.Item("OVPM").GetValue("U_liablePrTx", 0).Trim() == "Y";
            bool payNoDoc = oForm.DataSources.DBDataSources.Item("OVPM").GetValue("PayNoDoc", 0).Trim() == "Y";

            if (!isWTLiable) //|| PayNoDoc != "Y"
            {
                SAPbouiCOM.EditText oEdit = oForm.Items.Item("PrBaseE").Specific;
                oEdit.Value = "";

                oEdit = oForm.Items.Item("PrBsDscr").Specific;
                oEdit.Value = "";

                oEdit = oForm.Items.Item("AmtPrTxE").Specific;
                oEdit.Value = "";

                oForm.Items.Item("26").Click();
            }

            if (!payNoDoc && isWTLiable)
            {
                SAPbouiCOM.CheckBox oliablePrTx = oForm.Items.Item("liablePrTx").Specific;
                oliablePrTx.Checked = false;
                Program.uiApp.SetStatusBarMessage("Payment on Account should be checked for Profit Taxes",
                    SAPbouiCOM.BoMessageTime.bmt_Short);
            }

            //fillAmountTaxes( oForm, out errorText);

            setVisibleFormItems(oForm);
        }

        public static void comboSelect(SAPbouiCOM.Form oForm, SAPbouiCOM.ItemEvent pVal, out string errorText)
        {
            errorText = null;

            try
            {
                if (!pVal.BeforeAction)
                {
                    if (pVal.ItemUID == "operationB")
                    {
                        SAPbouiCOM.ButtonCombo oButtonCombo = ((SAPbouiCOM.ButtonCombo)(oForm.Items.Item("operationB").Specific));

                        string selectedOperation = null;
                        if (oButtonCombo.Selected != null)
                        {
                            selectedOperation = oButtonCombo.Selected.Value;
                        }

                        oForm.Freeze(false);
                        oButtonCombo.Caption = BDOSResources.getTranslate("Operations");

                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            Program.uiApp.MessageBox(BDOSResources.getTranslate("ToCompleteOperationWriteDocument"));
                            return;
                        }
                        else if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            Program.uiApp.MessageBox(BDOSResources.getTranslate("OperationImpossibleInMode"));
                            return;
                        }

                        if (selectedOperation != null)
                        {
                            string isPayToBank = oForm.DataSources.DBDataSources.Item("OVPM").GetValue("IsPaytoBnk", 0).Trim();
                            string docType = oForm.DataSources.DBDataSources.Item("OVPM").GetValue("DocType", 0).Trim();

                            if (isPayToBank == "Y" || docType == "A")
                            {
                                List<int> docEntryList = new List<int>();
                                docEntryList.Add(Convert.ToInt32(oForm.DataSources.DBDataSources.Item("OVPM").GetValue("DocEntry", 0)));
                                string trsfrAcct = oForm.DataSources.DBDataSources.Item("OVPM").GetValue("TrsfrAcct", 0).Trim();
                                string bankProgram = CommonFunctions.getBankProgram(trsfrAcct);

                                if (bankProgram == "TBC")
                                {
                                    BDOSAuthenticationFormTBC.createForm(oForm, selectedOperation, docEntryList, false, null, null, out errorText);
                                }
                                else if (bankProgram == "BOG")
                                {
                                    BDOSAuthenticationFormBOG.createForm(oForm, selectedOperation, docEntryList, false, null, null, out errorText);
                                }
                            }
                            else
                            {
                                Program.uiApp.MessageBox(BDOSResources.getTranslate("TheDocumentIsNotForInternetBanking") + "!"); //დოკუმენტი არ არის განკუთვნილი ინტერნეტბანკში გადატვირთვისთვის!
                                return;
                            }
                            FormsB1.SimulateRefresh();
                        }
                    }

                    if (pVal.ItemUID == "setStatusB")
                    {
                        SAPbouiCOM.ButtonCombo oButtonCombo = ((SAPbouiCOM.ButtonCombo)(oForm.Items.Item("setStatusB").Specific));

                        string selectedOperation = null;
                        if (oButtonCombo.Selected != null)
                        {
                            selectedOperation = oButtonCombo.Selected.Value;
                        }

                        oForm.Freeze(false);
                        oButtonCombo.Caption = BDOSResources.getTranslate("SetStatus");

                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            Program.uiApp.MessageBox(BDOSResources.getTranslate("ToCompleteOperationWriteDocument"));
                            return;
                        }
                        else if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            Program.uiApp.MessageBox(BDOSResources.getTranslate("OperationImpossibleInMode"));
                            return;
                        }

                        if (selectedOperation != null)
                        {
                            int answer = Program.uiApp.MessageBox(BDOSResources.getTranslate("SelectedDocumentsStatusWillBe") + " \"" + BDOSResources.getTranslate(selectedOperation) + "\". " + BDOSResources.getTranslate("WouldYouWantToContinueTheOperation") + "?", 1, BDOSResources.getTranslate("Yes"), BDOSResources.getTranslate("No"), ""); //მონიშნული დოკუმენტების სტატუსი გახდება, გსურთ ოპერაციის გაგრძელება

                            if (answer == 2)
                            {
                                return;
                            }

                            List<int> docEntryList = new List<int>();
                            docEntryList.Add(Convert.ToInt32(oForm.DataSources.DBDataSources.Item("OVPM").GetValue("DocEntry", 0)));

                            List<string> infoList = BDOSInternetBanking.setStatusImport(docEntryList, selectedOperation);
                            for (int i = 0; i < infoList.Count; i++)
                            {
                                Program.uiApp.SetStatusBarMessage(infoList[i], SAPbouiCOM.BoMessageTime.bmt_Short, false);
                            }
                            FormsB1.SimulateRefresh();
                        }
                    }

                    if (pVal.ItemUID == "opTypeCB")
                    {
                        string opType = oForm.DataSources.DBDataSources.Item("OVPM").GetValue("U_opType", 0).Trim();
                        SAPbouiCOM.EditText oEditText;
                        SAPbouiCOM.ComboBox oComboBox;

                        if (opType == "transferToOwnAccount") //გადარიცხვა პირად ანგარიშზე
                        {
                            oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("tresrCodeE").Specific;
                            oEditText.Value = "";
                            oComboBox = (SAPbouiCOM.ComboBox)oForm.Items.Item("crdActCuCB").Specific;
                            oComboBox.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue);
                            oComboBox = (SAPbouiCOM.ComboBox)oForm.Items.Item("chrgDtlsCB").Specific;
                            oComboBox.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue);
                            oComboBox = (SAPbouiCOM.ComboBox)oForm.Items.Item("dsptTypeCB").Specific;
                            oComboBox.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue);
                        }
                        else if (opType == "currencyExchange") //კონვერტაცია
                        {
                            oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("tresrCodeE").Specific;
                            oEditText.Value = "";
                            oComboBox = (SAPbouiCOM.ComboBox)oForm.Items.Item("chrgDtlsCB").Specific;
                            oComboBox.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue);
                            oComboBox = (SAPbouiCOM.ComboBox)oForm.Items.Item("dsptTypeCB").Specific;
                            oComboBox.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue);
                        }
                        else if (opType == "treasuryTransfer") //სახაზინო გადარიცხვა
                        {
                            oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("creditActE").Specific;
                            oEditText.Value = "";
                            oComboBox = (SAPbouiCOM.ComboBox)oForm.Items.Item("crdActCuCB").Specific;
                            oComboBox.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue);
                            oComboBox = (SAPbouiCOM.ComboBox)oForm.Items.Item("chrgDtlsCB").Specific;
                            oComboBox.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue);
                            oComboBox = (SAPbouiCOM.ComboBox)oForm.Items.Item("dsptTypeCB").Specific;
                            oComboBox.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue);
                        }
                    }
                }
                else if (pVal.BeforeAction)
                {
                    if (pVal.ItemUID == "operationB")
                    {
                        oForm.Freeze(true);
                    }
                    if (pVal.ItemUID == "setStatusB")
                    {
                        oForm.Freeze(true);
                    }
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

        public static void formDataLoad(SAPbouiCOM.Form oForm)
        {
            setVisibleFormItems(oForm);

            fillAmountTaxes(oForm);
        }

        public static void setVisibleFormItems(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.Item oItem;
            oForm.Freeze(true);

            try
            {
                oForm.Items.Item("statusCB").Enabled = false;
                string docEntry = oForm.DataSources.DBDataSources.Item("OVPM").GetValue("DocEntry", 0).Trim();
                string opType = oForm.DataSources.DBDataSources.Item("OVPM").GetValue("U_opType", 0).Trim();
                string docType = oForm.DataSources.DBDataSources.Item("OVPM").GetValue("DocType", 0).Trim();
                string PayNoDoc = oForm.DataSources.DBDataSources.Item("OVPM").GetValue("PayNoDoc", 0).Trim();
                string CardCode = oForm.DataSources.DBDataSources.Item("OVPM").GetValue("CardCode", 0).Trim();
                string DocType = oForm.DataSources.DBDataSources.Item("OVPM").GetValue("DocType", 0).Trim();
                string draftKey = oForm.DataSources.DBDataSources.Item("OVPM").GetValue("DocEntry", 0).Trim();
                string transId = oForm.DataSources.DBDataSources.Item("OVPM").GetValue("TransId", 0);

                bool ProfitTaxValuesVisible = (ProfitTaxTypeIsSharing && DocType == "S");

                bool docEntryIsEmpty = string.IsNullOrEmpty(docEntry);
                bool transIdIsEmpty = string.IsNullOrEmpty(transId);

                string liablePrTx = oForm.DataSources.DBDataSources.Item("OVPM").GetValue("U_liablePrTx", 0).Trim();
                oForm.Items.Item("liablePrTx").Enabled = (docEntryIsEmpty);
                oForm.Items.Item("PrBaseE").Enabled = (liablePrTx == "Y" && docEntryIsEmpty);

                string uniqueID_lf_ProfitBaseCFL = "CFL_ProfitBase";
                oForm.Items.Item("PrBaseE").Specific.ChooseFromListUID = uniqueID_lf_ProfitBaseCFL;
                oForm.Items.Item("PrBaseE").Specific.ChooseFromListAlias = "Code";

                oForm.Items.Item("liablePrTx").Visible = ProfitTaxValuesVisible;
                oForm.Items.Item("PrBaseS").Visible = ProfitTaxValuesVisible;
                oForm.Items.Item("PrBaseE").Visible = ProfitTaxValuesVisible;
                oForm.Items.Item("PrBsDscr").Visible = ProfitTaxValuesVisible;
                oForm.Items.Item("PrBaseLB").Visible = ProfitTaxValuesVisible;
                oForm.Items.Item("AmtPrTxS").Visible = ProfitTaxValuesVisible;
                oForm.Items.Item("AmtPrTxE").Visible = ProfitTaxValuesVisible;
                oForm.Items.Item("FillAmtTxs").Visible = ProfitTaxValuesVisible;
                oForm.Items.Item("ChngDcDt").Visible = false; //(draftKey != "" && oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE);

                //საშემოსავლო + საპენსიო
                bool PensionVisible = (DocType == "S");
                oForm.Items.Item("BDOSWhtS").Visible = PensionVisible;
                oForm.Items.Item("BDOSPnPhS").Visible = PensionVisible;
                oForm.Items.Item("BDOSPnCoS").Visible = PensionVisible;
                oForm.Items.Item("BDOSWhtAmt").Visible = PensionVisible;
                oForm.Items.Item("BDOSPnPhAm").Visible = PensionVisible;
                oForm.Items.Item("BDOSPnCoAm").Visible = PensionVisible;
                oForm.Items.Item("PnCoDiffAm").Visible = PensionVisible;
                oForm.Items.Item("CalcPhcEnt").Visible = PensionVisible;
                oForm.Items.Item("BDOSWhtAmt").Enabled = PensionVisible && transIdIsEmpty;
                oForm.Items.Item("BDOSPnPhAm").Enabled = PensionVisible && transIdIsEmpty;
                oForm.Items.Item("BDOSPnCoAm").Enabled = PensionVisible && transIdIsEmpty;
                oForm.Items.Item("PnCoDiffAm").Enabled = PensionVisible && transIdIsEmpty;
                oForm.Items.Item("CalcPhcEnt").Enabled = PensionVisible && transIdIsEmpty;
                //საშემოსავლო + საპენსიო

                oItem = oForm.Items.Item("opTypeCB");
                oItem.Enabled = docEntryIsEmpty;

                if (string.IsNullOrEmpty(opType) || opType == "transferToOwnAccount" || opType == "currencyExchange" || opType == "treasuryTransfer")
                {
                    try
                    {
                        SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("71").Specific));

                        SAPbouiCOM.Columns oColumns = oMatrix.Columns;
                        SAPbouiCOM.Column oColumn;

                        oColumn = oColumns.Item("U_employee");
                        oColumn.Visible = false;
                        oColumn = oColumns.Item("U_employeeN");
                        oColumn.Visible = false;
                        oColumn = oColumns.Item("U_creditAcct");
                        oColumn.Visible = false;
                        oColumn = oColumns.Item("U_bankCode");
                        oColumn.Visible = false;
                        oColumn = oColumns.Item("U_accrCode");
                        oColumn.Visible = false;
                        oColumn = oColumns.Item("U_wTaxAmount");
                        oColumn.Visible = false;

                        oMatrix.AutoResizeColumns();
                    }
                    catch { }
                }

                Dictionary<string, string> dataForTransferType = getDataForTransferType(oForm);
                string transferType = getTransferType(dataForTransferType, out var errorText);

                oItem = oForm.Items.Item("rprtCodeS");
                oItem.Visible = false;
                oItem = oForm.Items.Item("rprtCodeCB");
                oItem.Visible = false;

                if (transferType == "TransferToForeignCurrencyPaymentOrderIo")
                {
                    oItem = oForm.Items.Item("rprtCodeS");
                    oItem.Visible = true;
                    oItem = oForm.Items.Item("rprtCodeCB");
                    oItem.Visible = true;
                }

                if (docType == "A")
                {
                    oItem = oForm.Items.Item("opTypeS");
                    oItem.Visible = true;
                    oItem = oForm.Items.Item("opTypeCB");
                    oItem.Visible = true;
                    oItem = oForm.Items.Item("chrgDtlsS");
                    oItem.Visible = false;
                    oItem = oForm.Items.Item("chrgDtlsCB");
                    oItem.Visible = false;
                    oItem = oForm.Items.Item("dsptTypeS");
                    oItem.Visible = false;
                    oItem = oForm.Items.Item("dsptTypeCB");
                    oItem.Visible = false;

                    if (opType == "transferToOwnAccount") //გადარიცხვა პირად ანგარიშზე
                    {
                        oItem = oForm.Items.Item("creditActS");
                        oItem.Visible = true;
                        oItem = oForm.Items.Item("creditActE");
                        oItem.Visible = true;
                        oItem = oForm.Items.Item("crdActCuCB");
                        oItem.Visible = false;
                        oItem = oForm.Items.Item("tresrCodeS");
                        oItem.Visible = false;
                        oItem = oForm.Items.Item("tresrCodeE");
                        oItem.Visible = false;

                        if (transferType == "TransferToOtherBankForeignCurrencyPaymentOrderIo" || transferType == "TransferToForeignCurrencyPaymentOrderIo")
                        {
                            oItem = oForm.Items.Item("chrgDtlsS");
                            oItem.Visible = true;
                            oItem = oForm.Items.Item("chrgDtlsCB");
                            oItem.Visible = true;
                        }
                        else if (transferType == "TransferToNationalCurrencyPaymentOrderIo")
                        {
                            oItem = oForm.Items.Item("dsptTypeS");
                            oItem.Visible = true;
                            oItem = oForm.Items.Item("dsptTypeCB");
                            oItem.Visible = true;
                        }
                    }
                    else if (opType == "currencyExchange") //კონვერტაცია
                    {
                        oItem = oForm.Items.Item("creditActS");
                        oItem.Visible = true;
                        oItem = oForm.Items.Item("creditActE");
                        oItem.Visible = true;
                        oItem = oForm.Items.Item("crdActCuCB");
                        oItem.Visible = true;
                        oItem = oForm.Items.Item("tresrCodeS");
                        oItem.Visible = false;
                        oItem = oForm.Items.Item("tresrCodeE");
                        oItem.Visible = false;

                        if (transferType == "TransferToOtherBankForeignCurrencyPaymentOrderIo" || transferType == "TransferToForeignCurrencyPaymentOrderIo")
                        {
                            oItem = oForm.Items.Item("chrgDtlsS");
                            oItem.Visible = true;
                            oItem = oForm.Items.Item("chrgDtlsCB");
                            oItem.Visible = true;
                        }
                        else if (transferType == "TransferToNationalCurrencyPaymentOrderIo")
                        {
                            oItem = oForm.Items.Item("dsptTypeS");
                            oItem.Visible = true;
                            oItem = oForm.Items.Item("dsptTypeCB");
                            oItem.Visible = true;
                        }
                    }
                    else if (opType == "treasuryTransfer") //სახაზინო გადარიცხვა
                    {
                        oItem = oForm.Items.Item("creditActS");
                        oItem.Visible = false;
                        oItem = oForm.Items.Item("creditActE");
                        oItem.Visible = false;
                        oItem = oForm.Items.Item("crdActCuCB");
                        oItem.Visible = false;
                        oItem = oForm.Items.Item("tresrCodeS");
                        oItem.Visible = true;
                        oItem = oForm.Items.Item("tresrCodeE");
                        oItem.Visible = true;
                    }
                    else if (opType == "paymentToEmployee" || opType == "salaryPayment")
                    {
                        if (transferType == "TransferToOtherBankForeignCurrencyPaymentOrderIo" || transferType == "TransferToForeignCurrencyPaymentOrderIo")
                        {
                            oItem = oForm.Items.Item("chrgDtlsS");
                            oItem.Visible = true;
                            oItem = oForm.Items.Item("chrgDtlsCB");
                            oItem.Visible = true;
                        }
                        else if (transferType == "TransferToNationalCurrencyPaymentOrderIo")
                        {
                            oItem = oForm.Items.Item("dsptTypeS");
                            oItem.Visible = true;
                            oItem = oForm.Items.Item("dsptTypeCB");
                            oItem.Visible = true;
                        }
                        oItem = oForm.Items.Item("tresrCodeS");
                        oItem.Visible = false;
                        oItem = oForm.Items.Item("tresrCodeE");
                        oItem.Visible = false;
                        oItem = oForm.Items.Item("creditActS");
                        oItem.Visible = false;
                        oItem = oForm.Items.Item("creditActE");
                        oItem.Visible = false;
                        oItem = oForm.Items.Item("crdActCuCB");
                        oItem.Visible = false;
                    }
                    else
                    {
                        oItem = oForm.Items.Item("creditActS");
                        oItem.Visible = false;
                        oItem = oForm.Items.Item("creditActE");
                        oItem.Visible = false;
                        oItem = oForm.Items.Item("crdActCuCB");
                        oItem.Visible = false;
                        oItem = oForm.Items.Item("tresrCodeS");
                        oItem.Visible = false;
                        oItem = oForm.Items.Item("tresrCodeE");
                        oItem.Visible = false;
                        oItem = oForm.Items.Item("chrgDtlsS");
                        oItem.Visible = false;
                        oItem = oForm.Items.Item("chrgDtlsCB");
                        oItem.Visible = false;
                        oItem = oForm.Items.Item("dsptTypeS");
                        oItem.Visible = false;
                        oItem = oForm.Items.Item("dsptTypeCB");
                        oItem.Visible = false;
                    }
                }
                else if (transferType == "TransferToOtherBankForeignCurrencyPaymentOrderIo" || transferType == "TransferToForeignCurrencyPaymentOrderIo")
                {
                    oItem = oForm.Items.Item("chrgDtlsS");
                    oItem.Visible = true;
                    oItem = oForm.Items.Item("chrgDtlsCB");
                    oItem.Visible = true;
                    oItem = oForm.Items.Item("dsptTypeS");
                    oItem.Visible = false;
                    oItem = oForm.Items.Item("dsptTypeCB");
                    oItem.Visible = false;
                }
                else if (transferType == "TransferToNationalCurrencyPaymentOrderIo" || transferType == "TreasuryTransferPaymentOrderIoBP")
                {
                    oItem = oForm.Items.Item("chrgDtlsS");
                    oItem.Visible = false;
                    oItem = oForm.Items.Item("chrgDtlsCB");
                    oItem.Visible = false;
                    oItem = oForm.Items.Item("dsptTypeS");
                    oItem.Visible = transferType == "TransferToNationalCurrencyPaymentOrderIo";
                    oItem = oForm.Items.Item("dsptTypeCB");
                    oItem.Visible = transferType == "TransferToNationalCurrencyPaymentOrderIo";
                }
                else
                {
                    oItem = oForm.Items.Item("opTypeS");
                    oItem.Visible = false;
                    oItem = oForm.Items.Item("opTypeCB");
                    oItem.Visible = false;
                    oItem = oForm.Items.Item("creditActS");
                    oItem.Visible = false;
                    oItem = oForm.Items.Item("creditActE");
                    oItem.Visible = false;
                    oItem = oForm.Items.Item("crdActCuCB");
                    oItem.Visible = false;
                    oItem = oForm.Items.Item("tresrCodeS");
                    oItem.Visible = false;
                    oItem = oForm.Items.Item("tresrCodeE");
                    oItem.Visible = false;
                    oItem = oForm.Items.Item("chrgDtlsS");
                    oItem.Visible = false;
                    oItem = oForm.Items.Item("chrgDtlsCB");
                    oItem.Visible = false;
                    oItem = oForm.Items.Item("dsptTypeS");
                    oItem.Visible = false;
                    oItem = oForm.Items.Item("dsptTypeCB");
                    oItem.Visible = false;
                }
            }
            catch (Exception ex)
            {
            }
            finally
            {
                oForm.Freeze(false);
                oForm.Update();
            }
        }

        public static void chooseFromList(SAPbouiCOM.Form oForm, SAPbouiCOM.IChooseFromListEvent oCFLEvento, string itemUID, bool beforeAction)
        {
            NumberFormatInfo Nfi = new NumberFormatInfo() { NumberDecimalSeparator = "." };

            try
            {
                string sCFL_ID = oCFLEvento.ChooseFromListUID;
                SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(sCFL_ID);

                if (beforeAction)
                {
                    if (sCFL_ID == "HouseBankAccount_CFL")
                    {

                    }
                    else if (sCFL_ID == "1") //Blanket Agreement
                    {
                        string project = oForm.DataSources.DBDataSources.Item("OVPM").GetValue("PrjCode", 0);
                        SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();
                        if (project != "")
                        {
                            SAPbouiCOM.Condition oCon = oCons.Add();
                            oCon.Alias = "Project";
                            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                            oCon.CondVal = project;
                        }
                        oCFL.SetConditions(oCons);
                    }
                }
                else
                {
                    SAPbouiCOM.DataTable oDataTable = null;
                    oDataTable = oCFLEvento.SelectedObjects;

                    if (oDataTable != null)
                    {
                        if (sCFL_ID == "HouseBankAccount_CFL")
                        {
                            string account = Convert.ToString(oDataTable.GetValue("Account", 0));
                            string currency;
                            SAPbouiCOM.EditText oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("creditActE").Specific;
                            oEditText.Value = account;
                            //oEditText.Value = CommonFunctions.accountParse(account, out currency);
                            try
                            {
                                SAPbouiCOM.ComboBox oComboBox = (SAPbouiCOM.ComboBox)oForm.Items.Item("crdActCuCB").Specific;
                                CommonFunctions.accountParse(account, out currency);
                                oComboBox.Select(currency, SAPbouiCOM.BoSearchKey.psk_ByValue);
                            }
                            catch { }
                            //if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            //{
                            //    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                            //}
                        }

                        else if (sCFL_ID == "CFL_ProfitBase")
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

                        else if (sCFL_ID == "1") //Blanket Agreement
                        {
                            var agrNo = Convert.ToString(oDataTable.GetValue("AbsID", 0));
                            var prjCode = Convert.ToString(oDataTable.GetValue("Project", 0));
                            var bpCurr = Convert.ToString(oDataTable.GetValue("BPCurr", 0));

                            LanguageUtils.IgnoreErrors<string>(() => oForm.Items.Item("95").Specific.Value = prjCode);

                            FilterInvoiceMatrix(oForm, agrNo, prjCode);

                            oForm.Items.Item("26").Click(); //Remark

                            oForm.Items.Item("95").Enabled = string.IsNullOrEmpty(prjCode);
                            oForm.Items.Item("234000005").Enabled = string.IsNullOrEmpty(agrNo);

                            IncomingPayment.SetUsBlaAgRtSAvailability(oForm, !string.IsNullOrEmpty(agrNo) && bpCurr != Program.LocalCurrency);
                        }

                        else if (sCFL_ID == "23") //Project
                        {
                            var prjCode = Convert.ToString(oDataTable.GetValue("PrjCode", 0));

                            oForm.Items.Item("234000005").Enabled = true;
                            LanguageUtils.IgnoreErrors<string>(() => oForm.Items.Item("234000005").Specific.Value = string.Empty);

                            FilterInvoiceMatrix(oForm, null, prjCode);

                            oForm.Items.Item("26").Click(); //Remark
                            oForm.Items.Item("95").Enabled = string.IsNullOrEmpty(prjCode);

                            IncomingPayment.SetUsBlaAgRtSAvailability(oForm);
                        }
                    }
                    setVisibleFormItems(oForm);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static void resizeForm(SAPbouiCOM.Form oForm)
        {
            reArrangeFormItems(oForm);
        }

        public static void reArrangeFormItems(SAPbouiCOM.Form oForm)
        {
            oForm.Freeze(true);
            try
            {
                SAPbouiCOM.Item oItem = oForm.Items.Item("20");

                int height = 15;
                int top = oItem.Top + oItem.Height;

                oItem = oForm.Items.Item("descrptS");
                oItem.Top = top;
                oItem = oForm.Items.Item("descrptE");
                oItem.Top = top;

                top = top + height + 1;

                oItem = oForm.Items.Item("addDescrpS");
                oItem.Top = top;
                oItem = oForm.Items.Item("addDescrpE");
                oItem.Top = top;

                //მოგების გადასახადი
                top = top + height + 1;
                oItem = oForm.Items.Item("liablePrTx");
                oItem.Top = top;

                top = top + height + 1;
                oItem = oForm.Items.Item("PrBaseS");
                oItem.Top = top;
                oItem = oForm.Items.Item("PrBaseE");
                oItem.Top = top;
                oItem = oForm.Items.Item("PrBsDscr");
                oItem.Top = top;

                top = top + height + 1;
                oItem = oForm.Items.Item("27");
                oItem.Top = top;
                oItem = oForm.Items.Item("26");
                oItem.Top = top;

                top = top + height + 1;
                oItem = oForm.Items.Item("60");
                oItem.Top = top;
                oItem = oForm.Items.Item("59");
                oItem.Top = top;

                oItem = oForm.Items.Item("13");
                top = oItem.Top + oItem.Height + 1;
                oItem = oForm.Items.Item("AmtPrTxS");
                oItem.Top = top;
                oItem.Left = oForm.Items.Item("37").Left;
                oItem = oForm.Items.Item("AmtPrTxE");
                oItem.Top = top;
                oItem.Left = oForm.Items.Item("13").Left;
                oItem = oForm.Items.Item("FillAmtTxs");
                oItem.Top = top;
                oItem.Left = oForm.Items.Item("234000001").Left;
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

        private static Dictionary<string, string> getDataForTransferType(SAPbouiCOM.Form oForm)
        {
            try
            {
                //დოკუმენტის მონაცემები --->
                string docType = oForm.DataSources.DBDataSources.Item("OVPM").GetValue("DocType", 0).Trim();
                string opType = oForm.DataSources.DBDataSources.Item("OVPM").GetValue("U_opType", 0).Trim();
                string diffCurr = oForm.DataSources.DBDataSources.Item("OVPM").GetValue("DiffCurr", 0).Trim();
                //string docCurr = oForm.DataSources.DBDataSources.Item("OVPM").GetValue("DocCurr", 0).Trim();
                string docCurr = CommonFunctions.getAccountCurrency(oForm.DataSources.DBDataSources.Item("OVPM").GetValue("TrsfrAcct", 0).Trim());
                docCurr = docCurr == "##" ? Program.LocalCurrency : docCurr;
                //docCurr = CommonFunctions.getCurrencyInternationalCode(docCurr);
                //docCurr = diffCurr == "Y" ? Program.LocalCurrency : docCurr;
                string description = oForm.DataSources.DBDataSources.Item("OVPM").GetValue("U_descrpt", 0).Trim();
                string chargeDetails = oForm.DataSources.DBDataSources.Item("OVPM").GetValue("U_chrgDtls", 0).Trim();
                string reportCode = oForm.DataSources.DBDataSources.Item("OVPM").GetValue("U_rprtCode", 0).Trim();
                string dispatchType = oForm.DataSources.DBDataSources.Item("OVPM").GetValue("U_dsptchType", 0).Trim();
                string isPayToBank = oForm.DataSources.DBDataSources.Item("OVPM").GetValue("IsPaytoBnk", 0).Trim();
                string docRate = FormsB1.ConvertDecimalToString(FormsB1.cleanStringOfNonDigits(oForm.DataSources.DBDataSources.Item("OVPM").GetValue("docRate", 0).ToString()));
                //დოკუმენტის მონაცემები <---

                //კომპანიის მონაცემები (გამგზავნი) --->
                string trsfrAcct = oForm.DataSources.DBDataSources.Item("OVPM").GetValue("TrsfrAcct", 0).Trim(); // ბუღ. ანგარიში (გამგზავნი)
                string bankProgram = CommonFunctions.getBankProgram(trsfrAcct); //ბანკის პროგრამა (გამგზავნი)
                string bankCode = CommonFunctions.getBankCode(trsfrAcct); //ბანკის კოდი (გამგზავნი)
                //კომპანიის მონაცემები (გამგზავნი) <---

                //კომპანიის/თანამშრომლის მონაცემები (მიმღები) --->
                string creditAcct = null;
                string creditBankCode = null;
                string creditAcctCurrency = null;
                string exchangeCurrency = null;
                if (opType == "paymentToEmployee")
                {
                    try
                    {
                        SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("71").Specific));
                        creditAcct = oMatrix.Columns.Item("U_creditAcct").Cells.Item(1).Specific.Value;
                        creditBankCode = oMatrix.Columns.Item("U_bankCode").Cells.Item(1).Specific.Value;
                        creditAcctCurrency = docCurr; //ანგარიშის ვალუტა (მიმღები)
                    }
                    catch { }
                    if (string.IsNullOrEmpty(creditAcct))
                    {
                        creditAcct = oForm.DataSources.DBDataSources.Item("VPM4").GetValue("U_creditAcct", 0).Trim(); //ანგარიში (მიმღები)
                        creditBankCode = oForm.DataSources.DBDataSources.Item("VPM4").GetValue("U_bankCode", 0).Trim(); //CommonFunctions.getBankCode( null, creditAcct); //ბანკის კოდი (მიმღები)
                        creditAcctCurrency = docCurr; //ანგარიშის ვალუტა (მიმღები)
                    }
                }
                else
                {
                    creditAcct = oForm.DataSources.DBDataSources.Item("OVPM").GetValue("U_creditAcct", 0).Trim(); //ანგარიში (მიმღები)
                    creditBankCode = CommonFunctions.getBankCode(null, creditAcct); //ბანკის კოდი (მიმღები)
                    creditAcctCurrency = null; //ანგარიშის ვალუტა (მიმღები)
                    if (string.IsNullOrEmpty(creditAcct) == false)
                    {
                        creditAcct = CommonFunctions.accountParse(creditAcct, out creditAcctCurrency);
                    }
                    exchangeCurrency = oForm.DataSources.DBDataSources.Item("OVPM").GetValue("U_crdtActCur", 0).Trim(); //კონვერტაციის ვალუტა
                    exchangeCurrency = CommonFunctions.getCurrencyInternationalCode(exchangeCurrency);
                }
                //კომპანიის/თანამშრომლის მონაცემები (მიმღები) <---

                //კონტრაგენტის მონაცემები (მიმღები) --->
                string bpBnkCode = oForm.DataSources.DBDataSources.Item("OVPM").GetValue("PBnkCode", 0).Trim(); //ბანკის კოდი
                string bpBAccount = oForm.DataSources.DBDataSources.Item("OVPM").GetValue("PBnkAccnt", 0).Trim(); //ანგარიში
                string bpBAccountCurrency = null; //ვალუტა
                if (string.IsNullOrEmpty(bpBAccount) == false)
                {
                    bpBAccount = CommonFunctions.accountParse(bpBAccount, out bpBAccountCurrency);
                }
                string RecipientCity = oForm.DataSources.DBDataSources.Item("OCRD").GetValue("City", 0).Trim();
                string BeneficiaryRegistrationCountryCode = oForm.DataSources.DBDataSources.Item("OCRD").GetValue("Country", 0).Trim();
                string BeneficiaryAddress = oForm.DataSources.DBDataSources.Item("OCRD").GetValue("Address", 0).Trim();
                string treasuryCode = oForm.DataSources.DBDataSources.Item("OVPM").GetValue("U_tresrCode", 0).Trim();
                string cardCode = oForm.DataSources.DBDataSources.Item("OVPM").GetValue("CardCode", 0).Trim();
                string isBPAccountTreasury = CommonFunctions.isBPAccountTreasury(cardCode, bpBnkCode, oForm.DataSources.DBDataSources.Item("OVPM").GetValue("PBnkAccnt", 0).Trim()) ? "Y" : "N";
                treasuryCode = isBPAccountTreasury == "Y" ? bpBAccount : treasuryCode;
                //კონტრაგენტის მონაცემები (მიმღები) <---

                Dictionary<string, string> dataForTransferType = new Dictionary<string, string>();
                dataForTransferType.Add("docType", docType);
                dataForTransferType.Add("opType", opType);
                dataForTransferType.Add("diffCurr", diffCurr);
                dataForTransferType.Add("docCurr", docCurr);
                dataForTransferType.Add("description", description);
                dataForTransferType.Add("chargeDetails", chargeDetails);
                dataForTransferType.Add("reportCode", reportCode);
                dataForTransferType.Add("RecipientCity", RecipientCity);
                dataForTransferType.Add("BeneficiaryAddress", BeneficiaryAddress);
                dataForTransferType.Add("BeneficiaryRegistrationCountryCode", BeneficiaryRegistrationCountryCode);
                dataForTransferType.Add("dispatchType", dispatchType);
                dataForTransferType.Add("isPayToBank", isPayToBank);
                dataForTransferType.Add("docRate", docRate);

                dataForTransferType.Add("trsfrAcct", trsfrAcct);
                dataForTransferType.Add("bankCode", bankCode);
                dataForTransferType.Add("bankProgram", bankProgram);
                dataForTransferType.Add("creditAcct", creditAcct);
                dataForTransferType.Add("creditBankCode", creditBankCode);
                dataForTransferType.Add("creditAcctCurrency", creditAcctCurrency);
                dataForTransferType.Add("exchangeCurrency", exchangeCurrency);

                dataForTransferType.Add("bpBnkCode", bpBnkCode);
                dataForTransferType.Add("bpBAccount", bpBAccount);
                dataForTransferType.Add("bpBAccountCurrency", bpBAccountCurrency);
                dataForTransferType.Add("treasuryCode", treasuryCode);
                dataForTransferType.Add("isBPAccountTreasury", isBPAccountTreasury);

                return dataForTransferType;
            }
            catch
            {
                return null;
            }
        }

        public static Dictionary<string, string> getDataForTransferType(SAPbobsCOM.Recordset oRecordSet)
        {
            try
            {
                //დოკუმენტის მონაცემები --->
                string docType = oRecordSet.Fields.Item("DocType").Value.ToString();
                string opType = oRecordSet.Fields.Item("U_opType").Value.ToString();
                string diffCurr = oRecordSet.Fields.Item("DiffCurr").Value.ToString();
                //string docCurr = oRecordSet.Fields.Item("DocCurr").Value.ToString();
                string docCurr = oRecordSet.Fields.Item("DebitAccountCurrencyCode").Value.ToString();
                //docCurr = CommonFunctions.getCurrencyInternationalCode(docCurr);
                //docCurr = diffCurr == "Y" ? Program.LocalCurrency : docCurr;
                string description = oRecordSet.Fields.Item("U_descrpt").Value.ToString();
                string chargeDetails = oRecordSet.Fields.Item("U_chrgDtls").Value.ToString();
                string reportCode = oRecordSet.Fields.Item("U_rprtCode").Value.ToString();
                string dispatchType = oRecordSet.Fields.Item("U_dsptchType").Value.ToString();
                string isPayToBank = oRecordSet.Fields.Item("IsPaytoBnk").Value.ToString();
                //დოკუმენტის მონაცემები <---

                //კომპანიის მონაცემები (გამგზავნი) --->
                string trsfrAcct = oRecordSet.Fields.Item("TrsfrAcct").Value.ToString(); // ბუღ. ანგარიში (გამგზავნი)
                string bankProgram = CommonFunctions.getBankProgram(trsfrAcct); //ბანკის პროგრამა (გამგზავნი)
                string bankCode = CommonFunctions.getBankCode(trsfrAcct); //ბანკის კოდი (გამგზავნი)
                //კომპანიის მონაცემები (გამგზავნი) <---

                //კომპანიის/თანამშრომლის მონაცემები (მიმღები) --->
                string creditAcct = null;
                string creditBankCode = null;
                string creditAcctCurrency = null;
                string exchangeCurrency = null;
                if (opType == "paymentToEmployee")
                {
                    creditAcct = oRecordSet.Fields.Item("U_creditAcctEmp").Value.ToString(); //ანგარიში (მიმღები)
                    creditBankCode = oRecordSet.Fields.Item("U_bankCodeEmp").Value.ToString();//CommonFunctions.getBankCode( null, creditAcct); //ბანკის კოდი (მიმღები)
                    creditAcctCurrency = docCurr; //ანგარიშის ვალუტა (მიმღები)
                }
                else
                {
                    creditAcct = oRecordSet.Fields.Item("U_creditAcct").Value.ToString(); //ანგარიში (მიმღები)
                    creditBankCode = CommonFunctions.getBankCode(null, creditAcct); //ბანკის კოდი (მიმღები)
                    creditAcctCurrency = null; //ანგარიშის ვალუტა (მიმღები)
                    if (string.IsNullOrEmpty(creditAcct) == false)
                    {
                        creditAcct = CommonFunctions.accountParse(creditAcct, out creditAcctCurrency);
                    }
                    exchangeCurrency = oRecordSet.Fields.Item("U_crdtActCur").Value.ToString(); //კონვერტაციის ვალუტა
                    exchangeCurrency = CommonFunctions.getCurrencyInternationalCode(exchangeCurrency);
                }
                //კომპანიის/თანამშრომლის მონაცემები (მიმღები) <---

                //კონტრაგენტის მონაცემები (მიმღები) --->
                string bpBnkCode = oRecordSet.Fields.Item("PBnkCode").Value.ToString(); //ბანკის კოდი
                string bpBAccount = oRecordSet.Fields.Item("PBnkAccnt").Value.ToString(); //ანგარიში
                string bpBAccountCurrency = null; //ვალუტა
                if (string.IsNullOrEmpty(bpBAccount) == false)
                {
                    bpBAccount = CommonFunctions.accountParse(bpBAccount, out bpBAccountCurrency);
                }
                string RecipientCity = oRecordSet.Fields.Item("RecipientCity").Value.ToString();
                string BeneficiaryRegistrationCountryCode = oRecordSet.Fields.Item("BeneficiaryRegistrationCountryCode").Value.ToString();
                string BeneficiaryAddress = oRecordSet.Fields.Item("BeneficiaryAddress").Value.ToString();
                string treasuryCode = oRecordSet.Fields.Item("U_tresrCode").Value.ToString();
                string cardCode = oRecordSet.Fields.Item("CardCode").Value.ToString();
                string isBPAccountTreasury = CommonFunctions.isBPAccountTreasury(cardCode, bpBnkCode, oRecordSet.Fields.Item("PBnkAccnt").Value.ToString()) ? "Y" : "N";
                treasuryCode = isBPAccountTreasury == "Y" ? bpBAccount : treasuryCode;
                //კონტრაგენტის მონაცემები (მიმღები) <---

                Dictionary<string, string> dataForTransferType = new Dictionary<string, string>();
                dataForTransferType.Add("docType", docType);
                dataForTransferType.Add("opType", opType);
                dataForTransferType.Add("diffCurr", diffCurr);
                dataForTransferType.Add("docCurr", docCurr);
                dataForTransferType.Add("description", description);
                dataForTransferType.Add("chargeDetails", chargeDetails);
                dataForTransferType.Add("reportCode", reportCode);
                dataForTransferType.Add("RecipientCity", RecipientCity);
                dataForTransferType.Add("BeneficiaryAddress", BeneficiaryAddress);
                dataForTransferType.Add("BeneficiaryRegistrationCountryCode", BeneficiaryRegistrationCountryCode);
                dataForTransferType.Add("dispatchType", dispatchType);
                dataForTransferType.Add("isPayToBank", isPayToBank);
                //dataForTransferType.Add("docRate", docRate);

                dataForTransferType.Add("trsfrAcct", trsfrAcct);
                dataForTransferType.Add("bankCode", bankCode);
                dataForTransferType.Add("bankProgram", bankProgram);
                dataForTransferType.Add("creditAcct", creditAcct);
                dataForTransferType.Add("creditBankCode", creditBankCode);
                dataForTransferType.Add("creditAcctCurrency", creditAcctCurrency);
                dataForTransferType.Add("exchangeCurrency", exchangeCurrency);

                dataForTransferType.Add("bpBnkCode", bpBnkCode);
                dataForTransferType.Add("bpBAccount", bpBAccount);
                dataForTransferType.Add("bpBAccountCurrency", bpBAccountCurrency);
                dataForTransferType.Add("treasuryCode", treasuryCode);
                dataForTransferType.Add("isBPAccountTreasury", isBPAccountTreasury);

                return dataForTransferType;
            }
            catch
            {
                return null;
            }
        }

        public static Dictionary<string, object> getDataForImport(SAPbobsCOM.Recordset oRecordSet, Dictionary<string, string> dataForTransferType, string transferType)
        {
            try
            {
                Dictionary<string, object> dataForImport = new Dictionary<string, object>();

                //დოკუმენტის მონაცემები --->
                string docType = dataForTransferType["docType"];
                string opType = dataForTransferType["opType"];
                string diffCurr = dataForTransferType["diffCurr"];
                string docCurr = oRecordSet.Fields.Item("DebitAccountCurrencyCode").Value.ToString(); //dataForTransferType["docCurr"];
                string description = dataForTransferType["description"];
                string chargeDetails = dataForTransferType["chargeDetails"];
                string reportCode = dataForTransferType["reportCode"];
                string dispatchType = dataForTransferType["dispatchType"];
                string isPayToBank = dataForTransferType["isPayToBank"];
                //დოკუმენტის მონაცემები <---

                //კომპანიის მონაცემები (გამგზავნი) --->
                string trsfrAcct = dataForTransferType["trsfrAcct"]; // ბუღ. ანგარიში (გამგზავნი)
                string bankCode = dataForTransferType["bankCode"]; //ბანკის კოდი (გამგზავნი)
                string bankProgram = dataForTransferType["bankProgram"]; //ბანკის პროგრამა (გამგზავნი)
                //კომპანიის მონაცემები (გამგზავნი) <---

                //კომპანიის/თანამშრომლის მონაცემები (მიმღები) --->
                string creditAcct = dataForTransferType["creditAcct"]; //ანგარიში (მიმღები)
                string creditBankCode = dataForTransferType["creditBankCode"]; //ბანკის კოდი (მიმღები)
                string creditAcctCurrency = dataForTransferType["creditAcctCurrency"]; //ანგარიშის ვალუტა (მიმღები)
                string exchangeCurrency = dataForTransferType["exchangeCurrency"]; //კონვერტაციის ვალუტა
                //კომპანიის/თანამშრომლის მონაცემები (მიმღები) <---

                //კონტრაგენტის მონაცემები (მიმღები) --->
                string bpBnkCode = dataForTransferType["bpBnkCode"]; //ბანკის კოდი
                string bpBAccount = dataForTransferType["bpBAccount"]; //ანგარიში
                string bpBAccountCurrency = dataForTransferType["bpBAccountCurrency"]; //ვალუტა
                string treasuryCode = dataForTransferType["treasuryCode"];
                //კონტრაგენტის მონაცემები (მიმღები) <---

                string CardCode = oRecordSet.Fields.Item("CardCode").Value.ToString();
                string CreditAccount = bpBAccount;
                string CreditAccountCurrencyCode = bpBAccountCurrency;
                string BeneficiaryName = oRecordSet.Fields.Item("CardName").Value.ToString();
                string BeneficiaryTaxCode = oRecordSet.Fields.Item("BeneficiaryTaxCode").Value.ToString();
                string BeneficiaryAddress = oRecordSet.Fields.Item("BeneficiaryAddress").Value.ToString();
                string RecipientCity = oRecordSet.Fields.Item("RecipientCity").Value.ToString();
                string BeneficiaryBankCode = bpBnkCode;
                string BeneficiaryBankName = oRecordSet.Fields.Item("BeneficiaryBankName").Value.ToString();

                string TransferCurrency = docCurr;
                string BeneficiaryRegistrationCountryCode = oRecordSet.Fields.Item("BeneficiaryRegistrationCountryCode").Value.ToString();

                if (opType == "paymentToEmployee")
                {
                    CardCode = oRecordSet.Fields.Item("U_employee").Value.ToString();
                    CreditAccount = creditAcct;
                    CreditAccountCurrencyCode = creditAcctCurrency;
                    BeneficiaryName = oRecordSet.Fields.Item("U_employeeN").Value.ToString();
                    BeneficiaryTaxCode = oRecordSet.Fields.Item("BeneficiaryTaxCodeEmp").Value.ToString();
                    BeneficiaryAddress = oRecordSet.Fields.Item("BeneficiaryAddressEmp").Value.ToString();
                    BeneficiaryBankCode = creditBankCode;
                    BeneficiaryBankName = oRecordSet.Fields.Item("BeneficiaryBankNameEmp").Value.ToString();
                    BeneficiaryRegistrationCountryCode = oRecordSet.Fields.Item("BeneficiaryRegistrationCountryCodeEmp").Value.ToString();
                }
                else
                {
                    Dictionary<string, string> CompanyInfo = CommonFunctions.getCompanyInfo();

                    if (bankProgram == "TBC")
                    {
                        if (transferType == "TransferToOwnAccountPaymentOrderIo" || transferType == "CurrencyExchangePaymentOrderIo" || (docType == "A" && (transferType == "TransferToOtherBankNationalCurrencyPaymentOrderIo" || transferType == "TransferToOtherBankForeignCurrencyPaymentOrderIo")))
                        {
                            CreditAccount = creditAcct;
                            CreditAccountCurrencyCode = creditAcctCurrency;
                            BeneficiaryName = CompanyInfo["CompnyName"];
                            BeneficiaryTaxCode = CompanyInfo["FreeZoneNo"];
                            BeneficiaryAddress = CompanyInfo["CompnyAddr"];
                            BeneficiaryBankCode = creditBankCode;
                            BeneficiaryBankName = CommonFunctions.getBankName(creditBankCode);
                        }
                    }
                    else if (bankProgram == "BOG")
                    {
                        if (transferType == "TransferToOwnAccountPaymentOrderIo" || transferType == "CurrencyExchangePaymentOrderIo" || (docType == "A" && (transferType == "TransferToNationalCurrencyPaymentOrderIo" || transferType == "TransferToForeignCurrencyPaymentOrderIo")))
                        {
                            CreditAccount = creditAcct;
                            CreditAccountCurrencyCode = creditAcctCurrency;
                            BeneficiaryName = CompanyInfo["CompnyName"];
                            BeneficiaryTaxCode = CompanyInfo["FreeZoneNo"];
                            BeneficiaryAddress = CompanyInfo["CompnyAddr"];
                            BeneficiaryBankCode = creditBankCode;
                            BeneficiaryBankName = CommonFunctions.getBankName(creditBankCode);
                        }
                    }
                    if (transferType == "TreasuryTransferPaymentOrderIo" || transferType == "TreasuryTransferPaymentOrderIoBP") //სახაზინო გადარიცხვა
                    {
                        CreditAccount = treasuryCode;
                        BeneficiaryBankCode = "TRESGE22"; //მიმღები ბანკის RTGS კოდი / სავალდებულო
                        BeneficiaryName = "სახელმწიფო ხაზინა"; //მიმღების დასახელება
                    }
                }

                if (docType == "A" && opType != "paymentToEmployee")
                {
                    CardCode = "";
                }

                var Amount = decimal.Zero;
                var AmountFC = decimal.Zero;
                var AmountLC = Math.Round(Convert.ToDecimal(oRecordSet.Fields.Item("TrsfrSum").Value), 2);
                if (TransferCurrency == Program.LocalCurrency)
                    Amount = Convert.ToDecimal(oRecordSet.Fields.Item("TrsfrSum").Value);
                else
                {
                    Amount = Convert.ToDecimal(oRecordSet.Fields.Item("TrsfrSumFC").Value);
                    AmountFC = Math.Round(Convert.ToDecimal(oRecordSet.Fields.Item("TrsfrSumFC").Value), 2);
                }
                Amount = Math.Round(Amount, 2); //დამრგვალება აუცილებლად უნდა იყოს 2 ციფრამდე, ინტ.ბანკის გამო

                dataForImport.Add("DebitBankCode", bankCode);
                dataForImport.Add("CreditAccount", CreditAccount);
                dataForImport.Add("CreditAccountCurrencyCode", CreditAccountCurrencyCode);
                dataForImport.Add("Currency", TransferCurrency);
                dataForImport.Add("BeneficiaryName", BeneficiaryName);
                dataForImport.Add("Amount", Amount);
                dataForImport.Add("AmountLC", AmountLC);
                dataForImport.Add("AmountFC", AmountFC);
                dataForImport.Add("BeneficiaryTaxCode", BeneficiaryTaxCode);
                dataForImport.Add("BeneficiaryAddress", BeneficiaryAddress);
                dataForImport.Add("RecipientCity", RecipientCity);
                dataForImport.Add("BeneficiaryBankCode", BeneficiaryBankCode);
                dataForImport.Add("BeneficiaryBankName", BeneficiaryBankName);
                dataForImport.Add("BeneficiaryRegistrationCountryCode", BeneficiaryRegistrationCountryCode);

                string currencyTemp;
                dataForImport.Add("DebitAccount", CommonFunctions.accountParse(oRecordSet.Fields.Item("DebitAccount").Value, out currencyTemp));
                dataForImport.Add("DebitAccountCurrencyCode", oRecordSet.Fields.Item("DebitAccountCurrencyCode").Value.ToString());

                dataForImport.Add("IntermediaryBankCode", oRecordSet.Fields.Item("IntermediaryBankCode").Value.ToString());
                dataForImport.Add("IntermediaryBankName", oRecordSet.Fields.Item("IntermediaryBankName").Value.ToString());
                dataForImport.Add("ChargeDetails", chargeDetails);
                dataForImport.Add("reportCode", reportCode);
                dataForImport.Add("TaxpayerCode", oRecordSet.Fields.Item("TaxpayerCode").Value.ToString());
                dataForImport.Add("TaxpayerName", oRecordSet.Fields.Item("TaxpayerName").Value.ToString());
                dataForImport.Add("TreasuryCode", treasuryCode);
                dataForImport.Add("DocEntry", oRecordSet.Fields.Item("DocEntry").Value.ToString());
                dataForImport.Add("AdditionalDescription", oRecordSet.Fields.Item("U_addDescrpt").Value.ToString());
                dataForImport.Add("Description", description);
                dataForImport.Add("CardCode", CardCode);
                dataForImport.Add("DispatchType", dispatchType);
                dataForImport.Add("TransferType", transferType);
                dataForImport.Add("DocRate", Convert.ToDecimal(oRecordSet.Fields.Item("DocRate").Value));

                return dataForImport;
            }
            catch
            {
                return null;
            }
        }

        public static string getTransferType(Dictionary<string, string> dataForTransferType, out string errorText)
        {
            errorText = null;
            string transferType = null;

            if (dataForTransferType == null)
            {
                return null;
            }

            //დოკუმენტის მონაცემები --->
            string docType = dataForTransferType["docType"] == null ? "" : dataForTransferType["docType"];
            string opType = dataForTransferType["opType"] == null ? "" : dataForTransferType["opType"];
            string diffCurr = dataForTransferType["diffCurr"] == null ? "" : dataForTransferType["diffCurr"];
            string docCurr = dataForTransferType["docCurr"] == null ? "" : dataForTransferType["docCurr"];
            string description = dataForTransferType["description"] == null ? "" : dataForTransferType["description"];
            string chargeDetails = dataForTransferType["chargeDetails"] == null ? "" : dataForTransferType["chargeDetails"];
            string isPayToBank = dataForTransferType["isPayToBank"] == null ? "" : dataForTransferType["isPayToBank"];
            //დოკუმენტის მონაცემები <---

            //კომპანიის მონაცემები (გამგზავნი) --->
            string trsfrAcct = dataForTransferType["trsfrAcct"] == null ? "" : dataForTransferType["trsfrAcct"]; // ბუღ. ანგარიში (გამგზავნი)
            string bankCode = dataForTransferType["bankCode"] == null ? "" : dataForTransferType["bankCode"]; //ბანკის კოდი (გამგზავნი)
            string bankProgram = dataForTransferType["bankProgram"] == null ? "" : dataForTransferType["bankProgram"]; //ბანკის პროგრამა (გამგზავნი)
            //კომპანიის მონაცემები (გამგზავნი) <---

            //კომპანიის/თანამშრომლის მონაცემები (მიმღები) --->
            string creditAcct = dataForTransferType["creditAcct"] == null ? "" : dataForTransferType["creditAcct"]; //ანგარიში (მიმღები)
            string creditBankCode = dataForTransferType["creditBankCode"] == null ? "" : dataForTransferType["creditBankCode"]; //ბანკის კოდი (მიმღები)
            string creditAcctCurrency = dataForTransferType["creditAcctCurrency"] == null ? "" : dataForTransferType["creditAcctCurrency"]; //ანგარიშის ვალუტა (მიმღები)
            string exchangeCurrency = dataForTransferType["exchangeCurrency"] == null ? "" : dataForTransferType["exchangeCurrency"]; //კონვერტაციის ვალუტა
            //კომპანიის/თანამშრომლის მონაცემები (მიმღები) <---

            //კონტრაგენტის მონაცემები (მიმღები) --->
            string bpBnkCode = dataForTransferType["bpBnkCode"] == null ? "" : dataForTransferType["bpBnkCode"]; //ბანკის კოდი
            string bpBAccount = dataForTransferType["bpBAccount"] == null ? "" : dataForTransferType["bpBAccount"]; //ანგარიში
            string bpBAccountCurrency = dataForTransferType["bpBAccountCurrency"] == null ? "" : dataForTransferType["bpBAccountCurrency"]; //ვალუტა
            string treasuryCode = dataForTransferType["treasuryCode"] == null ? "" : dataForTransferType["treasuryCode"];
            string isBPAccountTreasury = dataForTransferType["isBPAccountTreasury"] == null ? "" : dataForTransferType["isBPAccountTreasury"]; //ბპ-ის ანგარიში არის თუ არა სახაზინო
            //კონტრაგენტის მონაცემები (მიმღები) <---

            if (string.IsNullOrEmpty(bankCode) || string.IsNullOrEmpty(trsfrAcct))
            {
                return null;
            }

            if (docType == "A" && opType == "treasuryTransfer")
            {
                transferType = "TreasuryTransferPaymentOrderIo"; //საბიუჯეტო გადარიცხვა
            }
            else if (docType == "A" && (opType == "transferToOwnAccount" || opType == "currencyExchange" || opType == "paymentToEmployee") && string.IsNullOrEmpty(creditBankCode) == false)
            {
                if (opType == "currencyExchange" && creditBankCode == bankCode) //ერთნაირი ბანკის ანგარიშებია
                {
                    transferType = "CurrencyExchangePaymentOrderIo"; //კონვერტაცია
                }
                else if (opType == "transferToOwnAccount" || opType == "paymentToEmployee")
                {
                    if (creditBankCode == bankCode && bankCode == "TBCBGE22" && opType == "transferToOwnAccount") //ერთნაირი ბანკის ანგარიშებია  და TBC
                    {
                        transferType = "TransferToOwnAccountPaymentOrderIo"; //გადარიცხვა საკუთარ ანგარიშზე
                    }
                    else
                    {
                        //docCurr <-> creditAcctCurrency
                        if (bankCode == "TBCBGE22")
                        {
                            if (creditBankCode == bankCode)
                            {
                                transferType = "TransferWithinBankPaymentOrderIo"; //გადარიცხვა თიბისი ბანკის ფილიალებში
                            }
                            else if (docCurr == "GEL")
                            {
                                transferType = "TransferToOtherBankNationalCurrencyPaymentOrderIo"; //გადარიცხვა სხვა ბანკში (ეროვნული ვალუტა)
                            }
                            else if (docCurr != "GEL")
                            {
                                transferType = "TransferToOtherBankForeignCurrencyPaymentOrderIo"; //გადარიცხვა სხვა ბანკში (უცხოური ვალუტა)
                            }
                        }
                        else if (bankCode == "BAGAGE22")
                        {
                            if (docCurr == "GEL")
                            {
                                transferType = "TransferToNationalCurrencyPaymentOrderIo"; //გადარიცხვა (ეროვნული ვალუტა)
                            }
                            else if (docCurr != "GEL")
                            {
                                transferType = "TransferToForeignCurrencyPaymentOrderIo"; //გადარიცხვა (უცხოური ვალუტა)
                            }
                        }
                    }
                }
            }
            else if (docType != "A")
            {
                if (isBPAccountTreasury == "Y")
                {
                    transferType = "TreasuryTransferPaymentOrderIoBP"; //საბიუჯეტო გადარიცხვა ბპ-თვის
                }
                else
                {
                    //docCurr <-> bpBAccountCurrency
                    if (bankCode == "TBCBGE22")
                    {
                        if (bpBnkCode == "TBCBGE22")
                        {
                            transferType = "TransferWithinBankPaymentOrderIo"; //გადარიცხვა თიბისი ბანკის ფილიალებში
                        }
                        else if (string.IsNullOrEmpty(bpBnkCode) == false && bpBnkCode != "TBCBGE22") //pBnkCode != bankCode)
                        {
                            if (docCurr == "GEL")
                            {
                                transferType = "TransferToOtherBankNationalCurrencyPaymentOrderIo"; //გადარიცხვა სხვა ბანკში (ეროვნული ვალუტა)
                            }
                            else if (docCurr != "GEL")
                            {
                                transferType = "TransferToOtherBankForeignCurrencyPaymentOrderIo"; //გადარიცხვა სხვა ბანკში (უცხოური ვალუტა)
                            }
                        }
                    }
                    else if (bankCode == "BAGAGE22")
                    {
                        if (string.IsNullOrEmpty(bpBnkCode) == false)
                        {
                            if (docCurr == "GEL")
                            {
                                transferType = "TransferToNationalCurrencyPaymentOrderIo"; //გადარიცხვა (ეროვნული ვალუტა)
                            }
                            else if (docCurr != "GEL")
                            {
                                transferType = "TransferToForeignCurrencyPaymentOrderIo"; //გადარიცხვა (უცხოური ვალუტა)
                            }
                        }
                    }
                }
            }
            else
            {
                transferType = null;
            }

            return transferType;
        }

        private static void checkFillDoc(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            string trsfrAcct = oForm.DataSources.DBDataSources.Item("OVPM").GetValue("TrsfrAcct", 0).Trim(); // ბუღ. ანგარიში (გამგზავნი)
            string bankProgram = CommonFunctions.getBankProgram(trsfrAcct); //ბანკის პროგრამა (გამგზავნი)
            string opType = oForm.DataSources.DBDataSources.Item("OVPM").GetValue("U_opType", 0).Trim();
            string docType = oForm.DataSources.DBDataSources.Item("OVPM").GetValue("DocType", 0).Trim();
            string isPayToBank = oForm.DataSources.DBDataSources.Item("OVPM").GetValue("IsPaytoBnk", 0).Trim();

            if ((isPayToBank == "Y" || docType == "A") && opType != "other" && opType != "salaryPayment")
            {
                if (opType == "paymentToEmployee")
                {
                    List<int> oList = new List<int>();
                    SAPbouiCOM.DBDataSource oDBDataSource = oForm.DataSources.DBDataSources.Item("VPM4");

                    for (int i = 0; i < oDBDataSource.Size; i++)
                    {
                        string employee = oDBDataSource.GetValue("U_employee", i).Trim();
                        if (string.IsNullOrEmpty(employee) == false)
                        {
                            oList.Add(Convert.ToInt32(employee));
                        }
                    }

                    List<int> result = new List<int>();

                    if (oList.Count > 1)
                    {
                        result = oList.Select(o => o).Distinct().ToList();
                        if (result.Count > 1)
                        {
                            errorText = BDOSResources.getTranslate("EmployeesMustNotBeDifferent") + " : " + string.Join(",", result);
                            return;
                        }
                    }
                }
                if (bankProgram == "BOG")
                {
                    checkFillDocForBOG(oForm, out errorText);
                }
                else if (bankProgram == "TBC")
                {
                    checkFillDocForTBC(oForm, out errorText);
                }
                if ((bankProgram == "BOG" || bankProgram == "TBC") && string.IsNullOrEmpty(errorText))
                {
                    string status = oForm.DataSources.DBDataSources.Item("OVPM").GetValue("U_status", 0).Trim();
                    if (string.IsNullOrEmpty(status) || status == "empty" || status == "notToUpload")
                    {
                        SAPbouiCOM.ComboBox oComboBox = (SAPbouiCOM.ComboBox)oForm.Items.Item("statusCB").Specific;
                        oComboBox.Select("readyToLoad", SAPbouiCOM.BoSearchKey.psk_ByValue);
                    }
                }
            }
            else
            {
                SAPbouiCOM.ComboBox oComboBox = (SAPbouiCOM.ComboBox)oForm.Items.Item("statusCB").Specific;
                oComboBox.Select("notToUpload", SAPbouiCOM.BoSearchKey.psk_ByValue);
            }
        }

        private static void checkFillDocForTBC(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            Dictionary<string, string> dataForTransferType = getDataForTransferType(oForm);
            string transferType = getTransferType(dataForTransferType, out errorText);

            if (dataForTransferType == null)
            {
                errorText = BDOSResources.getTranslate("GeneralError");
                return;
            }
            //დოკუმენტის მონაცემები --->
            string docType = dataForTransferType["docType"];
            string opType = dataForTransferType["opType"];
            string diffCurr = dataForTransferType["diffCurr"];
            string docCurr = dataForTransferType["docCurr"];
            string description = dataForTransferType["description"];
            string chargeDetails = dataForTransferType["chargeDetails"];
            string beneficiaryAddress = dataForTransferType["BeneficiaryAddress"];
            string dispatchType = dataForTransferType["dispatchType"];
            string isPayToBank = dataForTransferType["isPayToBank"];
            //დოკუმენტის მონაცემები <---

            //კომპანიის მონაცემები (გამგზავნი) --->
            string trsfrAcct = dataForTransferType["trsfrAcct"]; // ბუღ. ანგარიში (გამგზავნი)
            string bankCode = dataForTransferType["bankCode"]; //ბანკის კოდი (გამგზავნი)
            string bankProgram = dataForTransferType["bankProgram"]; //ბანკის პროგრამა (გამგზავნი)
            //კომპანიის მონაცემები (გამგზავნი) <---

            //კომპანიის/თანამშრომლის მონაცემები (მიმღები) --->
            string creditAcct = dataForTransferType["creditAcct"]; //ანგარიში (მიმღები)
            string creditBankCode = dataForTransferType["creditBankCode"]; //ბანკის კოდი (მიმღები)
            string creditAcctCurrency = dataForTransferType["creditAcctCurrency"]; //ანგარიშის ვალუტა (მიმღები)
            string exchangeCurrency = dataForTransferType["exchangeCurrency"]; //კონვერტაციის ვალუტა
            //კომპანიის/თანამშრომლის მონაცემები (მიმღები) <---

            //კონტრაგენტის მონაცემები (მიმღები) --->
            string bpBnkCode = dataForTransferType["bpBnkCode"]; //ბანკის კოდი
            string bpBAccount = dataForTransferType["bpBAccount"]; //ანგარიში
            string bpBAccountCurrency = dataForTransferType["bpBAccountCurrency"]; //ვალუტა
            string treasuryCode = dataForTransferType["treasuryCode"];
            //კონტრაგენტის მონაცემები (მიმღები) <---

            if ((isPayToBank == "Y" || docType == "A") && opType != "other")
            {
                if (string.IsNullOrEmpty(transferType))
                {
                    errorText = BDOSResources.getTranslate("MakeSureToFillInTheDocuments") + "! "; //გადარიცხვის ტიპის დადგენა ვერ მოხერხდა, გადაამოწმეთ დოკუმენტის შევსების სისწორე
                    return;
                }

                string creditAcctTmp = null; //შევინახავთ მიმღების ვალუტას (იცვლება ტიპების მიხედვით)

                if (bankProgram == "TBC")
                {
                    if (docType == "A" && string.IsNullOrEmpty(opType))
                    {
                        errorText = BDOSResources.getTranslate("TheFollowingFieldIsMandatory") + " : \"" + oForm.Items.Item("opTypeS").Specific.caption + "\""; //აუცილებელია შემდეგი ველების შევსება
                        return;
                    }
                    else if (transferType == "TransferWithinBankPaymentOrderIo") //გადარიცხვა თიბისი ბანკის ფილიალებში
                    {
                        creditAcctTmp = bpBAccountCurrency;
                        if (string.IsNullOrEmpty(description))
                        {
                            errorText = BDOSResources.getTranslate("TheFollowingFieldIsMandatory") + " : \"" + oForm.Items.Item("descrptS").Specific.caption + "\""; //აუცილებელია შემდეგი ველების შევსება
                            return;
                        }
                    }
                    else if (transferType == "TreasuryTransferPaymentOrderIo" || transferType == "TreasuryTransferPaymentOrderIoBP") //საბიუჯეტო გადარიცხვა
                    {
                        creditAcctTmp = docCurr;
                        if (string.IsNullOrEmpty(treasuryCode))
                        {
                            errorText = BDOSResources.getTranslate("TheFollowingFieldIsMandatory") + " : \"" + BDOSResources.getTranslate("TreasuryCode") + "\""; //აუცილებელია შემდეგი ველების შევსება
                            return;
                        }
                    }
                    else if (transferType == "TransferToOwnAccountPaymentOrderIo") //გადარიცხვა საკუთარ ანგარიშზე
                    {
                        creditAcctTmp = creditAcctCurrency;
                        if (string.IsNullOrEmpty(creditAcct) || string.IsNullOrEmpty(description))
                        {
                            errorText = BDOSResources.getTranslate("TheFollowingFieldsAreMandatory") + " : \"" + oForm.Items.Item("creditActS").Specific.caption + "\", \"" + oForm.Items.Item("descrptS").Specific.caption + "\""; //აუცილებელია შემდეგი ველების შევსება
                            return;
                        }
                    }
                    else if (transferType == "CurrencyExchangePaymentOrderIo") //კონვერტაცია
                    {
                        creditAcctTmp = creditAcctCurrency;
                        if (string.IsNullOrEmpty(creditAcct) || string.IsNullOrEmpty(exchangeCurrency) || string.IsNullOrEmpty(description))
                        {
                            errorText = BDOSResources.getTranslate("TheFollowingFieldsAreMandatory") + " : \"" + oForm.Items.Item("creditActS").Specific.caption + "\", \"" + BDOSResources.getTranslate("CurrencyForExchange") + "\", \"" + oForm.Items.Item("descrptS").Specific.caption + "\""; //აუცილებელია შემდეგი ველების შევსება
                            return;
                        }
                        if (creditAcctCurrency != exchangeCurrency)
                        {
                            errorText = BDOSResources.getTranslate("CurrencyForExchangeAndTheCreditAccountSCurrencyIsDifferent") + "!"; //კონვერტაციის ვალუტა და მიმღები ანგარიშის ვალუტა განსხვავდება!";
                            return;
                        }
                    }
                    else if (docType != "A" && transferType == "TransferToOtherBankNationalCurrencyPaymentOrderIo") //გადარიცხვა სხვა ბანკში (ეროვნული ვალუტა)
                    {
                        creditAcctTmp = bpBAccountCurrency;
                        if (string.IsNullOrEmpty(description))
                        {
                            errorText = BDOSResources.getTranslate("TheFollowingFieldIsMandatory") + " : \"" + oForm.Items.Item("descrptS").Specific.caption + "\""; //აუცილებელია შემდეგი ველების შევსება
                            return;
                        }
                    }
                    else if (docType != "A" && transferType == "TransferToOtherBankForeignCurrencyPaymentOrderIo") //გადარიცხვა სხვა ბანკში (უცხოური ვალუტა)
                    {
                        creditAcctTmp = bpBAccountCurrency;
                        if (string.IsNullOrEmpty(chargeDetails) || string.IsNullOrEmpty(description) || string.IsNullOrEmpty(beneficiaryAddress))
                        {
                            errorText = BDOSResources.getTranslate("TheFollowingFieldsAreMandatory")
                                + " : \"" + oForm.Items.Item("chrgDtlsS").Specific.caption + "\", \""
                                + oForm.Items.Item("descrptS").Specific.caption
                                + "\", \"" + BDOSResources.getTranslate("BeneficiaryAddress") + "\""; //აუცილებელია შემდეგი ველების შევსება
                            return;
                        }
                    }
                    else if (docType == "A" && transferType == "TransferToOtherBankNationalCurrencyPaymentOrderIo") //გადარიცხვა სხვა ბანკში (ეროვნული ვალუტა)
                    {
                        creditAcctTmp = creditAcctCurrency;
                        if (string.IsNullOrEmpty(description))
                        {
                            errorText = BDOSResources.getTranslate("TheFollowingFieldIsMandatory") + " : \"" + oForm.Items.Item("descrptS").Specific.caption + "\""; //აუცილებელია შემდეგი ველების შევსება
                            return;
                        }
                    }
                    else if (docType == "A" && transferType == "TransferToOtherBankForeignCurrencyPaymentOrderIo") //გადარიცხვა სხვა ბანკში (უცხოური ვალუტა)
                    {
                        creditAcctTmp = creditAcctCurrency;
                        if (string.IsNullOrEmpty(chargeDetails) || string.IsNullOrEmpty(description) || string.IsNullOrEmpty(beneficiaryAddress))
                        {
                            errorText = BDOSResources.getTranslate("TheFollowingFieldsAreMandatory")
                                + " : \"" + oForm.Items.Item("chrgDtlsS").Specific.caption + "\", \""
                                + oForm.Items.Item("descrptS").Specific.caption
                                + "\", \"" + BDOSResources.getTranslate("BeneficiaryAddress") + "\""; //აუცილებელია შემდეგი ველების შევსება
                            return;
                        }
                    }

                    //დროებით
                    //if (string.IsNullOrEmpty(creditAcctTmp) == false && transferType != "CurrencyExchangePaymentOrderIo") //კონვერტაცია
                    //{
                    //    if (docCurr != creditAcctTmp)
                    //    {
                    //        errorText = BDOSResources.getTranslate("DocumentSCurrencyAndTheCreditAccountSCurrencyIsDifferent") + "!"; //დოკუმენტის ვალუტა და მიმღები ანგარიშის ვალუტა განსხვავდება!";
                    //        return;
                    //    }
                    //}
                }
            }
        }

        private static void checkFillDocForBOG(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            Dictionary<string, string> dataForTransferType = getDataForTransferType(oForm);
            string transferType = getTransferType(dataForTransferType, out errorText);

            if (dataForTransferType == null)
            {
                errorText = BDOSResources.getTranslate("GeneralError");
                return;
            }
            //დოკუმენტის მონაცემები --->
            string docType = dataForTransferType["docType"];
            string opType = dataForTransferType["opType"];
            string diffCurr = dataForTransferType["diffCurr"];
            string docCurr = dataForTransferType["docCurr"];
            string description = dataForTransferType["description"];
            string chargeDetails = dataForTransferType["chargeDetails"];
            string reportCode = dataForTransferType["reportCode"];
            string recipientCity = dataForTransferType["RecipientCity"];
            string beneficiaryAddress = dataForTransferType["BeneficiaryAddress"];
            string beneficiaryRegistrationCountryCode = dataForTransferType["BeneficiaryRegistrationCountryCode"];
            string dispatchType = dataForTransferType["dispatchType"];
            string isPayToBank = dataForTransferType["isPayToBank"];
            string docRate = dataForTransferType["docRate"];
            //დოკუმენტის მონაცემები <---

            //კომპანიის მონაცემები (გამგზავნი) --->
            string trsfrAcct = dataForTransferType["trsfrAcct"]; // ბუღ. ანგარიში (გამგზავნი)
            string bankCode = dataForTransferType["bankCode"]; //ბანკის კოდი (გამგზავნი)
            string bankProgram = dataForTransferType["bankProgram"]; //ბანკის პროგრამა (გამგზავნი)
            //კომპანიის მონაცემები (გამგზავნი) <---

            //კომპანიის/თანამშრომლის მონაცემები (მიმღები) --->
            string creditAcct = dataForTransferType["creditAcct"]; //ანგარიში (მიმღები)
            string creditBankCode = dataForTransferType["creditBankCode"]; //ბანკის კოდი (მიმღები)
            string creditAcctCurrency = dataForTransferType["creditAcctCurrency"]; //ანგარიშის ვალუტა (მიმღები)
            string exchangeCurrency = dataForTransferType["exchangeCurrency"]; //კონვერტაციის ვალუტა
            //კომპანიის/თანამშრომლის მონაცემები (მიმღები) <---

            //კონტრაგენტის მონაცემები (მიმღები) --->
            string bpBnkCode = dataForTransferType["bpBnkCode"]; //ბანკის კოდი
            string bpBAccount = dataForTransferType["bpBAccount"]; //ანგარიში
            string bpBAccountCurrency = dataForTransferType["bpBAccountCurrency"]; //ვალუტა
            string treasuryCode = dataForTransferType["treasuryCode"];
            //კონტრაგენტის მონაცემები (მიმღები) <---

            if ((isPayToBank == "Y" || docType == "A") && opType != "other")
            {
                if (string.IsNullOrEmpty(transferType))
                {
                    errorText = BDOSResources.getTranslate("MakeSureToFillInTheDocuments") + "! "; //გადარიცხვის ტიპის დადგენა ვერ მოხერხდა, გადაამოწმეთ დოკუმენტის შევსების სისწორე
                    return;
                }

                string creditAcctTmp = null; //შევინახავთ მიმღების ვალუტას (იცვლება ტიპების მიხედვით)

                if (bankProgram == "BOG")
                {
                    if (docType == "A" && string.IsNullOrEmpty(opType))
                    {
                        errorText = BDOSResources.getTranslate("TheFollowingFieldIsMandatory") + " : \"" + oForm.Items.Item("opTypeS").Specific.caption + "\""; //აუცილებელია შემდეგი ველების შევსება
                        return;
                    }
                    else if (transferType == "TreasuryTransferPaymentOrderIo" || transferType == "TreasuryTransferPaymentOrderIoBP") //საბიუჯეტო გადარიცხვა
                    {
                        creditAcctTmp = docCurr;
                        if (string.IsNullOrEmpty(treasuryCode) || string.IsNullOrEmpty(description))
                        {
                            errorText = BDOSResources.getTranslate("TheFollowingFieldsAreMandatory") + " : \"" + oForm.Items.Item("descrptS").Specific.caption + "\", \"" + BDOSResources.getTranslate("TreasuryCode") + "\""; //აუცილებელია შემდეგი ველების შევსება
                            return;
                        }
                    }
                    else if (transferType == "TransferToOwnAccountPaymentOrderIo") //გადარიცხვა საკუთარ ანგარიშზე
                    {
                        creditAcctTmp = creditAcctCurrency;
                        if (string.IsNullOrEmpty(creditAcct))
                        {
                            errorText = BDOSResources.getTranslate("TheFollowingFieldIsMandatory") + " : \"" + oForm.Items.Item("creditActS").Specific.caption + "\""; //აუცილებელია შემდეგი ველების შევსება
                            return;
                        }
                    }
                    else if (transferType == "CurrencyExchangePaymentOrderIo") //კონვერტაცია
                    {
                        creditAcctTmp = creditAcctCurrency;
                        if (string.IsNullOrEmpty(creditAcct) || string.IsNullOrEmpty(exchangeCurrency) || docRate == "0")
                        {
                            errorText = BDOSResources.getTranslate("TheFollowingFieldsAreMandatory") + " : \"" + oForm.Items.Item("creditActS").Specific.caption + "\", \"" + BDOSResources.getTranslate("CurrencyForExchange") + "\", \"" + BDOSResources.getTranslate("DocRate") + "\""; //აუცილებელია შემდეგი ველების შევსება
                            return;
                        }
                        if (creditAcctCurrency != exchangeCurrency)
                        {
                            errorText = BDOSResources.getTranslate("CurrencyForExchangeAndTheCreditAccountSCurrencyIsDifferent") + "!"; //კონვერტაციის ვალუტა და მიმღები ანგარიშის ვალუტა განსხვავდება!";
                            return;
                        }
                    }
                    else if (docType != "A" && transferType == "TransferToNationalCurrencyPaymentOrderIo") //გადარიცხვა (ეროვნული ვალუტა)
                    {
                        creditAcctTmp = bpBAccountCurrency;
                        if (string.IsNullOrEmpty(description) || string.IsNullOrEmpty(dispatchType))
                        {
                            errorText = BDOSResources.getTranslate("TheFollowingFieldsAreMandatory") + " : \"" + oForm.Items.Item("descrptS").Specific.caption + "\", \"" + BDOSResources.getTranslate("DispatchType") + "\""; //აუცილებელია შემდეგი ველების შევსება
                            return;
                        }
                    }
                    else if (docType != "A" && transferType == "TransferToForeignCurrencyPaymentOrderIo") //გადარიცხვა (უცხოური ვალუტა)
                    {
                        creditAcctTmp = bpBAccountCurrency;
                        if (string.IsNullOrEmpty(chargeDetails) || string.IsNullOrEmpty(description)
                            || string.IsNullOrEmpty(reportCode) || string.IsNullOrEmpty(beneficiaryAddress)
                            || string.IsNullOrEmpty(recipientCity) || string.IsNullOrEmpty(beneficiaryRegistrationCountryCode))
                        {
                            errorText = BDOSResources.getTranslate("TheFollowingFieldsAreMandatory")
                                + " : \"" + oForm.Items.Item("chrgDtlsS").Specific.caption
                                + "\", \"" + oForm.Items.Item("rprtCodeS").Specific.caption
                                + "\", \"" + oForm.Items.Item("descrptS").Specific.caption
                                + "\", \"" + BDOSResources.getTranslate("BeneficiaryAddress")
                                + "\", \"" + BDOSResources.getTranslate("RecipientCity")
                                + "\", \"" + BDOSResources.getTranslate("BeneficiaryRegistrationCountryCode") + "\"";
                            //აუცილებელია შემდეგი ველების შევსება
                            return;
                        }
                    }
                    else if (docType == "A" && transferType == "TransferToNationalCurrencyPaymentOrderIo") //გადარიცხვა (ეროვნული ვალუტა)
                    {
                        creditAcctTmp = creditAcctCurrency;
                        if (string.IsNullOrEmpty(description) || string.IsNullOrEmpty(dispatchType))
                        {
                            errorText = BDOSResources.getTranslate("TheFollowingFieldsAreMandatory") + " : \"" + oForm.Items.Item("descrptS").Specific.caption + "\", \"" + BDOSResources.getTranslate("DispatchType") + "\""; //აუცილებელია შემდეგი ველების შევსება
                            return;
                        }
                    }
                    else if (docType == "A" && transferType == "TransferToForeignCurrencyPaymentOrderIo") //გადარიცხვა (უცხოური ვალუტა)
                    {
                        creditAcctTmp = creditAcctCurrency;
                        if (string.IsNullOrEmpty(chargeDetails) || string.IsNullOrEmpty(description)
                             || string.IsNullOrEmpty(reportCode) || string.IsNullOrEmpty(beneficiaryAddress)
                             || string.IsNullOrEmpty(recipientCity) || string.IsNullOrEmpty(beneficiaryRegistrationCountryCode))
                        {
                            errorText = BDOSResources.getTranslate("TheFollowingFieldsAreMandatory")
                                + " : \"" + oForm.Items.Item("chrgDtlsS").Specific.caption
                                + "\", \"" + oForm.Items.Item("rprtCodeS").Specific.caption
                                + "\", \"" + oForm.Items.Item("descrptS").Specific.caption
                                + "\", \"" + BDOSResources.getTranslate("BeneficiaryAddress")
                                + "\", \"" + BDOSResources.getTranslate("RecipientCity")
                                + "\", \"" + BDOSResources.getTranslate("BeneficiaryRegistrationCountryCode") + "\"";
                            //აუცილებელია შემდეგი ველების შევსება
                            return;
                        }
                    }

                    //დროებით
                    //if (string.IsNullOrEmpty(creditAcctTmp) == false && transferType != "CurrencyExchangePaymentOrderIo") //კონვერტაცია
                    //{
                    //    if (docCurr != creditAcctTmp)
                    //    {
                    //        errorText = BDOSResources.getTranslate("DocumentSCurrencyAndTheCreditAccountSCurrencyIsDifferent") + "!"; //დოკუმენტის ვალუტა და მიმღები ანგარიშის ვალუტა განსხვავდება!";
                    //        return;
                    //    }
                    //}
                }
            }
        }

        public static void uiApp_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText;

            SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(BusinessObjectInfo.FormTypeEx, Program.currentFormCount);

            //შემოწმება
            if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD || BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE)
            {
                if (!BusinessObjectInfo.BeforeAction && !BusinessObjectInfo.ActionSuccess)
                {
                    BubbleEvent = false;
                }

                if (BusinessObjectInfo.BeforeAction)
                {
                    SAPbouiCOM.DBDataSource oDBDataSource = oForm.DataSources.DBDataSources.Item("OVPM");
                    if (oDBDataSource.GetValue("Canceled", 0) == "N" && !Program.cancellationTrans)
                    {
                        CalcPhysicalEntityTaxes(oForm);

                        // მოგების გადასახადი
                        if (ProfitTaxTypeIsSharing)
                        {
                            if (oDBDataSource.GetValue("DocType", 0) == "S")
                            {
                                bool TaxAccountsIsEmpty = ProfitTax.TaxAccountsIsEmpty();
                                if (TaxAccountsIsEmpty)
                                {
                                    Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("TaxAccounts") + " " + BDOSResources.getTranslate("YouCantLeaveEmpty"));
                                    Program.uiApp.MessageBox(BDOSResources.getTranslate("OperationUnsuccesfullSeeLog"));
                                    BubbleEvent = false;
                                }
                            }

                            if (oDBDataSource.GetValue("U_liablePrTx", 0) == "Y")
                            {
                                if (string.IsNullOrEmpty(oDBDataSource.GetValue("U_prBase", 0)))
                                {
                                    Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("TaxableObject") + " " + BDOSResources.getTranslate("YouCantLeaveEmpty"));
                                    Program.uiApp.MessageBox(BDOSResources.getTranslate("OperationUnsuccesfullSeeLog"));
                                    BubbleEvent = false;
                                }
                            }
                        }

                        CheckAccounts(oForm, out errorText);
                        if (errorText != null)
                        {
                            Program.uiApp.MessageBox(errorText);
                            BubbleEvent = false;
                        }

                        checkFillDoc(oForm, out errorText);
                        if (errorText != null)
                        {
                            SAPbouiCOM.ComboBox oComboBox = (SAPbouiCOM.ComboBox)oForm.Items.Item("statusCB").Specific;
                            oComboBox.Select("notToUpload", SAPbouiCOM.BoSearchKey.psk_ByValue);
                            Program.uiApp.SetStatusBarMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, false);
                            BubbleEvent = false;
                        }
                    }
                }

                if (BusinessObjectInfo.ActionSuccess && !BusinessObjectInfo.BeforeAction && BusinessObjectInfo.EventType != SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE) //BusinessObjectInfo.ActionSuccess != BusinessObjectInfo.BeforeAction
                {
                    //დოკუმენტის გატარების დროს გატარდეს ბუღლტრული გატარება
                    SAPbouiCOM.DBDataSource oDBDataSource = oForm.DataSources.DBDataSources.Item(0);
                    if (oDBDataSource.GetValue("Canceled", 0) == "N" && !Program.cancellationTrans)
                    {
                        string opType = oDBDataSource.GetValue("U_opType", 0).Trim();
                        if (opType != "salaryPayment" & opType != "paymentToEmployee")
                        {
                            string DocEntry = oDBDataSource.GetValue("DocEntry", 0);
                            string DocCurrency = oDBDataSource.GetValue("DocCurr", 0);
                            //decimal DocRate = Convert.ToDecimal( DocDBSourcePAYR.GetValue("DocRate", 0));
                            decimal DocRate = Convert.ToDecimal(FormsB1.cleanStringOfNonDigits(oDBDataSource.GetValue("DocRate", 0)));
                            string DocNum = oDBDataSource.GetValue("DocNum", 0);
                            DateTime DocDate = DateTime.ParseExact(oDBDataSource.GetValue("DocDate", 0), "yyyyMMdd", CultureInfo.InvariantCulture);

                            CommonFunctions.StartTransaction();

                            Program.JrnLinesGlobal = new DataTable();
                            DataTable reLines;
                            DataTable JrnLinesDT = createAdditionalEntries(oForm, null, null, null, DocCurrency, out reLines, DocRate);

                            JrnEntry(DocEntry, DocNum, DocDate, JrnLinesDT, reLines, out errorText);
                            if (errorText != null)
                            {
                                Program.uiApp.MessageBox(errorText);
                                BubbleEvent = false;
                            }
                            else
                            {
                                if (!BusinessObjectInfo.ActionSuccess)
                                {
                                    Program.JrnLinesGlobal = JrnLinesDT;
                                }
                            }
                            if (Program.oCompany.InTransaction)
                            {
                                //თუ დოკუმენტი გატარდა, მერე ვაკეთებს ბუღალტრულ გატარებას
                                if (BusinessObjectInfo.ActionSuccess && !BusinessObjectInfo.BeforeAction)
                                {
                                    CommonFunctions.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                                }
                                else
                                {
                                    CommonFunctions.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                                }
                            }
                            else
                            {
                                Program.uiApp.MessageBox("ტრანზაქციის დასრულების შეცდომა");
                                BubbleEvent = false;
                            }
                        }
                    }
                }
            }

            if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE && !BusinessObjectInfo.BeforeAction && BusinessObjectInfo.ActionSuccess)
            {
                if (Program.cancellationTrans && Program.canceledDocEntry != 0)
                {
                    cancellation(oForm, Program.canceledDocEntry);
                    Program.canceledDocEntry = 0;
                }
            }

            //if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE)
            //{
            //    if (BusinessObjectInfo.BeforeAction)
            //    {
            //        if (Program.cancellationTrans && Program.canceledDocEntry != 0)
            //        {

            //        }
            //        else
            //        {
            //            checkFillDoc(oForm, out errorText);
            //            if (errorText != null)
            //            {
            //                Program.uiApp.SetStatusBarMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, false);
            //                BubbleEvent = false;
            //            }
            //        }
            //    }
            //    else if (BusinessObjectInfo.BeforeAction == false && BusinessObjectInfo.ActionSuccess)
            //    {
            //        if (Program.cancellationTrans && Program.canceledDocEntry != 0)
            //        {
            //            cancellation(oForm, Program.canceledDocEntry, out errorText);
            //            Program.canceledDocEntry = 0;
            //        }
            //    }
            //}

            if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD && !BusinessObjectInfo.BeforeAction)
            {
                //when "Keep Visible" is not selected Program.uiApp.Forms.ActiveForm = List of Outgoing Payments form (Type = 10045), so we need check
                if (Program.uiApp.Forms.ActiveForm.Type == 426) //Keep Visible Case!!!
                    oForm = Program.uiApp.Forms.ActiveForm;
                formDataLoad(oForm);
            }
        }

        public static DataTable createAdditionalEntries(SAPbouiCOM.Form oForm, SAPbobsCOM.GeneralData oGeneralData, DataTable DTSource, DataTable DTSourceVPM2, string docCurrency, out DataTable reLines, decimal docRate)
        {
            DataTable jeLines = JournalEntry.JournalEntryTable();

            reLines = ProfitTax.ProfitTaxTable();
            DataRow reLinesRow = null;
            DataTable AccountTable = CommonFunctions.GetOACTTable();

            SAPbouiCOM.DBDataSource DBDataSourceTable = null;
            int JEcount = 0;

            SAPbouiCOM.DBDataSource docDBSource = null;
            SAPbouiCOM.DBDataSource BPDataSourceTable = null;
            SAPbouiCOM.DBDataSources docDBSources = null;
            if (oForm == null)
            {
                JEcount = DTSourceVPM2.Rows.Count;
                ProfitTaxTypeIsSharing = ProfitTax.ProfitTaxTypeIsSharing();
            }
            else
            {
                docDBSources = oForm.DataSources.DBDataSources;
                DBDataSourceTable = docDBSources.Item("VPM2");
                JEcount = DBDataSourceTable.Size;
                docDBSource = docDBSources.Item("OVPM");

                BPDataSourceTable = docDBSources.Item("OCRD");
            }

            DateTime docDate = DateTime.ParseExact(CommonFunctions.getChildOrDbDataSourceValue(docDBSource, null, DTSource, "DocDate", 0).ToString(), "yyyyMMdd", CultureInfo.InvariantCulture);
            string CardCode = CommonFunctions.getChildOrDbDataSourceValue(BPDataSourceTable, null, DTSource, "CardCode", 0).ToString();

            SAPbobsCOM.BusinessPartners oBP = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);
            string vatCode = "";
            string TaxType = "";
            if (oBP.GetByKey(CardCode))
            {
                vatCode = oBP.VatGroup;
                TaxType = oBP.UserFields.Fields.Item("U_BDO_TaxTyp").Value;
            }

            docCurrency = docCurrency == CommonFunctions.getLocalCurrency() ? "" : docCurrency;

            //დღგ-ის გატარება
            if (TaxType == "12") // = 18
            {
                for (int i = 0; i < JEcount; i++)
                {
                    string InvType = CommonFunctions.getChildOrDbDataSourceValue(DBDataSourceTable, null, DTSourceVPM2, "InvType", i).ToString();

                    if (InvType != "18")
                    {
                        continue;
                    }

                    int InvoiceEntry = Convert.ToInt32(CommonFunctions.getChildOrDbDataSourceValue(DBDataSourceTable, null, DTSourceVPM2, "DocEntry", i));
                    SAPbobsCOM.Documents oInvoice = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices);
                    oInvoice.GetByKey(InvoiceEntry);

                    decimal sumApplied = Convert.ToDecimal(CommonFunctions.getChildOrDbDataSourceValue(DBDataSourceTable, null, DTSourceVPM2, "AppliedFC", i));

                    if (sumApplied == 0)
                    {
                        sumApplied = Convert.ToDecimal(CommonFunctions.getChildOrDbDataSourceValue(DBDataSourceTable, null, DTSourceVPM2, "SumApplied", i));
                    }

                    sumApplied = sumApplied * Convert.ToDecimal(oInvoice.DocRate);

                    SAPbobsCOM.VatGroups oVatCode = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVatGroups);
                    oVatCode.GetByKey(vatCode);

                    string DebitAccount = oVatCode.UserFields.Fields.Item("U_BDOSAccF").Value;
                    string CreditAccount = oVatCode.TaxAccount;

                    decimal vatRate = LandedCosts.GetVatGroupRate(vatCode);
                    decimal TaxAmount = sumApplied * vatRate / (100 + vatRate);
                    decimal TaxAmountFC = docCurrency == "" ? 0 : TaxAmount / docRate;

                    if (TaxAmount > 0)
                    {
                        JournalEntry.AddJournalEntryRow(AccountTable, jeLines, "Full", DebitAccount, CreditAccount, TaxAmount, TaxAmountFC, docCurrency, "", "", "", "", "", "", "", "");
                    }
                }
            }

            //მოგების გადასახადის გატარება
            if (ProfitTaxTypeIsSharing)
            {
                string U_liablePrTx = CommonFunctions.getChildOrDbDataSourceValue(docDBSource, null, DTSource, "U_liablePrTx", 0).ToString(); //docDBSource.GetValue("U_liablePrTx", 0).Trim();
                decimal NoDocSum = FormsB1.cleanStringOfNonDigits(CommonFunctions.getChildOrDbDataSourceValue(docDBSource, null, DTSource, "NoDocSum", 0).ToString());

                string DebitAccount = CommonFunctions.getOADM("U_BDO_CapAcc").ToString();
                string CreditAccount = CommonFunctions.getOADM("U_BDO_TaxAcc").ToString();
                decimal U_BDO_PrTxRt = Convert.ToDecimal(CommonFunctions.getOADM("U_BDO_PrTxRt").ToString(), CultureInfo.InvariantCulture);

                if (U_liablePrTx == "Y" & NoDocSum > 0)
                {
                    string prBase = CommonFunctions.getChildOrDbDataSourceValue(docDBSource, null, DTSource, "U_prBase", 0).ToString().Trim();
                    decimal TaxAmount = NoDocSum * U_BDO_PrTxRt / (100 - U_BDO_PrTxRt);
                    decimal TaxAmountFC = docCurrency == "" ? 0 : TaxAmount / docRate;

                    JournalEntry.AddJournalEntryRow(AccountTable, jeLines, "Full", DebitAccount, CreditAccount, TaxAmount, TaxAmountFC, docCurrency,
                                                            "", "", "", "", "", "", "", "");

                    reLinesRow = reLines.Rows.Add();

                    reLinesRow["debitAccount"] = DebitAccount;
                    reLinesRow["creditAccount"] = CreditAccount;
                    reLinesRow["prBase"] = prBase;
                    reLinesRow["txType"] = "Accrual";
                    reLinesRow["amtTx"] = NoDocSum;
                    reLinesRow["amtPrTx"] = TaxAmount;
                }

                for (int i = 0; i < JEcount; i++)
                {
                    string InvType = CommonFunctions.getChildOrDbDataSourceValue(DBDataSourceTable, null, DTSourceVPM2, "InvType", i).ToString();

                    if (InvType == "18" || InvType == "204")
                    {
                        decimal SumApplied = FormsB1.cleanStringOfNonDigits(CommonFunctions.getChildOrDbDataSourceValue(DBDataSourceTable, null, DTSourceVPM2, "SumApplied", i).ToString());
                        decimal TaxAmount = SumApplied * U_BDO_PrTxRt / (100 - U_BDO_PrTxRt);
                        decimal TaxAmountFC = docCurrency == "" ? 0 : TaxAmount / docRate;

                        int InvoiceEntry = Convert.ToInt32(CommonFunctions.getChildOrDbDataSourceValue(DBDataSourceTable, null, DTSourceVPM2, "DocEntry", i));
                        SAPbobsCOM.Documents oInvoice;

                        if (InvType == "18")
                        {
                            oInvoice = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices);
                            oInvoice.GetByKey(InvoiceEntry);
                            U_liablePrTx = oInvoice.UserFields.Fields.Item("U_nonEconExp").Value;
                        }
                        else
                        {
                            oInvoice = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDownPayments);
                            oInvoice.GetByKey(InvoiceEntry);
                            U_liablePrTx = oInvoice.UserFields.Fields.Item("U_liablePrTx").Value;
                        }

                        if (U_liablePrTx == "Y")
                        {
                            string prBase = oInvoice.UserFields.Fields.Item("U_prBase").Value;

                            JournalEntry.AddJournalEntryRow(AccountTable, jeLines, "Full", DebitAccount, CreditAccount, TaxAmount, TaxAmountFC, docCurrency,
                                                        "", "", "", "", "", "", "", "");

                            reLinesRow = reLines.Rows.Add();

                            reLinesRow["debitAccount"] = DebitAccount;
                            reLinesRow["creditAccount"] = CreditAccount;
                            reLinesRow["prBase"] = prBase;
                            reLinesRow["txType"] = "Accrual";
                            reLinesRow["amtTx"] = SumApplied;
                            reLinesRow["amtPrTx"] = TaxAmount;
                        }
                    }
                }
            }

            string wTCode = CommonFunctions.getChildOrDbDataSourceValue(BPDataSourceTable, null, DTSource, "WtCode", 0).ToString();
            bool isWTLiable = CommonFunctions.getChildOrDbDataSourceValue(BPDataSourceTable, null, DTSource, "WTLiable", 0).ToString() == "Y";

            if (isWTLiable)
            {
                decimal whTaxAmt = FormsB1.cleanStringOfNonDigits(CommonFunctions.getChildOrDbDataSourceValue(docDBSource, null, DTSource, "U_BDOSWhtAmt", 0).ToString());
                decimal whTaxAmtFC = docCurrency == "" ? 0 : whTaxAmt / docRate;

                decimal pensEmployedAmt = FormsB1.cleanStringOfNonDigits(CommonFunctions.getChildOrDbDataSourceValue(docDBSource, null, DTSource, "U_BDOSPnPhAm", 0).ToString()); //დასაქმებული
                decimal pensEmployedAmtFC = docCurrency == "" ? 0 : pensEmployedAmt / docRate;

                decimal pensEmployerAmt = FormsB1.cleanStringOfNonDigits(CommonFunctions.getChildOrDbDataSourceValue(docDBSource, null, DTSource, "U_BDOSPnCoAm", 0).ToString()); //დამსაქმებელი
                decimal pensEmployerAmtFC = docCurrency == "" ? 0 : pensEmployerAmt / docRate;

                decimal pensEmployerDiffAmt = FormsB1.cleanStringOfNonDigits(CommonFunctions.getChildOrDbDataSourceValue(docDBSource, null, DTSource, "U_BDOSPnCoDiffAm", 0).ToString(), true); //დამსაქმებელი განსხვავების თანხა

                if (pensEmployedAmt > 0 && pensEmployerAmt > 0)
                {
                    //bool physicalEntityTax = CommonFunctions.getValue("OWHT", "U_BDOSPhisTx", "WTCode", wTCode).ToString() == "Y";
                    string pensionCoWTCode = CommonFunctions.getOADM("U_BDOSPnCoP").ToString();
                    string pensionPhWTCode = CommonFunctions.getOADM("U_BDOSPnPh").ToString();

                    string debitAccount;
                    string creditAccount;
                    string distrRule1 = "";
                    string distrRule2 = "";
                    string distrRule3 = "";
                    string distrRule4 = "";
                    string distrRule5 = "";

                    string project = CommonFunctions.getChildOrDbDataSourceValue(docDBSource, null, DTSource, "PrjCode", 0).ToString();

                    debitAccount = CommonFunctions.getValue("OWHT", "Account", "WTCode", wTCode).ToString(); //BP-ს ძირითადი WTCode-ს ანგარიში
                    JournalEntry.AddJournalEntryRow(AccountTable, jeLines, "OnlyDebit", debitAccount, "", whTaxAmt + pensEmployedAmt, whTaxAmtFC + pensEmployedAmtFC, docCurrency,
                                                        distrRule1, distrRule2, distrRule3, distrRule4, distrRule5, project, "", "");

                    creditAccount = CommonFunctions.getValue("OWHT", "U_BdgtDbtAcc", "WTCode", wTCode).ToString(); //BP-ს ძირითადი WTCode-ს ვალდებულების ანგარიში
                    JournalEntry.AddJournalEntryRow(AccountTable, jeLines, "OnlyCredit", "", creditAccount, whTaxAmt, whTaxAmtFC, docCurrency,
                                                        distrRule1, distrRule2, distrRule3, distrRule4, distrRule5, project, "", "");

                    creditAccount = CommonFunctions.getValue("OWHT", "Account", "WTCode", pensionPhWTCode).ToString(); //U_BdgtDbtAcc დასაქმებულის საპენსიოს ვალდებულების ანგარიში
                    JournalEntry.AddJournalEntryRow(AccountTable, jeLines, "OnlyCredit", "", creditAccount, pensEmployedAmt, pensEmployedAmtFC, docCurrency,
                                    distrRule1, distrRule2, distrRule3, distrRule4, distrRule5, project, "", "");

                    debitAccount = CommonFunctions.getValue("OWHT", "Account", "WTCode", pensionCoWTCode).ToString(); // დამსაქმებლის საპენსიოს ანგარიში
                    creditAccount = CommonFunctions.getValue("OWHT", "U_BdgtDbtAcc", "WTCode", pensionCoWTCode).ToString(); // დამსაქმებლის საპენსიოს ვალდებულების ანგარიში

                    JournalEntry.AddJournalEntryRow(AccountTable, jeLines, "Full", debitAccount, creditAccount, pensEmployerAmt, pensEmployerAmtFC, docCurrency,
                                                        distrRule1, distrRule2, distrRule3, distrRule4, distrRule5, project, "", "");

                    if (pensEmployerDiffAmt > 0)
                    {
                        debitAccount = CommonFunctions.getPeriodsCategory("GLLossXdif", docDate.Year.ToString()); //საკურსო სხვაობის ხარჯი - დებეტი (Loss)

                        JournalEntry.AddJournalEntryRow(AccountTable, jeLines, "OnlyCredit", "", creditAccount, pensEmployerDiffAmt, decimal.Zero, string.Empty,
                                    distrRule1, distrRule2, distrRule3, distrRule4, distrRule5, project, "", "");
                        JournalEntry.AddJournalEntryRow(AccountTable, jeLines, "OnlyDebit", debitAccount, "", pensEmployerDiffAmt, decimal.Zero, string.Empty,
                                    distrRule1, distrRule2, distrRule3, distrRule4, distrRule5, project, "", "");
                    }
                    else if (pensEmployerDiffAmt < 0)
                    {
                        debitAccount = creditAccount; //საპენსიო შუალედური - დებეტი
                        creditAccount = CommonFunctions.getPeriodsCategory("GLGainXdif", docDate.Year.ToString()); //საკურსო სხვაობის ხარჯი - კრედიტი (Gain)

                        JournalEntry.AddJournalEntryRow(AccountTable, jeLines, "OnlyCredit", "", creditAccount, Math.Abs(pensEmployerDiffAmt), decimal.Zero, string.Empty,
                                    distrRule1, distrRule2, distrRule3, distrRule4, distrRule5, project, "", "");
                        JournalEntry.AddJournalEntryRow(AccountTable, jeLines, "OnlyDebit", debitAccount, "", Math.Abs(pensEmployerDiffAmt), decimal.Zero, string.Empty,
                                    distrRule1, distrRule2, distrRule3, distrRule4, distrRule5, project, "", "");
                    }
                }
            }

            return jeLines;
        }

        public static void JrnEntry(string DocEntry, string DocNum, DateTime DocDate, DataTable JrnLinesDT, DataTable reLines, out string errorText)
        {
            try
            {
                JournalEntry.JrnEntry(DocEntry, "46", "Outgoing payment: " + DocNum, DocDate, JrnLinesDT, out errorText);

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

                ProfitTax.AddRecord(reLines, "46", "Outgoing payment: " + DocNum, out errorText);
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
        }

        public static void cancellation(SAPbouiCOM.Form oForm, int docEntry)
        {
            try
            {
                JournalEntry.cancellation(oForm, docEntry, "46", out var errorText);
                if (!string.IsNullOrEmpty(errorText))
                {
                    throw new Exception(errorText);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static void uiApp_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText;
            if (FormUID == "OutgoingPaymentNewDate")
            {
                //if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && !pVal.BeforeAction)
                //{
                //    SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                //    if (pVal.ItemUID == "1")
                //    {
                //        string newDate = oForm.Items.Item("newDate").Specific.Value;
                //        changeDocDateRate(oForm, newDate);
                //    }
                //}
            }
            else
            {
                if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
                {
                    SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
                    {
                        if (pVal.BeforeAction)
                        {
                            createFormItems(oForm, out errorText);
                            Program.FORM_LOAD_FOR_VISIBLE = true;
                            Program.FORM_LOAD_FOR_ACTIVATE = true;

                            formDataLoad(oForm);
                        }
                    }

                    else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_COMBO_SELECT)
                    {
                        if (pVal.BeforeAction)
                        { }
                        else
                        {
                            if (pVal.ItemUID == "110") //Item - WTax Code
                            {
                                CalcPhysicalEntityTaxes(oForm, true);
                            }

                            else if (pVal.ItemUID == "opTypeCB" || pVal.ItemUID == "18" || pVal.ItemUID == "107")
                            {
                                if (pVal.ItemUID == "107" && oForm.DataSources.DBDataSources.Item("OVPM").GetValue("IsPaytoBnk", 0).Trim() != "Y")
                                {
                                    return;
                                }
                                setVisibleFormItems(oForm);
                            }
                        }

                        oForm.Freeze(true);
                        comboSelect(oForm, pVal, out errorText);
                        oForm.Freeze(false);
                    }

                    else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE)
                    {
                        if (pVal.BeforeAction)
                        { }
                        else
                        {
                            if (Program.FORM_LOAD_FOR_ACTIVATE)
                            {
                                setVisibleFormItems(oForm);
                                Program.FORM_LOAD_FOR_ACTIVATE = false;
                            }

                            if (Program.openPaymentMeans)
                            {
                                oForm.Freeze(false);

                                Program.openPaymentMeans = false;
                                setVisibleFormItems(oForm);

                                if (Program.openPaymentMeansByPostDateChange)
                                {
                                    Program.openPaymentMeansByPostDateChange = false;
                                    try
                                    {
                                        string docEntry = oForm.DataSources.DBDataSources.Item("OVPM").GetValue("DocEntry", 0);
                                        oForm.Freeze(true);
                                        Program.uiApp.ActivateMenuItem("5907"); //Save as Draft

                                        Program.uiApp.OpenForm((SAPbouiCOM.BoFormObjectEnum)140, "", docEntry);
                                    }
                                    catch
                                    {
                                        Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("DocumentCantSaveAsDraftAutomatically") + " " + BDOSResources.getTranslate("TryItManually"), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                    }
                                    finally
                                    {
                                        oForm.Freeze(false);
                                    }
                                }
                            }
                        }
                    }

                    else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
                    {
                        if (pVal.BeforeAction)
                        {
                            if (pVal.ItemUID == "1")
                                CommonFunctions.fillDocRate(oForm, "OVPM");
                        }
                        else
                        {
                            if (pVal.ItemUID == "57" || pVal.ItemUID == "56" || pVal.ItemUID == "58")
                            {
                                setVisibleFormItems(oForm);

                                oForm.Items.Item("95").Enabled = true;
                                oForm.Items.Item("234000005").Enabled = true;

                                IncomingPayment.SetUsBlaAgRtSAvailability(oForm);
                            }

                            else if (pVal.ItemUID == "ChngDcDt")
                            {
                                CurrentForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);
                                int answer = Program.uiApp.MessageBox(BDOSResources.getTranslate("AfterChangeDateFormWillCloseReopen") + ". " + BDOSResources.getTranslate("WouldYouWantToContinueTheOperation") + "?", 1, BDOSResources.getTranslate("Yes"), BDOSResources.getTranslate("No"), ""); //თარიღის ცვლილების შემდეგ ფორმა დაიხურება და ხელახლა გაიხსნება

                                if (answer == 2)
                                {
                                    return;
                                }

                                SAPbouiCOM.Form noForm = null;
                                createFormNewDate(noForm, out errorText);
                            }

                            else if (pVal.ItemUID == "UsBlaAgRtS" && !pVal.InnerEvent)
                            {
                                CommonFunctions.fillDocRate(oForm, "OVPM");
                            }

                            else if ((pVal.ItemUID == "liablePrTx" || pVal.ItemUID == "37") && !pVal.InnerEvent)
                            {
                                oForm.Freeze(true);
                                taxes_OnClick(oForm);
                                oForm.Freeze(false);
                            }
                        }
                    }

                    else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_VALIDATE)
                    {
                        if (pVal.BeforeAction)
                        { }
                        else
                        {
                            if ((pVal.ItemUID == "5" || pVal.ItemUID == "234000005") && !pVal.InnerEvent)
                            {
                                setVisibleFormItems(oForm);

                                if (pVal.ItemUID == "5") //Business Partner
                                {
                                    var prjCode = oForm.DataSources.DBDataSources.Item("OVPM").GetValue("PrjCode", 0);
                                    var agrNo = oForm.DataSources.DBDataSources.Item("OVPM").GetValue("AgrNo", 0);
                                    oForm.Items.Item("95").Enabled = string.IsNullOrEmpty(prjCode);
                                    oForm.Items.Item("95").Specific.Value = "";
                                    oForm.Items.Item("234000005").Enabled = string.IsNullOrEmpty(agrNo);
                                    oForm.Items.Item("234000005").Specific.Value = "";

                                    IncomingPayment.SetUsBlaAgRtSAvailability(oForm);
                                }
                            }
                        }
                    }

                    else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                    {
                        if (pVal.ItemUID == "creditActE" || pVal.ItemUID == "PrBaseE" || pVal.ItemUID == "95" || pVal.ItemUID == "234000005" || pVal.ItemUID == "5")
                        {
                            SAPbouiCOM.IChooseFromListEvent oCFLEvento = (SAPbouiCOM.IChooseFromListEvent)pVal;
                            chooseFromList(oForm, oCFLEvento, pVal.ItemUID, pVal.BeforeAction);
                        }
                    }
                    else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK)
                    {
                        if (pVal.BeforeAction)
                        {
                            if (pVal.ItemUID == "1")
                            {
                                //fillPhysicalEntityTaxes(oForm, out errorText);
                            }
                            else if (pVal.ItemUID == "10") //Item - Posting Date
                            {
                                oForm.Freeze(true);
                                oldAmountLC = Convert.ToDecimal(oForm.DataSources.DBDataSources.Item("OVPM").GetValue("TrsfrSum", 0), CultureInfo.InvariantCulture);
                                oldAmountFC = Convert.ToDecimal(oForm.DataSources.DBDataSources.Item("OVPM").GetValue("TrsfrSumFC", 0), CultureInfo.InvariantCulture);
                                oldDocRate = FormsB1.cleanStringOfNonDigits(oForm.DataSources.DBDataSources.Item("OVPM").GetValue("docRate", 0).ToString());

                                Program.paymentInvoices = new DataTable();
                                Program.paymentInvoices.Columns.Add("DocNum", typeof(string));
                                Program.paymentInvoices.Columns.Add("ObjType", typeof(string));
                                Program.paymentInvoices.Columns.Add("TotalPayment", typeof(decimal));

                                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("20").Specific;
                                int rowCount = oMatrix.RowCount;
                                int row = 1;

                                while (row <= rowCount)
                                {
                                    if (oMatrix.GetCellSpecific("10000127", row).Checked)
                                    {
                                        string docNum = oMatrix.GetCellSpecific("1", row).Value;
                                        string objType = oMatrix.GetCellSpecific("45", row).Value;
                                        decimal totalPayment = FormsB1.cleanStringOfNonDigits(oMatrix.GetCellSpecific("24", row).Value);

                                        var paymentInvRow = Program.paymentInvoices.NewRow();
                                        paymentInvRow["DocNum"] = docNum;
                                        paymentInvRow["ObjType"] = objType;
                                        paymentInvRow["TotalPayment"] = totalPayment;

                                        Program.paymentInvoices.Rows.Add(paymentInvRow);
                                    }
                                    row++;
                                }
                                oForm.Freeze(false);
                            }
                        }
                        else
                        {
                            if (pVal.ItemUID == "FillAmtTxs" && !pVal.InnerEvent)
                            {
                                fillAmountTaxes(oForm);
                            }
                            else if (pVal.ItemUID == "CalcPhcEnt")
                            {
                                CalcPhysicalEntityTaxes(oForm);
                            }
                        }
                    }

                    else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_RESIZE && !pVal.BeforeAction)
                    {
                        resizeForm(oForm);
                    }

                    else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_LOST_FOCUS && !pVal.BeforeAction)
                    {
                        //if (pVal.ItemUID == "13") //Item - Payment On Account (EditText)
                        //{
                        //    CalcPhysicalEntityTaxes(oForm, true);
                        //}
                    }

                    if (pVal.ItemChanged)
                    {
                        if (pVal.ItemUID == "10") //Item - Posting Date
                        {
                            string docEntry = oForm.DataSources.DBDataSources.Item("OVPM").GetValue("DocEntry", 0);
                            if (!string.IsNullOrEmpty(docEntry))
                            {
                                SAPbobsCOM.Payments oPayments = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPaymentsDrafts);
                                bool draft = oPayments.GetByKey(Convert.ToInt32(docEntry));

                                if (draft)
                                {
                                    if (!pVal.BeforeAction && oldAmountLC > 0)
                                    {
                                        changePostingDate(oForm);
                                    }
                                }
                            }
                            else
                            {
                                string docType = oForm.DataSources.DBDataSources.Item("OVPM").GetValue("DocType", 0).Trim();
                                if (docType != "A")
                                {
                                    oForm.Items.Item("234000005").Enabled = true;
                                    oForm.Items.Item("234000005").Specific.Value = "";
                                }
                                IncomingPayment.SetUsBlaAgRtSAvailability(oForm);
                            }
                        }
                        else if (pVal.ItemUID == "13" && !pVal.BeforeAction && !pVal.InnerEvent) //Item - Payment on Account (EditText)
                        {
                            CalcPhysicalEntityTaxes(oForm, true);
                        }
                        else if (pVal.ItemUID == "152" && !pVal.BeforeAction) //Item - WTax Base Sum (EditText)
                        {
                            CalcPhysicalEntityTaxes(oForm, true);
                        }
                    }
                }
            }
        }

        public static void CalcPhysicalEntityTaxes(SAPbouiCOM.Form oForm, bool isChangedWTaxBaseSum = false)
        {
            oForm.Freeze(true);
            SAPbobsCOM.SBObob oSBOBob = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);
            try
            {
                //var docEntry = oForm.DataSources.DBDataSources.Item("OVPM").GetValue("DocEntry", 0);
                //if (!string.IsNullOrEmpty(docEntry))
                //    return;
                string transId = oForm.DataSources.DBDataSources.Item("OVPM").GetValue("TransId", 0);
                if (!string.IsNullOrEmpty(transId))
                    return;

                var oDBDataSource = oForm.DataSources.DBDataSources.Item("OVPM");
                var oDBDataSourceBP = oForm.DataSources.DBDataSources.Item("OCRD");

                var isWTLiable = oDBDataSourceBP.GetValue("WTLiable", 0) == "Y";

                if (isWTLiable)
                {
                    var docDateStr = oDBDataSource.GetValue("DocDate", 0);
                    if (string.IsNullOrEmpty(docDateStr))
                        throw new Exception(BDOSResources.getTranslate("DocDate") + " " + BDOSResources.getTranslate("YouCantLeaveEmpty"));

                    DateTime docDate = DateTime.ParseExact(docDateStr, "yyyyMMdd", CultureInfo.InvariantCulture);
                    var docRate = Convert.ToDecimal(oDBDataSource.GetValue("DocRate", 0), CultureInfo.InvariantCulture);
                    var docCurr = oDBDataSource.GetValue("DocCurr", 0).Trim();
                    var isForeignCurrency = docCurr != Program.LocalCurrency;

                    decimal whTaxAmt = decimal.Zero;
                    decimal pensEmployedAmt = decimal.Zero; //დასაქმებული
                    decimal pensEmployerAmt = decimal.Zero; //დამსაქმებელი
                    decimal pensEmployerDiffAmt = decimal.Zero; //დამსაქმებლის საპენსიოების განსხვავების თანხა AP(AP Reserve) Invoice-სა და Payment-ს შორის

                    //invoices
                    if (!isChangedWTaxBaseSum)
                    {
                        SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("20").Specific;

                        for (var i = 1; i <= oMatrix.RowCount; i++)
                        {
                            if (oMatrix.GetCellSpecific("10000127", i).Checked)
                            {
                                var invDocNum = oMatrix.GetCellSpecific("1", i).Value;
                                var invType = oMatrix.GetCellSpecific("45", i).Value;
                                var invDocDate = DateTime.ParseExact(oMatrix.GetCellSpecific("21", i).Value, "yyyyMMdd", CultureInfo.InvariantCulture);
                                (string wtCode, decimal invDocRate)? invData = GetWTCodeFromInvoices(invType, invDocNum);

                                if (!invData.HasValue && string.IsNullOrEmpty(invData.Value.wtCode)) continue;

                                var invWTCode = invData.Value.wtCode;
                                var invWTaxAmt = FormsB1.cleanStringOfNonDigits(oMatrix.GetCellSpecific("72", i).Value);
                                string invTotalPaymentStr = oMatrix.GetCellSpecific("24", i).Value;
                                var invTotalPayment = FormsB1.cleanStringOfNonDigits(invTotalPaymentStr);
                                var invCurr = new string(invTotalPaymentStr.Where(char.IsLetter).ToArray());

                                var invGrossAmt = invWTaxAmt + invTotalPayment;

                                var invRate = 1.0m;
                                if (invCurr != Program.LocalCurrency)
                                    invRate = invCurr == docCurr ? docRate : Convert.ToDecimal(oSBOBob.GetCurrencyRate(invCurr, docDate).Fields.Item("CurrencyRate").Value, NumberFormatInfo.InvariantInfo);

                                (decimal whTaxAmt, decimal pensEmployedAmt, decimal pensEmployerAmt) physicalEntityTaxesAmt = CommonFunctions.CalcPhysicalEntityTaxes(invGrossAmt * invRate, docDate, invWTCode);
                                whTaxAmt += physicalEntityTaxesAmt.whTaxAmt;
                                pensEmployedAmt += physicalEntityTaxesAmt.pensEmployedAmt;
                                pensEmployerAmt += physicalEntityTaxesAmt.pensEmployerAmt;

                                var invOldDocRate = 1.0m;
                                var invOldPensEmployerAmtLC = 0.0m;
                                if (invType == "18" && invCurr != Program.LocalCurrency && invCurr == docCurr) //Only for AP Invoice || AP Reserve Invoice
                                {
                                    invOldDocRate = invData.Value.invDocRate;
                                    invOldPensEmployerAmtLC = physicalEntityTaxesAmt.pensEmployerAmt * invOldDocRate / docRate;
                                }

                                pensEmployerDiffAmt += invOldPensEmployerAmtLC > 0 ? physicalEntityTaxesAmt.pensEmployerAmt - invOldPensEmployerAmtLC : 0.0m;
                            }
                        }
                    }

                    //WTax Base Sum (on account)
                    var wTCode = oDBDataSource.GetValue("WtCode", 0);
                    if (!string.IsNullOrEmpty(wTCode))
                    {
                        var wTaxBaseSum = FormsB1.cleanStringOfNonDigits(oForm.Items.Item("152").Specific.Value);
                        if (wTaxBaseSum > 0)
                        {
                            (decimal whTaxAmt, decimal pensEmployedAmt, decimal pensEmployerAmt) physicalEntityTaxesWTaxBaseSum = CommonFunctions.CalcPhysicalEntityTaxes(wTaxBaseSum, docDate, wTCode);

                            double wTaxAmtDoc = Convert.ToDouble(CommonFunctions.roundAmountByGeneralSettings(physicalEntityTaxesWTaxBaseSum.whTaxAmt + physicalEntityTaxesWTaxBaseSum.pensEmployedAmt, "Sum"));
                            oForm.Items.Item("111").Specific.Value = oSBOBob.Format_MoneyToString(wTaxAmtDoc, SAPbobsCOM.BoMoneyPrecisionTypes.mpt_Sum).Fields.Item(0).Value;

                            whTaxAmt += physicalEntityTaxesWTaxBaseSum.whTaxAmt * (docRate > 0 ? docRate : 1);
                            pensEmployedAmt += physicalEntityTaxesWTaxBaseSum.pensEmployedAmt * (docRate > 0 ? docRate : 1);
                            pensEmployerAmt += physicalEntityTaxesWTaxBaseSum.pensEmployerAmt * (docRate > 0 ? docRate : 1);
                        }
                    }

                    if (!isChangedWTaxBaseSum)
                    {
                        oForm.Items.Item("BDOSPnPhAm").Specific.Value = FormsB1.ConvertDecimalToString(CommonFunctions.roundAmountByGeneralSettings(pensEmployedAmt, "Sum"));
                        oForm.Items.Item("BDOSPnCoAm").Specific.Value = FormsB1.ConvertDecimalToString(CommonFunctions.roundAmountByGeneralSettings(pensEmployerAmt, "Sum"));
                        oForm.Items.Item("PnCoDiffAm").Specific.Value = FormsB1.ConvertDecimalToString(CommonFunctions.roundAmountByGeneralSettings(pensEmployerDiffAmt, "Sum"));
                        oForm.Items.Item("BDOSWhtAmt").Specific.Value = FormsB1.ConvertDecimalToString(CommonFunctions.roundAmountByGeneralSettings(whTaxAmt, "Sum"));
                    }

                    oForm.Items.Item("26").Click();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                Marshal.ReleaseComObject(oSBOBob);
                oForm.Freeze(false);
            }
        }

        public static (string wtCode, decimal invDocRate)? GetWTCodeFromInvoices(string invType, string invDocNum)
        {
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                StringBuilder qury = new StringBuilder();
                qury.Append("SELECT T1.\"WTCode\", \n");
                qury.Append("T11.\"DocRate\" AS \"InvDocRate\" \n");
                qury.Append("FROM \"PCH5\" AS T1 \n");
                qury.Append("INNER JOIN \"OPCH\" T11 ON T1.\"AbsEntry\" = T11.\"DocEntry\" \n");
                qury.Append($"WHERE T1.\"ObjType\" = {invType} AND T11.\"DocNum\" = {invDocNum} \n");
                qury.Append("UNION \n");
                qury.Append("SELECT T2.\"WTCode\", \n");
                qury.Append("T22.\"DocRate\" AS \"InvDocRate\" \n");
                qury.Append("FROM \"DPO5\" AS T2 \n");
                qury.Append("INNER JOIN \"ODPO\" T22 ON T2.\"AbsEntry\" = T22.\"DocEntry\" \n");
                qury.Append($"WHERE T2.\"ObjType\" = {invType} AND T22.\"DocNum\" = {invDocNum} \n");
                qury.Append("UNION \n");
                qury.Append("SELECT T3.\"WTCode\", \n");
                qury.Append("T33.\"DocRate\" AS \"InvDocRate\" \n");
                qury.Append("FROM \"RPC5\" AS T3 \n");
                qury.Append("INNER JOIN \"ORPC\" T33 ON T3.\"AbsEntry\" = T33.\"DocEntry\" \n");
                qury.Append($"WHERE T3.\"ObjType\" = {invType} AND T33.\"DocNum\" = {invDocNum} ");

                oRecordSet.DoQuery(qury.ToString());

                if (!oRecordSet.EoF)
                    return (oRecordSet.Fields.Item("WTCode").Value, Convert.ToDecimal(oRecordSet.Fields.Item("InvDocRate").Value, CultureInfo.InvariantCulture));
                else
                    return null;
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

        private static void changePostingDate(SAPbouiCOM.Form oForm)
        {
            string docCurrency = oForm.DataSources.DBDataSources.Item("OVPM").GetValue("docCurr", 0).Trim();
            decimal docRate = FormsB1.cleanStringOfNonDigits(oForm.DataSources.DBDataSources.Item("OVPM").GetValue("docRate", 0).ToString());
            string transferAccount = oForm.DataSources.DBDataSources.Item("OVPM").GetValue("TrsfrAcct", 0).Trim();
            string trsfrAcctCurr = CommonFunctions.getAccountCurrency(transferAccount);
            string bpCurr = oForm.DataSources.DBDataSources.Item("OCRD").GetValue("Currency", 0).Trim();

            try
            {
                oForm.Freeze(true);

                Program.newPostDateStr = oForm.DataSources.DBDataSources.Item("OVPM").GetValue("DocDate", 0);
                if (docCurrency == Program.LocalCurrency)
                {
                    Program.transferSumFC = oldAmountLC;
                    Program.overallAmount = oldAmountLC;

                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)(oForm.Items.Item("20").Specific);
                    int rowCount = oMatrix.RowCount;

                    for (int i = 1; i <= rowCount; i++)
                    {
                        string docNum = oMatrix.GetCellSpecific("1", i).Value;
                        string objType = oMatrix.GetCellSpecific("45", i).Value;
                        DataRow[] foundRows = Program.paymentInvoices.Select("DocNum = '" + docNum + "' AND ObjType = '" + objType + "'");

                        if (foundRows.Count() > 0)
                        {
                            string totalPaymentFCStr = oMatrix.GetCellSpecific("24", i).Value;
                            string invCurr = totalPaymentFCStr.Substring(0, 3);

                            if (invCurr == docCurrency)
                            {
                                oMatrix.Columns.Item("10000127").Cells.Item(i).Specific.Checked = true;
                            }
                        }
                    }
                }

                else if (docCurrency != Program.LocalCurrency && trsfrAcctCurr == "##" && docRate != 0)
                {
                    Program.transferSumFC = CommonFunctions.roundAmountByGeneralSettings(oldAmountLC / docRate, "Sum");
                    if (bpCurr == Program.LocalCurrency || bpCurr == "##")
                    {
                        Program.overallAmount = CommonFunctions.roundAmountByGeneralSettings(oldAmountLC, "Sum");
                    }
                    else
                    {
                        Program.overallAmount = Program.transferSumFC;
                    }

                    decimal newAmountFC = CommonFunctions.roundAmountByGeneralSettings(oldAmountLC / docRate, "Sum");
                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)(oForm.Items.Item("20").Specific);
                    int rowCount = oMatrix.RowCount;

                    for (int i = 1; i <= rowCount; i++)
                    {
                        if (newAmountFC == 0)
                            break;

                        decimal totalPaymentFC;
                        string docNum = oMatrix.GetCellSpecific("1", i).Value;
                        string objType = oMatrix.GetCellSpecific("45", i).Value;
                        DataRow[] foundRows = Program.paymentInvoices.Select("DocNum = '" + docNum + "' AND ObjType = '" + objType + "'");

                        if (foundRows.Count() > 0)
                        {
                            string totalPaymentFCStr = oMatrix.GetCellSpecific("24", i).Value;
                            string invCurr = totalPaymentFCStr.Substring(0, 3);

                            if (invCurr == docCurrency)
                            {
                                decimal balanceDue = FormsB1.cleanStringOfNonDigits(oMatrix.GetCellSpecific("7", i).Value);
                                totalPaymentFC = Math.Min(balanceDue, newAmountFC);
                                newAmountFC = newAmountFC - totalPaymentFC;

                                oMatrix.Columns.Item("10000127").Cells.Item(i).Specific.Checked = true;
                                oMatrix.Columns.Item("24").Cells.Item(i).Specific.Value = FormsB1.ConvertDecimalToStringForEditboxStrings(totalPaymentFC);
                            }
                        }
                    }

                    oForm.Items.Item("21").Specific.value = FormsB1.ConvertDecimalToStringForEditboxStrings(CommonFunctions.roundAmountByGeneralSettings(oldAmountLC / Program.transferSumFC, "Rate", CommonFunctions.RoundingDirection.Down));

                    if (newAmountFC > 0)
                    {
                        decimal paymentOnAcct = CommonFunctions.roundAmountByGeneralSettings(newAmountFC * docRate, "Sum");
                        oForm.Items.Item("37").Specific.Checked = true;
                        oForm.Items.Item("13").Specific.value = FormsB1.ConvertDecimalToStringForEditboxStrings(paymentOnAcct);
                    }
                    else if (oForm.Items.Item("37").Specific.Checked)
                    {
                        oForm.Items.Item("37").Specific.Checked = false;
                    }
                }
                else
                {
                    Program.transferSumFC = oldAmountFC;

                    if (bpCurr == Program.LocalCurrency || bpCurr == "##")
                    {
                        Program.overallAmount = CommonFunctions.roundAmountByGeneralSettings(oldAmountFC * docRate, "Sum");
                    }
                    else
                    {
                        Program.overallAmount = oldAmountFC;
                    }

                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)(oForm.Items.Item("20").Specific);
                    int rowCount = oMatrix.RowCount;

                    for (int i = 1; i <= rowCount; i++)
                    {
                        string docNum = oMatrix.GetCellSpecific("1", i).Value;
                        string objType = oMatrix.GetCellSpecific("45", i).Value;
                        DataRow[] foundRows = Program.paymentInvoices.Select("DocNum = '" + docNum + "' AND ObjType = '" + objType + "'");

                        if (foundRows.Count() > 0)
                        {
                            string totalPaymentFCStr = oMatrix.GetCellSpecific("24", i).Value;
                            string invCurr = totalPaymentFCStr.Substring(0, 3);

                            if (invCurr == docCurrency)
                            {
                                oMatrix.Columns.Item("10000127").Cells.Item(i).Specific.Checked = true;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                var st = new System.Diagnostics.StackTrace(ex, true);
                // Get the top stack frame
                var frame = st.GetFrame(0);
                // Get the line number from the stack frame
                var line = frame.GetFileLineNumber();

                Program.uiApp.StatusBar.SetSystemMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                Program.uiApp.StatusBar.SetSystemMessage("st: " + st, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                Program.uiApp.StatusBar.SetSystemMessage("frame: " + frame, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                Program.uiApp.StatusBar.SetSystemMessage("line: " + line, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }

            Program.openPaymentMeansByPostDateChange = true;
            oForm.Items.Item("234000001").Click();
        }

        //private static void changeDocDateRate(SAPbouiCOM.Form oFormDate, string newDate)
        //{
        //    SAPbouiCOM.Form oForm = CurrentForm;
        //    DateTime newDocDate = Convert.ToDateTime(DateTime.ParseExact(newDate, "yyyyMMdd", CultureInfo.InvariantCulture));

        //    string docEntry = oForm.DataSources.DBDataSources.Item("OVPM").GetValue("DocEntry", 0);
        //    SAPbobsCOM.Payments oPayments = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPaymentsDrafts);

        //    if (oPayments.GetByKey(Convert.ToInt32(docEntry)))
        //    {
        //        decimal transferSum = Convert.ToDecimal(oPayments.TransferSum, CultureInfo.InvariantCulture);
        //        string docCurrency = oPayments.DocCurrency;
        //        SAPbobsCOM.SBObob oSBOBob = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);
        //        var docRate = oSBOBob.GetCurrencyRate(docCurrency, newDocDate).Fields.Item("CurrencyRate").Value;
        //        decimal docRateDcml = Convert.ToDecimal(docRate, CultureInfo.InvariantCulture);
        //        bool cashFlowRelevant = CommonFunctions.isAccountCashFlowRelevant(oPayments.TransferAccount);
        //        //decimal transferSumFC = CommonFunctions.roundAmountByGeneralSettings(transferSum / docRateDcml, "Sum");

        //        decimal transferSumFC = CommonFunctions.roundAmountByGeneralSettings(transferSum / Convert.ToDecimal(oPayments.DocRate, CultureInfo.InvariantCulture), "Sum");

        //        oPayments.DocDate = newDocDate;
        //        oPayments.TaxDate = newDocDate;
        //        oPayments.DueDate = newDocDate;
        //        oPayments.TransferDate = newDocDate;
        //        oPayments.VatDate = newDocDate;

        //        oPayments.DocRate = docRate;
        //        //oPayments.TransferSum = Convert.ToDouble(transferSumFC, NumberFormatInfo.InvariantInfo);

        //        for (int j = 0; j < oPayments.Invoices.Count; j++)
        //        {
        //            if (transferSumFC == 0)
        //                break;

        //            oPayments.Invoices.SetCurrentLine(j);

        //            decimal sumApplied = Convert.ToDecimal(oPayments.Invoices.SumApplied, NumberFormatInfo.InvariantInfo);
        //            decimal appliedFC = CommonFunctions.roundAmountByGeneralSettings(sumApplied / docRateDcml, "Sum");

        //            appliedFC = Math.Min(appliedFC, transferSumFC);


        //            oPayments.Invoices.AppliedFC = Convert.ToDouble(appliedFC, NumberFormatInfo.InvariantInfo);
        //            //oPayments.Invoices.SumApplied = Convert.ToDouble(sumApplied, NumberFormatInfo.InvariantInfo);

        //            transferSumFC = transferSumFC - appliedFC;
        //        }

        //        if (transferSumFC > 0)
        //        {
        //            if (oPayments.DocTypte == SAPbobsCOM.BoRcptTypes.rAccount)
        //            {
        //                oPayments.AccountPayments.SetCurrentLine(0);
        //                oPayments.AccountPayments.GrossAmount = Convert.ToDouble(transferSum - (transferSumFC * docRateDcml), NumberFormatInfo.InvariantInfo);
        //            }
        //            else
        //            {
        //                oPayments.DocTypte = SAPbobsCOM.BoRcptTypes.rAccount;
        //                //oPayments.AccountPayments.AccountCode = GLAccountCode;
        //                //oPayments.AccountPayments.ProjectCode = projectCod;
        //                oPayments.AccountPayments.GrossAmount = Convert.ToDouble(transferSumFC * docRateDcml, NumberFormatInfo.InvariantInfo);
        //                oPayments.AccountPayments.Add();
        //            }

        //            //oPayments.PayToBankAccountNo
        //        }

        //        int returnCode = oPayments.SaveDraftToDocument();
        //        if (returnCode != 0)
        //        {
        //            int errCode;
        //            string errMsg;
        //            Program.oCompany.GetLastError(out errCode, out errMsg);

        //            oPayments.Update();
        //        }

        //        //SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //        //SAPbobsCOM.Recordset oRecordSetPDF2 = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //        //if (bpCurrency == "##" || String.IsNullOrEmpty(bpCurrency))
        //        //{
        //        //    var invoicesDT = GetPaymentInvoices(Convert.ToInt32(docEntry), PaymentType.Draft, newDocDate);
        //        //    var currencies = invoicesDT.AsEnumerable().Select(x => x["DocCur"]);
        //        //    string firstcurrency = (string)currencies.FirstOrDefault();
        //        //    int otherCurrenciesCount = invoicesDT.AsEnumerable().Where(x => (string)x["DocCur"] != firstcurrency).Count();
        //        //    int firstCurrencyCount = invoicesDT.AsEnumerable().Where(x => (string)x["DocCur"] == firstcurrency).Count();

        //        //    if (otherCurrenciesCount > 0)
        //        //    {
        //        //        Program.uiApp.SetStatusBarMessage(BDOSResources.getTranslate("InvoicesDifferentCurrenciesError"), SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //        //        return;
        //        //    }
        //        //    else
        //        //    {
        //        //        DiffCurr = (docCurrency != firstcurrency) ? "Y" : "N";
        //        //        docCurrency = firstcurrency;
        //        //    }
        //        //}

        //        oFormDate.Close();
        //        oForm.Close();
        //        Program.uiApp.OpenForm((SAPbouiCOM.BoFormObjectEnum)140, "", docEntry);
        //        //SAPbouiCOM.Form oJournalForm = Program.uiApp.OpenForm((SAPbouiCOM.BoFormObjectEnum)140, "", docEntry);
        //    }
        //}

        //private static void changeDocDateRate(SAPbouiCOM.Form oFormDate, string newDate)
        //{
        //    //SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm("426", Program.currentFormCount);
        //    SAPbouiCOM.Form oForm = CurrentForm;
        //    DateTime DocDate = Convert.ToDateTime(DateTime.ParseExact(newDate, "yyyyMMdd", CultureInfo.InvariantCulture));
        //    string DocCurr = oForm.DataSources.DBDataSources.Item(0).GetValue("DocCurr", 0);
        //    string DocEntry = oForm.DataSources.DBDataSources.Item(0).GetValue("DocEntry", 0);
        //    decimal TrsfrSum = Convert.ToDecimal(oForm.DataSources.DBDataSources.Item(0).GetValue("TrsfrSum", 0), CultureInfo.InvariantCulture);
        //    decimal NoDocSumFC = Convert.ToDecimal(oForm.DataSources.DBDataSources.Item(0).GetValue("NoDocSumFC", 0), CultureInfo.InvariantCulture);
        //    string DiffCurr = oForm.DataSources.DBDataSources.Item(0).GetValue("DiffCurr", 0);

        //    string bpCurrency = oForm.DataSources.DBDataSources.Item(1).GetValue("Currency", 0);

        //    SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //    SAPbobsCOM.Recordset oRecordSetPDF2 = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //    if (bpCurrency == "##" || String.IsNullOrEmpty(bpCurrency))
        //    {
        //        var invoicesDT = GetPaymentInvoices(Convert.ToInt32(DocEntry), PaymentType.Draft, DocDate);
        //        var currencies = invoicesDT.AsEnumerable().Select(x => x["DocCur"]);
        //        string firstcurrency = (string)currencies.FirstOrDefault();
        //        int otherCurrenciesCount = invoicesDT.AsEnumerable().Where(x => (string)x["DocCur"] != firstcurrency).Count();
        //        int firstCurrencyCount = invoicesDT.AsEnumerable().Where(x => (string)x["DocCur"] == firstcurrency).Count();

        //        if (otherCurrenciesCount > 0)
        //        {
        //            Program.uiApp.SetStatusBarMessage(BDOSResources.getTranslate("InvoicesDifferentCurrenciesError"), SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //            return;
        //        }
        //        else
        //        {
        //            DiffCurr = (DocCurr != firstcurrency) ? "Y" : "N";
        //            DocCurr = firstcurrency;
        //        }
        //    }

        //    string errorText = null;
        //    decimal DocRate = 0;
        //    decimal TrsfrSumFC = 0;

        //    SAPbobsCOM.SBObob oSBOBob = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);

        //    if (DocCurr == CurrencyB1.getMainCurrency(out errorText))
        //    {
        //        string queryOPDF = @"update OPDF 
        //                            set 
        //                            ""DocDate"" = '" + newDate + @"',
        //                            ""TaxDate"" = '" + newDate + @"',
        //                            ""VatDate"" = '" + newDate + @"',
        //                            ""DocDueDate"" = '" + newDate + @"'

        //                            where ""DocEntry"" = " + DocEntry + @"";
        //        oRecordSet.DoQuery(queryOPDF);
        //    }
        //    else if (DiffCurr == "N")
        //    {
        //        SAPbobsCOM.Recordset RateRecordset = oSBOBob.GetCurrencyRate(DocCurr, DocDate);

        //        while (!RateRecordset.EoF)
        //        {
        //            DocRate = Convert.ToDecimal(RateRecordset.Fields.Item("CurrencyRate").Value);
        //            RateRecordset.MoveNext();
        //        }

        //        string queryOPDF = @"update OPDF 
        //                            set 
        //                            ""DocDate"" = '" + newDate + @"',
        //                            ""TaxDate"" = '" + newDate + @"',
        //                            ""VatDate"" = '" + newDate + @"',
        //                            ""DocDueDate"" = '" + newDate + @"',
        //                            ""DocRate"" = " + DocRate.ToString(CultureInfo.InvariantCulture) + @"
        //                            where ""DocEntry"" = " + DocEntry + @"";
        //        oRecordSet.DoQuery(queryOPDF);
        //    }
        //    else
        //    {
        //        SAPbobsCOM.Recordset RateRecordset = oSBOBob.GetCurrencyRate(DocCurr, DocDate);

        //        while (!RateRecordset.EoF)
        //        {
        //            DocRate = Convert.ToDecimal(RateRecordset.Fields.Item("CurrencyRate").Value);
        //            RateRecordset.MoveNext();
        //        }

        //        TrsfrSumFC = CommonFunctions.roundAmountByGeneralSettings(TrsfrSum / DocRate, "Sum");
        //        decimal TrsfrSumFCOld = TrsfrSumFC;

        //        string query = @"select * from PDF2
        //               where ""DocNum"" = " + DocEntry + @"";
        //        oRecordSet.DoQuery(query);

        //        while (!oRecordSet.EoF)
        //        {
        //            /////////////////////
        //            decimal AppliedFC = Convert.ToDecimal(oRecordSet.Fields.Item("AppliedFC").Value);
        //            string DocNum = oRecordSet.Fields.Item("DocNum").Value.ToString();
        //            string InvoiceDocEntry = oRecordSet.Fields.Item("DocEntry").Value.ToString();
        //            string InstId = oRecordSet.Fields.Item("InstId").Value.ToString();

        //            if (TrsfrSumFC == 0)
        //            {
        //                break;
        //            }

        //            if (AppliedFC > TrsfrSumFC)
        //            {
        //                AppliedFC = TrsfrSumFC;
        //            }

        //            TrsfrSumFC = TrsfrSumFC - AppliedFC;

        //            string queryPDF2 = @"update PDF2
        //            set 
        //            ""DocRate"" = " + DocRate.ToString(CultureInfo.InvariantCulture) + @",
        //            ""AppliedFC"" =" + AppliedFC.ToString(CultureInfo.InvariantCulture) + @",
        //            ""BfDcntSumF"" = " + AppliedFC.ToString(CultureInfo.InvariantCulture) + @",
        //            ""BfNetDcntF"" = " + AppliedFC.ToString(CultureInfo.InvariantCulture) + @"
        //            where ""DocNum"" = " + DocNum + @"
        //            and ""DocEntry"" =  " + InvoiceDocEntry + @"
        //            and ""InstId"" = " + InstId + "";

        //            oRecordSetPDF2.DoQuery(queryPDF2);
        //            /////////////////////

        //            oRecordSet.MoveNext();
        //        }

        //        query = @"update OPDF 
        //                            set 
        //                            ""DocDate"" = '" + newDate + @"',
        //                            ""TaxDate"" = '" + newDate + @"',
        //                            ""VatDate"" = '" + newDate + @"',
        //                            ""DocDueDate"" = '" + newDate + @"'" +
        //                            ((bpCurrency == "##") ? "" :
        //                            @",""DocRate"" = " + DocRate.ToString(CultureInfo.InvariantCulture) + @",
        //                            ""TrsfrSumFC"" = " + TrsfrSumFCOld.ToString(CultureInfo.InvariantCulture) + @",
        //                            ""DocTotalFC"" = " + TrsfrSumFCOld.ToString(CultureInfo.InvariantCulture)) +

        //                            @"where ""DocEntry"" = " + DocEntry + @"";
        //        oRecordSet.DoQuery(query);

        //        //დარჩენილი თანხით ინვოისების ჩახურვა
        //        if (TrsfrSumFC != 0)
        //        {
        //            DataTable Invoices = GetPaymentInvoices(Convert.ToInt32(DocEntry), PaymentType.Draft, DocDate);

        //            foreach (DataRow InvoicesRow in Invoices.Rows)
        //            {
        //                Decimal BalanceDue = Convert.ToDecimal(InvoicesRow["OpenAmountFC"]);
        //                Decimal AppliedFC = Convert.ToDecimal(InvoicesRow["AppliedFC"]);
        //                string DocNum = InvoicesRow["DocNum"].ToString();
        //                string InvoiceDocEntry = InvoicesRow["InvoiceDocEntry"].ToString();
        //                string InstlmntID = InvoicesRow["InstlmntID"].ToString();

        //                if (BalanceDue > AppliedFC)
        //                {
        //                    decimal AppliedFCNEW = Math.Min(BalanceDue, AppliedFC + TrsfrSumFC);

        //                    string queryPDF2 = @"update PDF2 set
        //                        ""AppliedFC"" = " + AppliedFCNEW.ToString(CultureInfo.InvariantCulture) + @",
        //                        ""BfDcntSumF"" = " + AppliedFCNEW.ToString(CultureInfo.InvariantCulture) + @",
        //                        ""BfNetDcntF"" = " + AppliedFCNEW.ToString(CultureInfo.InvariantCulture) + @"
        //                        where ""DocNum"" = " + DocNum + @"
        //                        and ""DocEntry"" = " + InvoiceDocEntry + @"
        //                        and ""InstId"" = " + InstlmntID;

        //                    oRecordSetPDF2.DoQuery(queryPDF2);
        //                    TrsfrSumFC = TrsfrSumFC - AppliedFCNEW + AppliedFC;
        //                }
        //            }
        //        }

        //        string noDocSumField = "NoDocSumFC";

        //        decimal onAccountSum = TrsfrSumFC;
        //        if ((bpCurrency == "##" || String.IsNullOrEmpty(bpCurrency)) && DocRate > 0)
        //        {
        //            noDocSumField = "NoDocSum";
        //            onAccountSum = onAccountSum * DocRate;
        //        }

        //        //დარჩენილი თანხის OnAccount-ზე გაშვება
        //        if (onAccountSum != 0)
        //        {
        //            query = @"update OPDF 
        //                set 
        //                " +
        //                $"\"{noDocSumField}\" = " + onAccountSum.ToString(CultureInfo.InvariantCulture) + @",
        //                ""PayNoDoc"" = 'Y'
        //                where ""DocEntry"" = " + DocEntry + @"";

        //            oRecordSet.DoQuery(query);
        //        }
        //        else
        //        {
        //            query = @"update OPDF 
        //                set 
        //                    " +
        //                $"\"{noDocSumField}\" = " + "0" + @",
        //                ""PayNoDoc"" = 'N'
        //                where ""DocEntry"" = " + DocEntry + @"";

        //            oRecordSet.DoQuery(query);
        //        }
        //    }

        //    Marshal.ReleaseComObject(oRecordSet);
        //    oRecordSet = null;

        //    Marshal.ReleaseComObject(oRecordSetPDF2);
        //    oRecordSetPDF2 = null;

        //    oFormDate.Close();
        //    oForm.Close();

        //    SAPbouiCOM.Form oJournalForm = Program.uiApp.OpenForm((SAPbouiCOM.BoFormObjectEnum)140, "", DocEntry);
        //}

        public static void createFormNewDate(SAPbouiCOM.Form oDocForm, out string errorText)
        {
            errorText = null;

            int left = 558 + 500;
            int Top = 200 + 300;

            //ფორმის აუცილებელი თვისებები
            Dictionary<string, object> formProperties = new Dictionary<string, object>();
            formProperties.Add("UniqueID", "OutgoingPaymentNewDate");
            formProperties.Add("BorderStyle", SAPbouiCOM.BoFormBorderStyle.fbs_Sizable);
            formProperties.Add("Title", BDOSResources.getTranslate("NewDate"));
            formProperties.Add("Left", left);
            formProperties.Add("Width", 200);
            formProperties.Add("Top", Top);
            formProperties.Add("Height", 10);
            formProperties.Add("Modality", SAPbouiCOM.BoFormModality.fm_Modal);

            SAPbouiCOM.Form oForm;
            bool newForm;
            bool formExist = FormsB1.createForm(formProperties, out oForm, out newForm, out errorText);

            if (formExist)
            {
                if (newForm)
                {
                    //ფორმის ელემენტების თვისებები
                    Dictionary<string, object> formItems = null;

                    Top = 1;
                    left = 6;


                    formItems = new Dictionary<string, object>();
                    string itemName = "newDate";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_DATE);
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
                    formItems.Add("ValueEx", DateTime.Now.ToString("yyyyMMdd"));

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    Top = Top + 19 + 5;
                    left = 6;

                    itemName = "1";
                    formItems = new Dictionary<string, object>();
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 120);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("Update"));

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }
                }

                oForm.Visible = true;
                //oForm.Select();
            }


            GC.Collect();


        }

        public static void fillAmountTaxes(SAPbouiCOM.Form oForm)
        {
            oForm.Freeze(true);
            try
            {
                if (oForm.TypeEx == "196") return; //Payments Means

                double AmountPr = 0;

                SAPbouiCOM.Item oItemPrTx = oForm.Items.Item("AmtPrTxE");
                oItemPrTx.Enabled = true;

                string liablePrTx = oForm.DataSources.DBDataSources.Item("OVPM").GetValue("U_liablePrTx", 0).Trim();

                double profitTaxRate = Convert.ToDouble(ProfitTax.GetProfitTaxRate());

                SAPbouiCOM.Item oItemNoDocSum = oForm.Items.Item("13");
                SAPbouiCOM.EditText oEditNoDocSum = (SAPbouiCOM.EditText)oItemNoDocSum.Specific;
                string noDocSumTx = oEditNoDocSum.Value;

                double noDocSum = Convert.ToDouble(noDocSumTx, CultureInfo.InvariantCulture);
                if (liablePrTx == "Y" && noDocSum > 0)
                {
                    AmountPr = AmountPr + Convert.ToDouble(CommonFunctions.roundAmountByGeneralSettings(Convert.ToDecimal(noDocSum / (100 - profitTaxRate) * profitTaxRate), "Sum"));
                }

                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("20").Specific;
                for (int i = 1; i < oMatrix.RowCount + 1; i++)
                {
                    string Payment = FormsB1.cleanStringOfNonDigits(oMatrix.Columns.Item("24").Cells.Item(i).Specific.Value).ToString();
                    decimal wTaxAmt = FormsB1.cleanStringOfNonDigits(oMatrix.Columns.Item("72").Cells.Item(i).Specific.Value);
                    string DocType = oMatrix.Columns.Item("45").Cells.Item(i).Specific.Value.ToString();
                    string DocNum = oMatrix.Columns.Item("1").Cells.Item(i).Specific.Value.ToString();
                    string Selected = oMatrix.Columns.Item("10000127").Cells.Item(i).Specific.Caption;

                    if (Selected == "Y" && (DocType == "204" || DocType == "18"))
                    {
                        bool prTx = GetNonEconExpAP(Convert.ToInt16(DocNum), DocType);
                        if (prTx)
                        {
                            AmountPr += Convert.ToDouble(CommonFunctions.roundAmountByGeneralSettings(
                                (Convert.ToDecimal(Payment) + wTaxAmt) / Convert.ToDecimal(100 - profitTaxRate) *
                                Convert.ToDecimal(profitTaxRate), "Sum"));
                        }
                    }
                }

                SAPbouiCOM.EditText oEditAmtPrTx = (SAPbouiCOM.EditText)oItemPrTx.Specific;

                oEditAmtPrTx.Value = FormsB1.ConvertDecimalToString(Convert.ToDecimal(AmountPr));

                oForm.Items.Item("26").Click();
                oItemPrTx.Enabled = false;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                oForm.Update();
                oForm.Freeze(false);
            }
        }

        //public static void fillPhysicalEntityTaxes(SAPbouiCOM.Form oForm, out string errorText)
        //{
        //    errorText = null;
        //    oForm.Freeze(true);

        //    try
        //    {
        //        SAPbouiCOM.DBDataSources docDBSources = oForm.DataSources.DBDataSources;

        //        string wtCode = oForm.Items.Item("110").Specific.Value.ToString();

        //        bool physicalEntityTax = (oForm.DataSources.DBDataSources.Item("OCRD").GetValue("WTLiable", 0) == "Y" &&
        //                        CommonFunctions.getValue("OWHT", "U_BDOSPhisTx", "WTCode", wtCode).ToString() == "Y");

        //        string docDatestr = docDBSources.Item("OVPM").GetValue("DocDate", 0);
        //        if (physicalEntityTax && string.IsNullOrEmpty(docDatestr))
        //        {
        //            errorText = BDOSResources.getTranslate("DocDate") + " " + BDOSResources.getTranslate("YouCantLeaveEmpty");
        //            return;
        //        }

        //        DateTime DocDate = DateTime.ParseExact(docDatestr, "yyyyMMdd", CultureInfo.InvariantCulture);
        //        Dictionary<string, decimal> PhysicalEntityPensionRates = WithholdingTax.GetPhysicalEntityPensionRates(DocDate, wtCode, out var errorTextCheck);

        //        if (physicalEntityTax && !string.IsNullOrEmpty(errorTextCheck))
        //        {
        //            errorText = errorTextCheck;
        //            return;
        //        }

        //        decimal TotalPensPhAm = oForm.Items.Item("111").Specific.Value;
        //        decimal TotalWhtAmt = 0;
        //        decimal TotalPensCoAm = 0;

        //        decimal PensPhAm = 0;
        //        decimal WhtAmt = 0;
        //        decimal PensCoAm = 0;
        //        decimal GrossAmount = 0;
        //        decimal GrossAmountFC = 0;

        //        SAPbouiCOM.Item oItemTxVal = oForm.Items.Item("13");
        //        SAPbouiCOM.EditText oEditTxVal = (SAPbouiCOM.EditText)oItemTxVal.Specific;
        //        string amtTxVal = oEditTxVal.Value;
        //        string wtCode2 = oForm.Items.Item("110").Specific.Value;
        //        bool onAccount = oForm.Items.Item("37").Specific.Checked;
        //        double rate = 0;
        //        bool lineIsChecked = false;
        //        SAPbouiCOM.Matrix oMatrix1 = oForm.Items.Item("20").Specific;
        //        decimal WhtAmtt = 0;
        //        decimal PnPhAmt = 0;
        //        decimal PnCoAm = 0;
        //        for (int row = 1; row <= oMatrix1.RowCount; row++)
        //        {
        //            SAPbouiCOM.CheckBox Edtfield = oMatrix1.Columns.Item("10000127").Cells.Item(row).Specific;
        //            bool checkedLine = (Edtfield.Checked);
        //            if (checkedLine && onAccount)
        //            {
        //                lineIsChecked = true;
        //                int docNum = (int)FormsB1.cleanStringOfNonDigits(oMatrix1.Columns.Item("1").Cells.Item(row).Specific.Value);
        //                SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //                string query = "select \"PCH1\".\"U_BDOSWhtAmt\", \"PCH1\".\"U_BDOSPnPhAm\", \"PCH1\".\"U_BDOSPnCoAm\" From \"PCH1\" " + "\n"
        //                + "join \"OPCH\" on \"OPCH\".\"DocEntry\" = \"PCH1\".\"DocEntry\" " + "\n"
        //                + "where \"OPCH\".\"DocNum\" = '" + docNum + "'";

        //                oRecordSet.DoQuery(query);

        //                while (!oRecordSet.EoF)
        //                {
        //                    WhtAmtt += Convert.ToDecimal((oRecordSet.Fields.Item("U_BDOSWhtAmt").Value));
        //                    PnPhAmt += (decimal)oRecordSet.Fields.Item("U_BDOSPnPhAm").Value;
        //                    PnCoAm += (decimal)oRecordSet.Fields.Item("U_BDOSPnCoAm").Value;

        //                    TotalPensPhAm += PnPhAmt;
        //                    TotalWhtAmt += WhtAmtt;
        //                    TotalPensCoAm += PnCoAm;

        //                    oRecordSet.MoveNext();
        //                }
        //                if (isPension(wtCode2, out rate))
        //                {
        //                    decimal payOnAcct = Convert.ToDecimal(oForm.Items.Item("13").Specific.Value);
        //                    decimal PhEnPens = Convert.ToDecimal(payOnAcct * 2 / 100);
        //                    decimal compPens = Convert.ToDecimal(payOnAcct * 2 / 100);
        //                    decimal whtax = Convert.ToDecimal((payOnAcct - PhEnPens) * (decimal)(rate) / 100);

        //                    TotalPensPhAm += PhEnPens;
        //                    TotalWhtAmt += whtax;
        //                    TotalPensCoAm += compPens;

        //                    SAPbouiCOM.EditText oEdit1 = oForm.Items.Item("BDOSWhtAmt").Specific;
        //                    oEdit1.Value = FormsB1.ConvertDecimalToString(whtax);

        //                    oEdit1 = oForm.Items.Item("BDOSPnPhAm").Specific;
        //                    oEdit1.Value = FormsB1.ConvertDecimalToString(PhEnPens);

        //                    oEdit1 = oForm.Items.Item("BDOSPnCoAm").Specific;
        //                    oEdit1.Value = FormsB1.ConvertDecimalToString(compPens);

        //                    oForm.Items.Item("26").Click();
        //                }
        //                else
        //                {
        //                    decimal payOnAcct = Convert.ToDecimal(oForm.Items.Item("13").Specific.Value);
        //                    decimal whtax = Convert.ToDecimal(payOnAcct * 20 / 100);

        //                    SAPbouiCOM.EditText oEdit1 = oForm.Items.Item("BDOSWhtAmt").Specific;
        //                    oEdit1.Value = FormsB1.ConvertDecimalToString(whtax);
        //                    oEdit1 = oForm.Items.Item("BDOSPnPhAm").Specific;
        //                    oEdit1.Value = FormsB1.ConvertDecimalToString(0);

        //                    oEdit1 = oForm.Items.Item("BDOSPnCoAm").Specific;
        //                    oEdit1.Value = FormsB1.ConvertDecimalToString(0);
        //                }
        //            }
        //        }


        //        if (onAccount && !lineIsChecked)
        //        {
        //            if (isPension(wtCode2, out rate))
        //            {
        //                decimal payOnAcct = Convert.ToDecimal(oForm.Items.Item("13").Specific.Value);
        //                decimal PhEnPens = Convert.ToDecimal(payOnAcct * 2 / 100);
        //                decimal compPens = Convert.ToDecimal(payOnAcct * 2 / 100);
        //                decimal whtax = Convert.ToDecimal((payOnAcct - PhEnPens) * (decimal)(rate) / 100);

        //                TotalPensPhAm += PhEnPens;
        //                TotalWhtAmt += whtax;
        //                TotalPensCoAm += compPens;

        //                SAPbouiCOM.EditText oEdit1 = oForm.Items.Item("BDOSWhtAmt").Specific;
        //                oEdit1.Value = FormsB1.ConvertDecimalToString(whtax);

        //                oEdit1 = oForm.Items.Item("BDOSPnPhAm").Specific;
        //                oEdit1.Value = FormsB1.ConvertDecimalToString(PhEnPens);

        //                oEdit1 = oForm.Items.Item("BDOSPnCoAm").Specific;
        //                oEdit1.Value = FormsB1.ConvertDecimalToString(compPens);

        //                oForm.Items.Item("26").Click();
        //            } else
        //            {
        //                decimal payOnAcct = Convert.ToDecimal(oForm.Items.Item("13").Specific.Value);
        //                decimal whtax = Convert.ToDecimal(payOnAcct * 20 / 100);

        //                SAPbouiCOM.EditText oEdit1 = oForm.Items.Item("BDOSWhtAmt").Specific;
        //                oEdit1.Value = FormsB1.ConvertDecimalToString(whtax);
        //                oEdit1 = oForm.Items.Item("BDOSPnPhAm").Specific;
        //                oEdit1.Value = FormsB1.ConvertDecimalToString(0);

        //                oEdit1 = oForm.Items.Item("BDOSPnCoAm").Specific;
        //                oEdit1.Value = FormsB1.ConvertDecimalToString(0);
        //            }
        //        }

        //        if (physicalEntityTax)
        //        {
        //            if (physicalEntityTax && amtTxVal != "" && docDBSources.Item("OVPM").GetValue("WtCode", 0).ToString().Trim() != "")
        //            {
        //                bool frgn = docDBSources.Item("OVPM").GetValue("DocCurr", 0).Trim() != CommonFunctions.getLocalCurrency();

        //                GrossAmount = Convert.ToDecimal(amtTxVal, CultureInfo.InvariantCulture);

        //                PensPhAm = CommonFunctions.roundAmountByGeneralSettings(GrossAmount * PhysicalEntityPensionRates["PensionWTaxRate"] / 100, "Sum");
        //                WhtAmt = CommonFunctions.roundAmountByGeneralSettings((GrossAmount - PensPhAm) * PhysicalEntityPensionRates["WTRate"] / 100, "Sum");
        //                PensCoAm = CommonFunctions.roundAmountByGeneralSettings(GrossAmount * PhysicalEntityPensionRates["PensionCoWTaxRate"] / 100, "Sum");

        //                SAPbouiCOM.EditText oEditWTax = oForm.Items.Item("111").Specific;
        //                oEditWTax.Value = FormsB1.ConvertDecimalToStringForEditboxStrings(PensPhAm + WhtAmt);

        //                if (frgn)
        //                {
        //                    GrossAmountFC = GrossAmount * Convert.ToDecimal(CommonFunctions.getChildOrDbDataSourceValue(docDBSources.Item("OVPM"), null, null, "DocRate", 0), CultureInfo.InvariantCulture);

        //                    TotalPensPhAm = CommonFunctions.roundAmountByGeneralSettings(GrossAmountFC * PhysicalEntityPensionRates["PensionWTaxRate"] / 100, "Sum");
        //                    TotalWhtAmt = CommonFunctions.roundAmountByGeneralSettings((GrossAmountFC - TotalPensPhAm) * PhysicalEntityPensionRates["WTRate"] / 100, "Sum");
        //                    TotalPensCoAm = CommonFunctions.roundAmountByGeneralSettings(GrossAmountFC * PhysicalEntityPensionRates["PensionCoWTaxRate"] / 100, "Sum");
        //                }
        //                else
        //                {
        //                    TotalPensPhAm = PensPhAm;
        //                    TotalWhtAmt = WhtAmt;
        //                    TotalPensCoAm = PensCoAm;
        //                }
        //            }

        //            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("20").Specific;
        //            SAPbouiCOM.Columns oColumns = oMatrix.Columns;
        //            string DocType;
        //            decimal wtaxInv;

        //            for (int i = 1; i < oMatrix.RowCount + 1; i++)
        //            {
        //                if (oColumns.Item("10000127").Cells.Item(i).Specific.Caption == "Y")
        //                {
        //                    docDatestr = oColumns.Item("21").Cells.Item(i).Specific.Value.ToString();
        //                    DocDate = DateTime.ParseExact(docDatestr, "yyyyMMdd", CultureInfo.InvariantCulture);
        //                    PhysicalEntityPensionRates = WithholdingTax.GetPhysicalEntityPensionRates(DocDate, wtCode, out errorTextCheck);

        //                    DocType = oColumns.Item("45").Cells.Item(i).Specific.Value.ToString().Trim();
        //                    GrossAmount = Convert.ToDecimal(FormsB1.cleanStringOfNonDigits(oColumns.Item("24").Cells.Item(i).Specific.Value), CultureInfo.InvariantCulture);
        //                    wtaxInv = Convert.ToDecimal(FormsB1.cleanStringOfNonDigits(oColumns.Item("72").Cells.Item(i).Specific.Value), CultureInfo.InvariantCulture);

        //                    decimal CurRate = Convert.ToDecimal(FormsB1.cleanStringOfNonDigits(oColumns.Item("41").Cells.Item(i).Specific.Value), CultureInfo.InvariantCulture);
        //                    if (CurRate != 0)
        //                    {
        //                        GrossAmount = GrossAmount * CurRate;
        //                        wtaxInv = wtaxInv * CurRate;
        //                    }

        //                    if (DocType == "19")
        //                    {
        //                        GrossAmount = GrossAmount * (-1);
        //                        wtaxInv = wtaxInv * (-1);
        //                    }

        //                    if ((DocType == "18" || DocType == "19" || DocType == "204") && wtaxInv != 0 && PhysicalEntityPensionRates["PensionWTaxRate"] != 0)
        //                    {
        //                        PensPhAm = CommonFunctions.roundAmountByGeneralSettings(wtaxInv * 100 * PhysicalEntityPensionRates["PensionWTaxRate"] / (100 * PhysicalEntityPensionRates["PensionWTaxRate"] + PhysicalEntityPensionRates["WTRate"] * (100 - PhysicalEntityPensionRates["PensionWTaxRate"])), "Sum");
        //                        TotalPensPhAm = TotalPensPhAm + PensPhAm;
        //                        if (PensPhAm != 0)
        //                        {
        //                            TotalWhtAmt = TotalWhtAmt + (Convert.ToDecimal(wtaxInv, CultureInfo.InvariantCulture) - PensPhAm);
        //                        }
        //                        TotalPensCoAm = TotalPensCoAm + CommonFunctions.roundAmountByGeneralSettings((GrossAmount + wtaxInv) * PhysicalEntityPensionRates["PensionCoWTaxRate"] / 100, "Sum");
        //                    }
        //                }
        //            }
        //        }

        //        for (int row = 1; row <= oMatrix1.RowCount; row++)
        //        {
        //            SAPbouiCOM.CheckBox Edtfield = oMatrix1.Columns.Item("10000127").Cells.Item(row).Specific;
        //            bool checkedLine = (Edtfield.Checked);
        //            if (checkedLine && !onAccount)
        //            {
        //                int docNum = (int)FormsB1.cleanStringOfNonDigits(oMatrix1.Columns.Item("1").Cells.Item(row).Specific.Value);
        //                SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //                string query = "select \"PCH1\".\"U_BDOSWhtAmt\", \"PCH1\".\"U_BDOSPnPhAm\", \"PCH1\".\"U_BDOSPnCoAm\" From \"PCH1\" " + "\n"
        //                + "join \"OPCH\" on \"OPCH\".\"DocEntry\" = \"PCH1\".\"DocEntry\" " + "\n"
        //                + "where \"OPCH\".\"DocNum\" = '" + docNum + "'";

        //                oRecordSet.DoQuery(query);

        //                while (!oRecordSet.EoF)
        //                {
        //                    WhtAmtt += Convert.ToDecimal((oRecordSet.Fields.Item("U_BDOSWhtAmt").Value));
        //                    PnPhAmt += (decimal)oRecordSet.Fields.Item("U_BDOSPnPhAm").Value;
        //                    PnCoAm += (decimal)oRecordSet.Fields.Item("U_BDOSPnCoAm").Value;

        //                    TotalPensPhAm += PnPhAmt;
        //                    TotalWhtAmt += WhtAmtt;
        //                    TotalPensCoAm += PnCoAm;

        //                    SAPbouiCOM.EditText oEdit1 = oForm.Items.Item("BDOSWhtAmt").Specific;
        //                    oEdit1.Value = FormsB1.ConvertDecimalToString(WhtAmtt);

        //                    oEdit1 = oForm.Items.Item("BDOSPnPhAm").Specific;
        //                    oEdit1.Value = FormsB1.ConvertDecimalToString(PnPhAmt);

        //                    oEdit1 = oForm.Items.Item("BDOSPnCoAm").Specific;
        //                    oEdit1.Value = FormsB1.ConvertDecimalToString(PnCoAm);

        //                    oForm.Items.Item("26").Click();

        //                    oRecordSet.MoveNext();
        //                }
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        errorText = ex.Message;
        //    }
        //    finally
        //    {
        //        oForm.Freeze(false);
        //        oForm.Update();
        //        GC.Collect();
        //    }
        //}

        public static bool GetNonEconExpAP(int DocNum, string DocType)
        {
            NumberFormatInfo Nfi = new NumberFormatInfo() { NumberDecimalSeparator = "." };
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string query = "";
            if (DocType == "18")
            {
                query = @"SELECT ""OPCH"".""U_nonEconExp"" AS ""PrTx""
                        FROM ""OPCH""
                        WHERE ""OPCH"".""DocNum""='" + DocNum + "'";
            }
            else if (DocType == "204")
            {
                query = @"SELECT ""ODPO"".""U_liablePrTx"" AS ""PrTx""
                        FROM ""ODPO""
                        WHERE ""ODPO"".""DocNum""='" + DocNum + "'";
            }

            oRecordSet.DoQuery(query);

            if (!oRecordSet.EoF)
            {
                return (oRecordSet.Fields.Item("PrTx").Value == "Y");
            }
            else
            {
                return false;
            }
        }

        public static string createDocumentTransferToOwnAccountType(SAPbouiCOM.DataTable oDataTable, int i, out int docEntry, out int docNum, out string errorText)
        {
            errorText = null;
            docEntry = 0;
            docNum = 0;
            SAPbobsCOM.Payments oPayments = null;
            oPayments = (SAPbobsCOM.Payments)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments);
            SAPbobsCOM.Payments oPaymentsNew = null;
            oPaymentsNew = (SAPbobsCOM.Payments)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments);
            SAPbobsCOM.SBObob oSBOBob = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);

            try
            {
                string localCurrency = CommonFunctions.getLocalCurrency();

                DateTime docDate = oDataTable.GetValue("DocumentDate", i);
                DateTime valueDate = oDataTable.GetValue("ValueDate", i);
                string GLAccountCode = oDataTable.GetValue("GLAccountCode", i);
                string projectCod = oDataTable.GetValue("Project", i);
                if (string.IsNullOrEmpty(GLAccountCode))
                    errorText = BDOSResources.getTranslate("TheFollowingFieldIsMandatory") + " : \"" + BDOSResources.getTranslate("GLAccountCode") + "\"! ";
                string cashFlowLineItemName = oDataTable.GetValue("CashFlowLineItemName", i);
                string accountNumber = oDataTable.GetValue("AccountNumber", i);
                string currency = oDataTable.GetValue("Currency", i);
                string currencySapCode = CommonFunctions.getCurrencySapCode(currency);
                if (string.IsNullOrEmpty(currencySapCode))
                    errorText = errorText + BDOSResources.getTranslate("CouldNotFindCurrency") + " \"" + currency + "\"! ";
                string partnerAccountNumber = oDataTable.GetValue("PartnerAccountNumber", i);
                string partnerCurrency = oDataTable.GetValue("PartnerCurrency", i);
                string partnerCurrencySapCode = CommonFunctions.getCurrencySapCode(partnerCurrency);
                if (string.IsNullOrEmpty(partnerCurrencySapCode))
                    errorText = errorText + BDOSResources.getTranslate("CouldNotFindCurrency") + " \"" + partnerCurrencySapCode + "\"! ";
                string transferAccount = CommonFunctions.getTransferAccount(accountNumber + currency);
                if (string.IsNullOrEmpty(transferAccount))
                    errorText = errorText + BDOSResources.getTranslate("CheckGLAccountForHouseBankAccount") + " \"" + accountNumber + currency + "\"! ";
                bool cashFlowRelevant = CommonFunctions.isAccountCashFlowRelevant(transferAccount);
                string cashFlowLineItemID = oDataTable.GetValue("CashFlowLineItemID", i);
                if (cashFlowRelevant && string.IsNullOrEmpty(cashFlowLineItemID))
                    errorText = errorText + BDOSResources.getTranslate("TheFollowingFieldIsMandatory") + " : \"" + BDOSResources.getTranslate("CashFlowLineItemID") + "\"! ";
                if (CommonFunctions.isAccountInHouseBankAccount(partnerAccountNumber + partnerCurrency) == false)
                    errorText = errorText + BDOSResources.getTranslate("CouldNotFindHouseBankAccount") + " \"" + partnerAccountNumber + partnerCurrency + "\"! ";

                if (string.IsNullOrEmpty(errorText) == false)
                {
                    errorText = BDOSResources.getTranslate("OperationCompletedUnSuccessfully") + "! " + errorText + BDOSResources.getTranslate("TableRow") + " : " + (i + 1);
                    return null;
                }

                oPayments.ProjectCode = projectCod;
                oPayments.DocObjectCode = SAPbobsCOM.BoPaymentsObjectType.bopot_OutgoingPayments;
                oPayments.DocTypte = SAPbobsCOM.BoRcptTypes.rAccount;
                oPayments.DocDate = docDate;
                oPayments.TaxDate = docDate;

                decimal docRate;
                decimal transferSumLC;
                decimal transferSumFC;
                decimal grossAmount;
                decimal amount = Convert.ToDecimal(oDataTable.GetValue("Amount", i), NumberFormatInfo.InvariantInfo);

                if ((currencySapCode == partnerCurrencySapCode) && currencySapCode == localCurrency)
                {
                    oPayments.DocCurrency = currencySapCode;
                    oPayments.LocalCurrency = SAPbobsCOM.BoYesNoEnum.tNO;
                    oPayments.DocRate = 0;
                    transferSumLC = amount;
                    transferSumFC = 0;
                    grossAmount = amount;
                }
                else if ((currencySapCode == partnerCurrencySapCode) && currencySapCode != localCurrency)
                {
                    oPayments.DocCurrency = currencySapCode;
                    oPayments.LocalCurrency = SAPbobsCOM.BoYesNoEnum.tNO;
                    oPayments.DocRate = oSBOBob.GetCurrencyRate(currencySapCode, docDate).Fields.Item("CurrencyRate").Value;
                    docRate = Convert.ToDecimal(oPayments.DocRate);
                    transferSumLC = amount * docRate;
                    transferSumFC = amount;
                    grossAmount = amount;
                }
                else
                {
                    errorText = BDOSResources.getTranslate("OperationCompletedUnSuccessfully") + "! " + BDOSResources.getTranslate("CheckCurrencies") + "! " + errorText + BDOSResources.getTranslate("TableRow") + " : " + (i + 1);
                    return null;
                }

                transferSumLC = CommonFunctions.roundAmountByGeneralSettings(transferSumLC, "Sum");
                transferSumFC = CommonFunctions.roundAmountByGeneralSettings(transferSumFC, "Sum");
                grossAmount = CommonFunctions.roundAmountByGeneralSettings(grossAmount, "Sum");
                amount = CommonFunctions.roundAmountByGeneralSettings(amount, "Sum");

                oPayments.TransferAccount = transferAccount;
                oPayments.TransferDate = docDate;
                oPayments.TransferSum = Convert.ToDouble(amount, NumberFormatInfo.InvariantInfo);

                if (CommonFunctions.IsDevelopment())
                {
                    oPayments.UserFields.Fields.Item("U_BDOSBdgCf").Value = oDataTable.GetValue("BudgetCashFlowID", i);
                    oPayments.UserFields.Fields.Item("U_BDOSBdgCfN").Value = oDataTable.GetValue("BudgetCashFlowName", i);
                }

                oPayments.UserFields.Fields.Item("U_opType").Value = "transferToOwnAccount";
                oPayments.UserFields.Fields.Item("U_status").Value = "downloadedFromTheBank";
                oPayments.UserFields.Fields.Item("U_paymentID").Value = oDataTable.GetValue("PaymentID", i);
                oPayments.UserFields.Fields.Item("U_creditAcct").Value = partnerAccountNumber + partnerCurrency;
                oPayments.UserFields.Fields.Item("U_descrpt").Value = oDataTable.GetValue("Description", i);
                oPayments.UserFields.Fields.Item("U_addDescrpt").Value = oDataTable.GetValue("AdditionalDescription", i);

                oPayments.UserFields.Fields.Item("U_docNumber").Value = oDataTable.GetValue("DocumentNumber", i);
                oPayments.UserFields.Fields.Item("U_transCode").Value = oDataTable.GetValue("TransactionCode", i);
                oPayments.UserFields.Fields.Item("U_ePaymentID").Value = oDataTable.GetValue("ExternalPaymentID", i);
                oPayments.UserFields.Fields.Item("U_opCode").Value = oDataTable.GetValue("OperationCode", i);

                //ცხრილური ნაწილი
                oPayments.AccountPayments.AccountCode = GLAccountCode;
                oPayments.AccountPayments.ProjectCode = projectCod;
                oPayments.AccountPayments.GrossAmount = Convert.ToDouble(grossAmount, NumberFormatInfo.InvariantInfo);
                oPayments.AccountPayments.Add();

                if (cashFlowRelevant)
                {
                    oPayments.PrimaryFormItems.CashFlowLineItemID = Convert.ToInt32(cashFlowLineItemID);
                    oPayments.PrimaryFormItems.AmountFC = Convert.ToDouble(transferSumFC, NumberFormatInfo.InvariantInfo);
                    if (oPayments.DocCurrency == localCurrency)
                        oPayments.PrimaryFormItems.AmountLC = Convert.ToDouble(transferSumLC, NumberFormatInfo.InvariantInfo);
                    oPayments.PrimaryFormItems.PaymentMeans = SAPbobsCOM.PaymentMeansTypeEnum.pmtBankTransfer;
                    oPayments.PrimaryFormItems.Add();
                }

                int returnCode = oPayments.Add();
                if (returnCode != 0)
                {
                    int errCode;
                    string errMsg;
                    Program.oCompany.GetLastError(out errCode, out errMsg);
                    errorText = BDOSResources.getTranslate("OperationCompletedUnSuccessfully") + "! " + errMsg + "! " + BDOSResources.getTranslate("TableRow") + " : " + (i + 1);
                    return null;
                }
                else
                {
                    bool newDoc = oPaymentsNew.GetByKey(Convert.ToInt32(Program.oCompany.GetNewObjectKey()));
                    if (newDoc)
                    {
                        docEntry = oPaymentsNew.DocEntry;
                        docNum = oPaymentsNew.DocNum;
                        oDataTable.SetValue("DocEntry", i, docEntry.ToString());
                        oDataTable.SetValue("DocNum", i, docNum.ToString());
                        return BDOSResources.getTranslate("OperationCompletedSuccessfully") + "! " + BDOSResources.getTranslate("TableRow") + " : " + (i + 1);
                    }
                    else
                        return "";
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
                return null;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oPayments);
                oPayments = null;
                Marshal.FinalReleaseComObject(oPaymentsNew);
                oPaymentsNew = null;
                Marshal.FinalReleaseComObject(oSBOBob);
                oSBOBob = null;
            }
        }

        public static string createDocumentTransferToBPType(SAPbouiCOM.DataTable oDataTable, int i, out int docEntry, out int docNum, out string errorText, string transactionType = null)
        {
            errorText = null;
            docEntry = 0;
            docNum = 0;
            SAPbobsCOM.Payments oPayments = null;
            oPayments = (SAPbobsCOM.Payments)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments);
            SAPbobsCOM.Payments oPaymentsNew = null;
            oPaymentsNew = (SAPbobsCOM.Payments)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments);
            SAPbobsCOM.SBObob oSBOBob = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);

            try
            {
                string localCurrency = CommonFunctions.getLocalCurrency();

                DateTime docDate = oDataTable.GetValue("DocumentDate", i);
                DateTime valueDate = oDataTable.GetValue("ValueDate", i);
                string GLAccountCode = oDataTable.GetValue("GLAccountCode", i);
                string projectCod = oDataTable.GetValue("Project", i);
                if (string.IsNullOrEmpty(GLAccountCode))
                    errorText = BDOSResources.getTranslate("TheFollowingFieldIsMandatory") + " : \"" + BDOSResources.getTranslate("GLAccountCode") + "\"! ";
                string cashFlowLineItemName = oDataTable.GetValue("CashFlowLineItemName", i);
                string accountNumber = oDataTable.GetValue("AccountNumber", i);
                string currency = oDataTable.GetValue("Currency", i);
                string currencySapCode = CommonFunctions.getCurrencySapCode(currency);
                if (string.IsNullOrEmpty(currencySapCode))
                    errorText = errorText + BDOSResources.getTranslate("CouldNotFindCurrency") + " \"" + currency + "\"! ";
                string partnerAccountNumber = transactionType == OperationTypeFromIntBank.TreasuryTransferPaymentOrderIoBP.ToString() ? oDataTable.GetValue("TreasuryCode", i) : oDataTable.GetValue("PartnerAccountNumber", i);
                string partnerCurrency = transactionType == OperationTypeFromIntBank.TreasuryTransferPaymentOrderIoBP.ToString() ? "GEL" : oDataTable.GetValue("PartnerCurrency", i);
                string partnerCurrencySapCode = CommonFunctions.getCurrencySapCode(partnerCurrency);
                if (string.IsNullOrEmpty(partnerCurrencySapCode))
                    errorText = errorText + BDOSResources.getTranslate("CouldNotFindCurrency") + " \"" + partnerCurrencySapCode + "\"! ";
                string transferAccount = CommonFunctions.getTransferAccount(accountNumber + currency);
                if (string.IsNullOrEmpty(transferAccount))
                    errorText = errorText + BDOSResources.getTranslate("CheckGLAccountForHouseBankAccount") + " \"" + accountNumber + currency + "\"! ";
                bool cashFlowRelevant = CommonFunctions.isAccountCashFlowRelevant(transferAccount);
                string cashFlowLineItemID = oDataTable.GetValue("CashFlowLineItemID", i);
                if (cashFlowRelevant && string.IsNullOrEmpty(cashFlowLineItemID))
                    errorText = errorText + BDOSResources.getTranslate("TheFollowingFieldIsMandatory") + " : \"" + BDOSResources.getTranslate("CashFlowLineItemID") + "\"! ";
                string partnerTaxCode = oDataTable.GetValue("PartnerTaxCode", i);
                string blnkAgr = oDataTable.GetValue("BlnkAgr", i);
                string useBlaAgRt = oDataTable.GetValue("UseBlaAgRt", i);

                SAPbobsCOM.Recordset oRecordSet = null;
                if (transactionType == OperationTypeFromIntBank.ReturnToCustomer.ToString())
                {
                    oRecordSet = CommonFunctions.getBPBankInfo(partnerAccountNumber + partnerCurrency, partnerTaxCode, "C");
                }
                else
                {
                    oRecordSet = CommonFunctions.getBPBankInfo(partnerAccountNumber + partnerCurrency, partnerTaxCode, "S");
                }


                if (oRecordSet == null)
                {
                    errorText = BDOSResources.getTranslate("CouldNotFindBusinessPartner") + "! " + BDOSResources.getTranslate("Account") + " \"" + partnerAccountNumber + currency + "\"";
                    if (!string.IsNullOrEmpty(partnerTaxCode))
                        errorText = errorText + ", " + BDOSResources.getTranslate("Tin") + " \"" + partnerTaxCode + "\"! ";
                    else errorText = errorText + "! ";
                }

                if (!string.IsNullOrEmpty(errorText))
                {
                    errorText = BDOSResources.getTranslate("OperationCompletedUnSuccessfully") + "! " + errorText + BDOSResources.getTranslate("TableRow") + " : " + (i + 1);
                    return null;
                }

                oPayments.ProjectCode = projectCod;
                oPayments.CardCode = oRecordSet.Fields.Item("CardCode").Value;
                oPayments.CardName = oRecordSet.Fields.Item("CardName").Value;
                string BPCurrency = oRecordSet.Fields.Item("Currency").Value;
                oPayments.PayToBankCountry = oRecordSet.Fields.Item("Country").Value;
                oPayments.PayToBankCode = oRecordSet.Fields.Item("BankCode").Value;
                oPayments.PayToBankAccountNo = oRecordSet.Fields.Item("Account").Value;
                oPayments.IsPayToBank = SAPbobsCOM.BoYesNoEnum.tYES;
                oPayments.ControlAccount = GLAccountCode;

                Marshal.FinalReleaseComObject(oRecordSet);
                oRecordSet = null;

                oPayments.DocObjectCode = SAPbobsCOM.BoPaymentsObjectType.bopot_OutgoingPayments;

                if (transactionType == OperationTypeFromIntBank.ReturnToCustomer.ToString())
                {
                    oPayments.DocTypte = SAPbobsCOM.BoRcptTypes.rCustomer;
                }
                else
                {
                    oPayments.DocTypte = SAPbobsCOM.BoRcptTypes.rSupplier;
                }

                oPayments.DocDate = docDate;
                oPayments.TaxDate = docDate;

                double docRateByBlnktAgr = 0;
                if (!string.IsNullOrEmpty(blnkAgr))
                {
                    oPayments.BlanketAgreement = Convert.ToInt32(blnkAgr);
                    oPayments.UserFields.Fields.Item("U_UseBlaAgRt").Value = useBlaAgRt;
                    string docCur;
                    if (useBlaAgRt == "Y")
                    {
                        docRateByBlnktAgr = Convert.ToDouble(BlanketAgreement.GetBlAgremeentCurrencyRate(Convert.ToInt32(blnkAgr), out docCur, docDate), NumberFormatInfo.InvariantInfo);
                    }
                }

                decimal docRate;
                decimal transferSumLC = 0;
                decimal transferSumFC = 0;
                decimal amount = Convert.ToDecimal(oDataTable.GetValue("Amount", i), NumberFormatInfo.InvariantInfo);

                if (currencySapCode == partnerCurrencySapCode)
                {
                    if (BPCurrency != "##" && BPCurrency != localCurrency)
                    {
                        if (partnerCurrencySapCode == localCurrency)
                        {
                            oPayments.DocCurrency = BPCurrency;
                            oPayments.LocalCurrency = SAPbobsCOM.BoYesNoEnum.tYES;
                            oPayments.DocRate = useBlaAgRt == "Y" ? docRateByBlnktAgr : oSBOBob.GetCurrencyRate(BPCurrency, docDate).Fields.Item("CurrencyRate").Value;
                            docRate = Convert.ToDecimal(oPayments.DocRate);
                            transferSumLC = amount;
                            transferSumFC = amount / docRate;
                        }
                        else if (partnerCurrencySapCode == BPCurrency)
                        {
                            oPayments.DocCurrency = BPCurrency;
                            oPayments.LocalCurrency = SAPbobsCOM.BoYesNoEnum.tNO;
                            oPayments.DocRate = useBlaAgRt == "Y" ? docRateByBlnktAgr : oSBOBob.GetCurrencyRate(partnerCurrencySapCode, docDate).Fields.Item("CurrencyRate").Value;
                            docRate = docRate = Convert.ToDecimal(oPayments.DocRate);
                            transferSumLC = amount * docRate;
                            transferSumFC = amount;
                        }
                        else
                        {
                            errorText = BDOSResources.getTranslate("OperationCompletedUnSuccessfully") + "! " + BDOSResources.getTranslate("CheckCurrencies") + "! " + errorText + BDOSResources.getTranslate("TableRow") + " : " + (i + 1);
                            return null;
                        }
                    }
                    else if (BPCurrency == localCurrency)
                    {
                        if (partnerCurrencySapCode == localCurrency)
                        {
                            oPayments.DocCurrency = BPCurrency;
                            oPayments.LocalCurrency = SAPbobsCOM.BoYesNoEnum.tNO;
                            oPayments.DocRate = 0;
                            transferSumLC = amount;
                            transferSumFC = 0;
                        }
                        else
                        {
                            errorText = BDOSResources.getTranslate("OperationCompletedUnSuccessfully") + "! " + BDOSResources.getTranslate("CheckCurrencies") + "! " + errorText + BDOSResources.getTranslate("TableRow") + " : " + (i + 1);
                            return null;
                        }
                    }
                    else if (BPCurrency == "##")
                    {
                        if (partnerCurrencySapCode == localCurrency)
                        {
                            oPayments.DocCurrency = partnerCurrencySapCode;
                            oPayments.LocalCurrency = SAPbobsCOM.BoYesNoEnum.tNO;
                            oPayments.DocRate = 0;
                            transferSumLC = amount;
                            transferSumFC = 0;
                        }
                        else if (partnerCurrencySapCode != localCurrency)
                        {
                            oPayments.DocCurrency = partnerCurrencySapCode;
                            oPayments.LocalCurrency = SAPbobsCOM.BoYesNoEnum.tNO;
                            oPayments.DocRate = useBlaAgRt == "Y" ? docRateByBlnktAgr : oSBOBob.GetCurrencyRate(partnerCurrencySapCode, docDate).Fields.Item("CurrencyRate").Value;
                            docRate = docRate = Convert.ToDecimal(oPayments.DocRate);
                            transferSumLC = amount * docRate;
                            transferSumFC = amount;
                        }
                        else
                        {
                            errorText = BDOSResources.getTranslate("OperationCompletedUnSuccessfully") + "! " + BDOSResources.getTranslate("CheckCurrencies") + "! " + errorText + BDOSResources.getTranslate("TableRow") + " : " + (i + 1);
                            return null;
                        }
                    }
                }
                else
                {
                    errorText = BDOSResources.getTranslate("OperationCompletedUnSuccessfully") + "! " + BDOSResources.getTranslate("CheckCurrencies") + "! " + errorText + BDOSResources.getTranslate("TableRow") + " : " + (i + 1);
                    return null;
                }

                oPayments.TransferAccount = transferAccount;
                oPayments.TransferDate = docDate;
                oPayments.TransferSum = Convert.ToDouble(amount, NumberFormatInfo.InvariantInfo);

                if (CommonFunctions.IsDevelopment())
                {
                    oPayments.UserFields.Fields.Item("U_BDOSBdgCf").Value = oDataTable.GetValue("BudgetCashFlowID", i);
                    oPayments.UserFields.Fields.Item("U_BDOSBdgCfN").Value = oDataTable.GetValue("BudgetCashFlowName", i);
                }

                oPayments.UserFields.Fields.Item("U_status").Value = "downloadedFromTheBank";
                oPayments.UserFields.Fields.Item("U_paymentID").Value = oDataTable.GetValue("PaymentID", i);
                oPayments.UserFields.Fields.Item("U_chrgDtls").Value = oDataTable.GetValue("ChargeDetail", i);
                oPayments.UserFields.Fields.Item("U_descrpt").Value = oDataTable.GetValue("Description", i);
                oPayments.UserFields.Fields.Item("U_addDescrpt").Value = oDataTable.GetValue("AdditionalDescription", i);

                oPayments.UserFields.Fields.Item("U_docNumber").Value = oDataTable.GetValue("DocumentNumber", i);
                oPayments.UserFields.Fields.Item("U_transCode").Value = oDataTable.GetValue("TransactionCode", i);
                oPayments.UserFields.Fields.Item("U_ePaymentID").Value = oDataTable.GetValue("ExternalPaymentID", i);
                oPayments.UserFields.Fields.Item("U_opCode").Value = oDataTable.GetValue("OperationCode", i);

                if (cashFlowRelevant)
                {
                    oPayments.PrimaryFormItems.CashFlowLineItemID = Convert.ToInt32(cashFlowLineItemID);
                    oPayments.PrimaryFormItems.AmountFC = Convert.ToDouble(transferSumFC, NumberFormatInfo.InvariantInfo);
                    if (oPayments.DocCurrency == localCurrency)
                        oPayments.PrimaryFormItems.AmountLC = Convert.ToDouble(transferSumLC, NumberFormatInfo.InvariantInfo);
                    oPayments.PrimaryFormItems.PaymentMeans = SAPbobsCOM.PaymentMeansTypeEnum.pmtBankTransfer;
                    oPayments.PrimaryFormItems.Add();
                }

                int returnCode = oPayments.Add();
                if (returnCode != 0)
                {
                    int errCode;
                    string errMsg;
                    Program.oCompany.GetLastError(out errCode, out errMsg);
                    errorText = BDOSResources.getTranslate("OperationCompletedUnSuccessfully") + "! " + errMsg + "! " + BDOSResources.getTranslate("TableRow") + " : " + (i + 1);
                    return null;
                }
                else
                {
                    bool newDoc = oPaymentsNew.GetByKey(Convert.ToInt32(Program.oCompany.GetNewObjectKey()));
                    if (newDoc)
                    {
                        docEntry = oPaymentsNew.DocEntry;
                        docNum = oPaymentsNew.DocNum;
                        oDataTable.SetValue("DocEntry", i, docEntry.ToString());
                        oDataTable.SetValue("DocNum", i, docNum.ToString());
                        return BDOSResources.getTranslate("OperationCompletedSuccessfully") + "! " + BDOSResources.getTranslate("TableRow") + " : " + (i + 1);
                    }
                    else
                        return "";
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
                return null;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oPayments);
                oPayments = null;
                Marshal.FinalReleaseComObject(oPaymentsNew);
                oPaymentsNew = null;
                Marshal.FinalReleaseComObject(oSBOBob);
                oSBOBob = null;
            }
        }

        public static string createDocumentTreasuryTransferType(SAPbouiCOM.DataTable oDataTable, int i, out int docEntry, out int docNum, out string errorText)
        {
            errorText = null;
            docEntry = 0;
            docNum = 0;
            SAPbobsCOM.Payments oPayments = null;
            oPayments = (SAPbobsCOM.Payments)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments);
            SAPbobsCOM.Payments oPaymentsNew = null;
            oPaymentsNew = (SAPbobsCOM.Payments)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments);
            SAPbobsCOM.SBObob oSBOBob = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);

            try
            {
                string localCurrency = CommonFunctions.getLocalCurrency();

                DateTime docDate = oDataTable.GetValue("DocumentDate", i);
                DateTime valueDate = oDataTable.GetValue("ValueDate", i);
                string GLAccountCode = oDataTable.GetValue("GLAccountCode", i);
                string projectCod = oDataTable.GetValue("Project", i);

                if (string.IsNullOrEmpty(GLAccountCode))
                    errorText = BDOSResources.getTranslate("TheFollowingFieldIsMandatory") + " : \"" + BDOSResources.getTranslate("GLAccountCode") + "\"! ";
                string cashFlowLineItemName = oDataTable.GetValue("CashFlowLineItemName", i);
                string accountNumber = oDataTable.GetValue("AccountNumber", i);
                string currency = oDataTable.GetValue("Currency", i);
                string currencySapCode = CommonFunctions.getCurrencySapCode(currency);
                if (string.IsNullOrEmpty(currencySapCode))
                    errorText = errorText + BDOSResources.getTranslate("CouldNotFindCurrency") + " \"" + currency + "\"! ";
                string treasuryCode = oDataTable.GetValue("TreasuryCode", i);
                if (string.IsNullOrEmpty(treasuryCode))
                    errorText = errorText + BDOSResources.getTranslate("TheFollowingFieldIsMandatory") + " : \"" + BDOSResources.getTranslate("TreasuryCode") + "\"! ";
                string transferAccount = CommonFunctions.getTransferAccount(accountNumber + currency);
                if (string.IsNullOrEmpty(transferAccount))
                    errorText = errorText + BDOSResources.getTranslate("CheckGLAccountForHouseBankAccount") + " \"" + accountNumber + currency + "\"! ";
                bool cashFlowRelevant = CommonFunctions.isAccountCashFlowRelevant(transferAccount);
                string cashFlowLineItemID = oDataTable.GetValue("CashFlowLineItemID", i);
                if (cashFlowRelevant && string.IsNullOrEmpty(cashFlowLineItemID))
                    errorText = errorText + BDOSResources.getTranslate("TheFollowingFieldIsMandatory") + " : \"" + BDOSResources.getTranslate("CashFlowLineItemID") + "\"! ";

                if (string.IsNullOrEmpty(errorText) == false)
                {
                    errorText = BDOSResources.getTranslate("OperationCompletedUnSuccessfully") + "! " + errorText + BDOSResources.getTranslate("TableRow") + " : " + (i + 1);
                    return null;
                }

                oPayments.ProjectCode = projectCod;
                oPayments.DocObjectCode = SAPbobsCOM.BoPaymentsObjectType.bopot_OutgoingPayments;
                oPayments.DocTypte = SAPbobsCOM.BoRcptTypes.rAccount;
                oPayments.DocDate = docDate;
                oPayments.TaxDate = docDate;
                oPayments.DocCurrency = currencySapCode;
                oPayments.LocalCurrency = SAPbobsCOM.BoYesNoEnum.tNO;

                decimal transferSumLC;
                decimal transferSumFC;
                decimal grossAmount;
                decimal amount = Convert.ToDecimal(oDataTable.GetValue("Amount", i), NumberFormatInfo.InvariantInfo);

                if (currencySapCode == localCurrency)
                {
                    oPayments.DocCurrency = currencySapCode;
                    oPayments.LocalCurrency = SAPbobsCOM.BoYesNoEnum.tNO;
                    oPayments.DocRate = 0;
                    transferSumLC = amount;
                    transferSumFC = 0;
                    grossAmount = amount;
                }
                else
                {
                    errorText = BDOSResources.getTranslate("OperationCompletedUnSuccessfully") + "! " + BDOSResources.getTranslate("CheckCurrencies") + "! " + errorText + BDOSResources.getTranslate("TableRow") + " : " + (i + 1);
                    return null;
                }

                transferSumLC = CommonFunctions.roundAmountByGeneralSettings(transferSumLC, "Sum");
                transferSumFC = CommonFunctions.roundAmountByGeneralSettings(transferSumFC, "Sum");
                grossAmount = CommonFunctions.roundAmountByGeneralSettings(grossAmount, "Sum");
                amount = CommonFunctions.roundAmountByGeneralSettings(amount, "Sum");

                oPayments.TransferAccount = transferAccount;
                oPayments.TransferDate = docDate;
                oPayments.TransferSum = Convert.ToDouble(amount, NumberFormatInfo.InvariantInfo);

                if (CommonFunctions.IsDevelopment())
                {
                    oPayments.UserFields.Fields.Item("U_BDOSBdgCf").Value = oDataTable.GetValue("BudgetCashFlowID", i);
                    oPayments.UserFields.Fields.Item("U_BDOSBdgCfN").Value = oDataTable.GetValue("BudgetCashFlowName", i);
                }

                oPayments.UserFields.Fields.Item("U_opType").Value = "treasuryTransfer";
                oPayments.UserFields.Fields.Item("U_status").Value = "downloadedFromTheBank";
                oPayments.UserFields.Fields.Item("U_paymentID").Value = oDataTable.GetValue("PaymentID", i);
                oPayments.UserFields.Fields.Item("U_tresrCode").Value = treasuryCode;
                oPayments.UserFields.Fields.Item("U_descrpt").Value = oDataTable.GetValue("Description", i);
                oPayments.UserFields.Fields.Item("U_addDescrpt").Value = oDataTable.GetValue("AdditionalDescription", i);

                oPayments.UserFields.Fields.Item("U_docNumber").Value = oDataTable.GetValue("DocumentNumber", i);
                oPayments.UserFields.Fields.Item("U_transCode").Value = oDataTable.GetValue("TransactionCode", i);
                oPayments.UserFields.Fields.Item("U_ePaymentID").Value = oDataTable.GetValue("ExternalPaymentID", i);
                oPayments.UserFields.Fields.Item("U_opCode").Value = oDataTable.GetValue("OperationCode", i);

                //ცხრილური ნაწილი
                oPayments.AccountPayments.AccountCode = GLAccountCode;
                oPayments.AccountPayments.ProjectCode = projectCod;
                oPayments.AccountPayments.GrossAmount = Convert.ToDouble(grossAmount, NumberFormatInfo.InvariantInfo);
                oPayments.AccountPayments.Add();

                if (cashFlowRelevant)
                {
                    oPayments.PrimaryFormItems.CashFlowLineItemID = Convert.ToInt32(cashFlowLineItemID);
                    oPayments.PrimaryFormItems.AmountFC = Convert.ToDouble(transferSumFC, NumberFormatInfo.InvariantInfo);
                    if (oPayments.DocCurrency == localCurrency)
                        oPayments.PrimaryFormItems.AmountLC = Convert.ToDouble(transferSumLC, NumberFormatInfo.InvariantInfo);
                    oPayments.PrimaryFormItems.PaymentMeans = SAPbobsCOM.PaymentMeansTypeEnum.pmtBankTransfer;
                    oPayments.PrimaryFormItems.Add();
                }

                int returnCode = oPayments.Add();
                if (returnCode != 0)
                {
                    int errCode;
                    string errMsg;
                    Program.oCompany.GetLastError(out errCode, out errMsg);
                    errorText = BDOSResources.getTranslate("OperationCompletedUnSuccessfully") + "! " + errMsg + "! " + BDOSResources.getTranslate("TableRow") + " : " + (i + 1);
                    return null;
                }
                else
                {
                    bool newDoc = oPaymentsNew.GetByKey(Convert.ToInt32(Program.oCompany.GetNewObjectKey()));
                    if (newDoc)
                    {
                        docEntry = oPaymentsNew.DocEntry;
                        docNum = oPaymentsNew.DocNum;
                        oDataTable.SetValue("DocEntry", i, docEntry.ToString());
                        oDataTable.SetValue("DocNum", i, docNum.ToString());
                        return BDOSResources.getTranslate("OperationCompletedSuccessfully") + "! " + BDOSResources.getTranslate("TableRow") + " : " + (i + 1);
                    }
                    else
                        return "";
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
                return null;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oPayments);
                oPayments = null;
                Marshal.FinalReleaseComObject(oPaymentsNew);
                oPaymentsNew = null;
                Marshal.FinalReleaseComObject(oSBOBob);
                oSBOBob = null;
            }
        }

        public static string createDocumentOtherExpensesType(SAPbouiCOM.DataTable oDataTable, int i, out int docEntry, out int docNum, out string errorText)
        {
            errorText = null;
            docEntry = 0;
            docNum = 0;
            SAPbobsCOM.Payments oPayments = null;
            oPayments = (SAPbobsCOM.Payments)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments);
            SAPbobsCOM.Payments oPaymentsNew = null;
            oPaymentsNew = (SAPbobsCOM.Payments)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments);
            SAPbobsCOM.SBObob oSBOBob = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);
            try
            {
                string localCurrency = CommonFunctions.getLocalCurrency();

                DateTime docDate = oDataTable.GetValue("DocumentDate", i);
                DateTime valueDate = oDataTable.GetValue("ValueDate", i);
                string GLAccountCode = oDataTable.GetValue("GLAccountCode", i);
                string projectCod = oDataTable.GetValue("Project", i);
                if (string.IsNullOrEmpty(GLAccountCode))
                    errorText = BDOSResources.getTranslate("TheFollowingFieldIsMandatory") + " : \"" + BDOSResources.getTranslate("GLAccountCode") + "\"! ";
                string cashFlowLineItemName = oDataTable.GetValue("CashFlowLineItemName", i);
                string accountNumber = oDataTable.GetValue("AccountNumber", i);
                string currency = oDataTable.GetValue("Currency", i);
                string currencySapCode = CommonFunctions.getCurrencySapCode(currency);
                if (string.IsNullOrEmpty(currencySapCode))
                    errorText = errorText + BDOSResources.getTranslate("CouldNotFindCurrency") + " \"" + currency + "\"! ";
                string partnerCurrency = oDataTable.GetValue("PartnerCurrency", i);
                string partnerCurrencySapCode = CommonFunctions.getCurrencySapCode(partnerCurrency);
                if (string.IsNullOrEmpty(partnerCurrencySapCode))
                    partnerCurrencySapCode = localCurrency;
                string transferAccount = CommonFunctions.getTransferAccount(accountNumber + currency);
                if (string.IsNullOrEmpty(transferAccount))
                    errorText = errorText + BDOSResources.getTranslate("CheckGLAccountForHouseBankAccount") + " \"" + accountNumber + currency + "\"! ";
                bool cashFlowRelevant = CommonFunctions.isAccountCashFlowRelevant(transferAccount);
                string cashFlowLineItemID = oDataTable.GetValue("CashFlowLineItemID", i);
                if (cashFlowRelevant && string.IsNullOrEmpty(cashFlowLineItemID))
                    errorText = errorText + BDOSResources.getTranslate("TheFollowingFieldIsMandatory") + " : \"" + BDOSResources.getTranslate("CashFlowLineItemID") + "\"! ";

                if (string.IsNullOrEmpty(errorText) == false)
                {
                    errorText = BDOSResources.getTranslate("OperationCompletedUnSuccessfully") + "! " + errorText + BDOSResources.getTranslate("TableRow") + " : " + (i + 1);
                    return null;
                }

                oPayments.ProjectCode = projectCod;
                oPayments.DocObjectCode = SAPbobsCOM.BoPaymentsObjectType.bopot_OutgoingPayments;
                oPayments.DocTypte = SAPbobsCOM.BoRcptTypes.rAccount;
                oPayments.DocDate = docDate;
                oPayments.TaxDate = docDate;

                decimal docRate;
                decimal transferSumLC;
                decimal transferSumFC;
                decimal grossAmount;
                decimal amount = Convert.ToDecimal(oDataTable.GetValue("Amount", i), NumberFormatInfo.InvariantInfo);

                if ((currencySapCode == partnerCurrencySapCode) && currencySapCode == localCurrency)
                {
                    oPayments.DocCurrency = currencySapCode;
                    oPayments.LocalCurrency = SAPbobsCOM.BoYesNoEnum.tNO;
                    oPayments.DocRate = 0;
                    transferSumLC = amount;
                    transferSumFC = 0;
                    grossAmount = amount;
                }
                else if ((currencySapCode == partnerCurrencySapCode) && currencySapCode != localCurrency)
                {
                    oPayments.DocCurrency = currencySapCode;
                    oPayments.LocalCurrency = SAPbobsCOM.BoYesNoEnum.tNO;
                    oPayments.DocRate = oSBOBob.GetCurrencyRate(currencySapCode, docDate).Fields.Item("CurrencyRate").Value;
                    docRate = Convert.ToDecimal(oPayments.DocRate);
                    transferSumLC = amount * docRate;
                    transferSumFC = amount;
                    grossAmount = amount;
                }
                else if ((currencySapCode != partnerCurrencySapCode) && currencySapCode == localCurrency)
                {
                    oPayments.DocCurrency = currencySapCode;
                    oPayments.LocalCurrency = SAPbobsCOM.BoYesNoEnum.tNO;
                    oPayments.DocRate = 0;
                    transferSumLC = amount;
                    transferSumFC = 0;
                    grossAmount = amount;
                }
                else if ((currencySapCode != partnerCurrencySapCode) && currencySapCode != localCurrency)
                {
                    oPayments.DocCurrency = currencySapCode;
                    oPayments.LocalCurrency = SAPbobsCOM.BoYesNoEnum.tNO;
                    oPayments.DocRate = oSBOBob.GetCurrencyRate(currencySapCode, docDate).Fields.Item("CurrencyRate").Value;
                    docRate = Convert.ToDecimal(oPayments.DocRate);
                    transferSumLC = amount * docRate;
                    transferSumFC = amount;
                    grossAmount = amount;
                }
                else
                {
                    errorText = BDOSResources.getTranslate("OperationCompletedUnSuccessfully") + "! " + BDOSResources.getTranslate("CheckCurrencies") + "! " + errorText + BDOSResources.getTranslate("TableRow") + " : " + (i + 1);
                    return null;
                }

                transferSumLC = CommonFunctions.roundAmountByGeneralSettings(transferSumLC, "Sum");
                transferSumFC = CommonFunctions.roundAmountByGeneralSettings(transferSumFC, "Sum");
                grossAmount = CommonFunctions.roundAmountByGeneralSettings(grossAmount, "Sum");
                amount = CommonFunctions.roundAmountByGeneralSettings(amount, "Sum");

                oPayments.TransferAccount = transferAccount;
                oPayments.TransferDate = docDate;
                oPayments.TransferSum = Convert.ToDouble(amount, NumberFormatInfo.InvariantInfo);

                if (CommonFunctions.IsDevelopment())
                {
                    oPayments.UserFields.Fields.Item("U_BDOSBdgCf").Value = oDataTable.GetValue("BudgetCashFlowID", i);
                    oPayments.UserFields.Fields.Item("U_BDOSBdgCfN").Value = oDataTable.GetValue("BudgetCashFlowName", i);
                }

                oPayments.UserFields.Fields.Item("U_opType").Value = "other";
                oPayments.UserFields.Fields.Item("U_status").Value = "downloadedFromTheBank";
                oPayments.UserFields.Fields.Item("U_paymentID").Value = oDataTable.GetValue("PaymentID", i);
                oPayments.UserFields.Fields.Item("U_descrpt").Value = oDataTable.GetValue("Description", i);
                oPayments.UserFields.Fields.Item("U_addDescrpt").Value = oDataTable.GetValue("AdditionalDescription", i);

                oPayments.UserFields.Fields.Item("U_docNumber").Value = oDataTable.GetValue("DocumentNumber", i);
                oPayments.UserFields.Fields.Item("U_transCode").Value = oDataTable.GetValue("TransactionCode", i);
                oPayments.UserFields.Fields.Item("U_ePaymentID").Value = oDataTable.GetValue("ExternalPaymentID", i);
                oPayments.UserFields.Fields.Item("U_opCode").Value = oDataTable.GetValue("OperationCode", i);

                //ცხრილური ნაწილი
                oPayments.AccountPayments.AccountCode = GLAccountCode;
                oPayments.AccountPayments.ProjectCode = projectCod;
                oPayments.AccountPayments.GrossAmount = Convert.ToDouble(grossAmount, NumberFormatInfo.InvariantInfo);
                oPayments.AccountPayments.Add();

                if (cashFlowRelevant)
                {
                    oPayments.PrimaryFormItems.CashFlowLineItemID = Convert.ToInt32(cashFlowLineItemID);
                    oPayments.PrimaryFormItems.AmountFC = Convert.ToDouble(transferSumFC, NumberFormatInfo.InvariantInfo);
                    if (oPayments.DocCurrency == localCurrency)
                        oPayments.PrimaryFormItems.AmountLC = Convert.ToDouble(transferSumLC, NumberFormatInfo.InvariantInfo);
                    oPayments.PrimaryFormItems.PaymentMeans = SAPbobsCOM.PaymentMeansTypeEnum.pmtBankTransfer;
                    oPayments.PrimaryFormItems.Add();
                }

                int returnCode = oPayments.Add();
                if (returnCode != 0)
                {
                    int errCode;
                    string errMsg;
                    Program.oCompany.GetLastError(out errCode, out errMsg);
                    errorText = BDOSResources.getTranslate("OperationCompletedUnSuccessfully") + "! " + errMsg + "! " + BDOSResources.getTranslate("TableRow") + " : " + (i + 1);
                    return null;
                }
                else
                {
                    bool newDoc = oPaymentsNew.GetByKey(Convert.ToInt32(Program.oCompany.GetNewObjectKey()));
                    if (newDoc)
                    {
                        docEntry = oPaymentsNew.DocEntry;
                        docNum = oPaymentsNew.DocNum;
                        oDataTable.SetValue("DocEntry", i, docEntry.ToString());
                        oDataTable.SetValue("DocNum", i, docNum.ToString());
                        return BDOSResources.getTranslate("OperationCompletedSuccessfully") + "! " + BDOSResources.getTranslate("TableRow") + " : " + (i + 1);
                    }
                    else
                        return "";
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
                return null;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oPayments);
                oPayments = null;
                Marshal.FinalReleaseComObject(oPaymentsNew);
                oPaymentsNew = null;
                Marshal.FinalReleaseComObject(oSBOBob);
                oSBOBob = null;
            }
        }

        public static string createDocumentCurrencyExchangeType(SAPbouiCOM.DataTable oDataTable, int i, out int docEntry, out int docNum, out string errorText)
        {
            errorText = null;
            docEntry = 0;
            docNum = 0;
            SAPbobsCOM.Payments oPayments = null;
            oPayments = (SAPbobsCOM.Payments)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments);
            SAPbobsCOM.Payments oPaymentsNew = null;
            oPaymentsNew = (SAPbobsCOM.Payments)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments);
            SAPbobsCOM.SBObob oSBOBob = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);

            try
            {
                string localCurrency = CommonFunctions.getLocalCurrency();
                string localCurrencyInternationalCode = CommonFunctions.getCurrencyInternationalCode(localCurrency);

                DateTime docDate = oDataTable.GetValue("DocumentDate", i);
                DateTime valueDate = oDataTable.GetValue("ValueDate", i);
                string GLAccountCode = oDataTable.GetValue("GLAccountCode", i);
                string projectCod = oDataTable.GetValue("Project", i);

                if (string.IsNullOrEmpty(GLAccountCode))
                    errorText = BDOSResources.getTranslate("TheFollowingFieldIsMandatory") + " : \"" + BDOSResources.getTranslate("GLAccountCode") + "\"! ";
                string cashFlowLineItemName = oDataTable.GetValue("CashFlowLineItemName", i);
                string accountNumber = oDataTable.GetValue("AccountNumber", i);
                string currency = oDataTable.GetValue("Currency", i);
                string currencySapCode = CommonFunctions.getCurrencySapCode(currency);
                if (string.IsNullOrEmpty(currencySapCode))
                    errorText = errorText + BDOSResources.getTranslate("CouldNotFindCurrency") + " \"" + currency + "\"! ";
                string partnerAccountNumber = oDataTable.GetValue("PartnerAccountNumber", i);
                string partnerCurrency = oDataTable.GetValue("PartnerCurrency", i);
                string partnerCurrencySapCode = CommonFunctions.getCurrencySapCode(partnerCurrency);
                string currencyExchange = oDataTable.GetValue("CurrencyExchange", i);
                string currencyExchangeSapCode = CommonFunctions.getCurrencySapCode(currencyExchange);

                //string bankCode = CommonFunctions.getBankCode( null, accountNumber + currency);
                //if (currencySapCode != localCurrency)
                //{
                //    partnerCurrency = localCurrencyInternationalCode;
                //    currencyExchange = localCurrencyInternationalCode;
                //    partnerCurrencySapCode = localCurrency;
                //    currencyExchangeSapCode = localCurrency;
                //    partnerAccountNumber = CommonFunctions.getHouseBankAccount( bankCode, partnerCurrency);
                //    if (string.IsNullOrEmpty(partnerAccountNumber))
                //        errorText = errorText + BDOSResources.getTranslate("CouldNotFindHouseBankAccount") + " " + BDOSResources.getTranslate("Currency") + " : \"" + partnerCurrency + "\"! ";
                //    else
                //        partnerAccountNumber = CommonFunctions.accountParse(partnerAccountNumber);
                //}
                //else
                //{
                if (string.IsNullOrEmpty(partnerCurrencySapCode))
                    errorText = errorText + BDOSResources.getTranslate("CouldNotFindCurrency") + " \"" + partnerCurrencySapCode + "\"! ";
                if (string.IsNullOrEmpty(currencyExchangeSapCode))
                    errorText = errorText + BDOSResources.getTranslate("CouldNotFindCurrency") + " \"" + currencyExchangeSapCode + "\"! ";
                if (CommonFunctions.isAccountInHouseBankAccount(partnerAccountNumber + partnerCurrency) == false)
                    errorText = errorText + BDOSResources.getTranslate("CouldNotFindHouseBankAccount") + " \"" + partnerAccountNumber + partnerCurrency + "\"! ";
                //}

                string transferAccount = CommonFunctions.getTransferAccount(accountNumber + currency);
                if (string.IsNullOrEmpty(transferAccount))
                    errorText = errorText + BDOSResources.getTranslate("CheckGLAccountForHouseBankAccount") + " \"" + accountNumber + currency + "\"! ";
                bool cashFlowRelevant = CommonFunctions.isAccountCashFlowRelevant(transferAccount);
                string cashFlowLineItemID = oDataTable.GetValue("CashFlowLineItemID", i);
                if (cashFlowRelevant && string.IsNullOrEmpty(cashFlowLineItemID))
                    errorText = errorText + BDOSResources.getTranslate("TheFollowingFieldIsMandatory") + " : \"" + BDOSResources.getTranslate("CashFlowLineItemID") + "\"! ";

                if (string.IsNullOrEmpty(errorText) == false)
                {
                    errorText = BDOSResources.getTranslate("OperationCompletedUnSuccessfully") + "! " + errorText + BDOSResources.getTranslate("TableRow") + " : " + (i + 1);
                    return null;
                }

                oPayments.ProjectCode = projectCod;
                oPayments.DocObjectCode = SAPbobsCOM.BoPaymentsObjectType.bopot_OutgoingPayments;
                oPayments.DocTypte = SAPbobsCOM.BoRcptTypes.rAccount;
                oPayments.DocDate = docDate;
                oPayments.TaxDate = docDate;

                decimal docRate;
                decimal transferSumLC;
                decimal transferSumFC;
                decimal grossAmount;
                decimal amount = Convert.ToDecimal(oDataTable.GetValue("Amount", i), NumberFormatInfo.InvariantInfo);

                if ((currencySapCode != partnerCurrencySapCode) && currencySapCode == localCurrency)
                {
                    oPayments.DocCurrency = currencySapCode;
                    oPayments.LocalCurrency = SAPbobsCOM.BoYesNoEnum.tNO;
                    oPayments.DocRate = 0;
                    transferSumLC = amount;
                    transferSumFC = 0;
                    grossAmount = amount;
                }
                else if ((currencySapCode != partnerCurrencySapCode) && currencySapCode != localCurrency)
                {
                    oPayments.DocCurrency = currencySapCode;
                    oPayments.LocalCurrency = SAPbobsCOM.BoYesNoEnum.tNO;
                    oPayments.DocRate = oSBOBob.GetCurrencyRate(currencySapCode, docDate).Fields.Item("CurrencyRate").Value;
                    docRate = Convert.ToDecimal(oPayments.DocRate);
                    transferSumLC = amount * docRate;
                    transferSumFC = amount;
                    grossAmount = amount;
                }
                else
                {
                    errorText = BDOSResources.getTranslate("OperationCompletedUnSuccessfully") + "! " + BDOSResources.getTranslate("CheckCurrencies") + "! " + errorText + BDOSResources.getTranslate("TableRow") + " : " + (i + 1);
                    return null;
                }

                transferSumLC = CommonFunctions.roundAmountByGeneralSettings(transferSumLC, "Sum");
                transferSumFC = CommonFunctions.roundAmountByGeneralSettings(transferSumFC, "Sum");
                grossAmount = CommonFunctions.roundAmountByGeneralSettings(grossAmount, "Sum");
                amount = CommonFunctions.roundAmountByGeneralSettings(amount, "Sum");

                oPayments.TransferAccount = transferAccount;
                oPayments.TransferDate = docDate;
                oPayments.TransferSum = Convert.ToDouble(amount, NumberFormatInfo.InvariantInfo);

                if (CommonFunctions.IsDevelopment())
                {
                    oPayments.UserFields.Fields.Item("U_BDOSBdgCf").Value = oDataTable.GetValue("BudgetCashFlowID", i);
                    oPayments.UserFields.Fields.Item("U_BDOSBdgCfN").Value = oDataTable.GetValue("BudgetCashFlowName", i);
                }

                oPayments.UserFields.Fields.Item("U_opType").Value = "currencyExchange";
                oPayments.UserFields.Fields.Item("U_status").Value = "downloadedFromTheBank";
                oPayments.UserFields.Fields.Item("U_paymentID").Value = oDataTable.GetValue("PaymentID", i);
                oPayments.UserFields.Fields.Item("U_creditAcct").Value = partnerAccountNumber + partnerCurrency;
                oPayments.UserFields.Fields.Item("U_crdtActCur").Value = partnerCurrencySapCode;
                oPayments.UserFields.Fields.Item("U_descrpt").Value = oDataTable.GetValue("Description", i);
                oPayments.UserFields.Fields.Item("U_addDescrpt").Value = oDataTable.GetValue("AdditionalDescription", i);

                oPayments.UserFields.Fields.Item("U_docNumber").Value = oDataTable.GetValue("DocumentNumber", i);
                oPayments.UserFields.Fields.Item("U_transCode").Value = oDataTable.GetValue("TransactionCode", i);
                oPayments.UserFields.Fields.Item("U_ePaymentID").Value = oDataTable.GetValue("ExternalPaymentID", i);
                oPayments.UserFields.Fields.Item("U_opCode").Value = oDataTable.GetValue("OperationCode", i);

                //ცხრილური ნაწილი
                oPayments.AccountPayments.AccountCode = GLAccountCode;
                oPayments.AccountPayments.ProjectCode = projectCod;
                oPayments.AccountPayments.GrossAmount = Convert.ToDouble(grossAmount, NumberFormatInfo.InvariantInfo);
                oPayments.AccountPayments.Add();

                if (cashFlowRelevant)
                {
                    oPayments.PrimaryFormItems.CashFlowLineItemID = Convert.ToInt32(cashFlowLineItemID);
                    oPayments.PrimaryFormItems.AmountFC = Convert.ToDouble(transferSumFC, NumberFormatInfo.InvariantInfo);
                    if (oPayments.DocCurrency == localCurrency)
                        oPayments.PrimaryFormItems.AmountLC = Convert.ToDouble(transferSumLC, NumberFormatInfo.InvariantInfo);
                    oPayments.PrimaryFormItems.PaymentMeans = SAPbobsCOM.PaymentMeansTypeEnum.pmtBankTransfer;
                    oPayments.PrimaryFormItems.Add();
                }

                int returnCode = oPayments.Add();
                if (returnCode != 0)
                {
                    int errCode;
                    string errMsg;
                    Program.oCompany.GetLastError(out errCode, out errMsg);
                    errorText = BDOSResources.getTranslate("OperationCompletedUnSuccessfully") + "! " + errMsg + "! " + BDOSResources.getTranslate("TableRow") + " : " + (i + 1);
                    return null;
                }
                else
                {
                    bool newDoc = oPaymentsNew.GetByKey(Convert.ToInt32(Program.oCompany.GetNewObjectKey()));
                    if (newDoc)
                    {
                        docEntry = oPaymentsNew.DocEntry;
                        docNum = oPaymentsNew.DocNum;
                        oDataTable.SetValue("DocEntry", i, docEntry.ToString());
                        oDataTable.SetValue("DocNum", i, docNum.ToString());

                        //incoming - ის დოკუმენტის მოძებნა და მასზე outgoing - ის მიბმა --->
                        IncomingPayment.attachOutgoingPayments(oDataTable.GetValue("PaymentID", i), oDataTable.GetValue("DocumentNumber", i), oDataTable.GetValue("ExternalPaymentID", i), docEntry.ToString(), "currencyExchange");
                        //incoming - ის დოკუმენტის მოძებნა და მასზე outgoing - ის მიბმა <---

                        return BDOSResources.getTranslate("OperationCompletedSuccessfully") + "! " + BDOSResources.getTranslate("TableRow") + " : " + (i + 1);
                    }
                    else
                        return "";
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
                return null;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oPayments);
                oPayments = null;
                Marshal.FinalReleaseComObject(oPaymentsNew);
                oPaymentsNew = null;
                Marshal.FinalReleaseComObject(oSBOBob);
                oSBOBob = null;
            }
        }

        //public static bool isPension(string wtcode, out double rate)
        //{
        //    SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //    rate = 0;

        //    string query = "select \"U_BDOSPhisTx\", \"Rate\" from OWHT " + "\n"
        //    + "where \"WTCode\" = '" + wtcode + "'";

        //    oRecordSet.DoQuery(query);
        //    if (!oRecordSet.EoF)
        //    {
        //        if (oRecordSet.Fields.Item("U_BDOSPhisTx").Value.ToString() == "Y")
        //        {
        //            rate = oRecordSet.Fields.Item("Rate").Value;
        //            return true;
        //        }
        //    }

        //    return false;
        //}

        //public static void fillWtax(SAPbouiCOM.Form oForm, bool forAdditionalJE, out decimal Wtax, out decimal WtPh, out decimal WtCo)
        //{
        //    double rate = 0;
        //    string wtCode = oForm.Items.Item("110").Specific.Value;

        //    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("20").Specific;
        //    Wtax = 0;
        //    WtPh = 0;
        //    WtCo = 0;

        //    decimal whtAmtAp = 0;
        //    decimal whtPnAp = 0;
        //    decimal whtCoAp = 0;
        //    for (int i = 1; i < oMatrix.RowCount + 1; i++)
        //    {
        //        if (oMatrix.Columns.Item("10000127").Cells.Item(i).Specific.Checked && oMatrix.Columns.Item("45").Cells.Item(i).Specific.Value == "18")
        //        {
        //            string docNum = oMatrix.Columns.Item("1").Cells.Item(i).Specific.Value;
        //            string WTCode = "";
        //            checkWTaxCodeFromMatrix(oMatrix, i, out WTCode, "OPCH", "PCH5", "18");

        //            var totalPayment = Convert.ToDecimal(FormsB1.cleanStringOfNonDigits(oMatrix.Columns.Item("24").Cells.Item(i).Specific.String));
        //            CalculateWTax(WTCode, totalPayment, out decimal wTaxAmt, out decimal pensionAmt);

        //            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //            string query = "select \"U_BDOSWhtAmt\", \"U_BDOSPnPhAm\", \"U_BDOSPnCoAm\" from PCH1 " + "\n"
        //                            + "where \"DocEntry\" = " + "\n"
        //                            + "(select \"DocEntry\" from OPCH " + "\n"
        //                            + "where \"DocNum\" = '" + docNum + "')";

        //            oRecordSet.DoQuery(query);
        //            if (!oRecordSet.EoF)
        //            {
        //                if (!forAdditionalJE)
        //                {
        //                    whtAmtAp += wTaxAmt;
        //                    whtPnAp += pensionAmt;
        //                    whtCoAp += pensionAmt;
        //                }
        //                else if (forAdditionalJE && oRecordSet.Fields.Item("U_BDOSPnPhAm").Value != 0 &&
        //                         !isInvoiceType(WTCode))
        //                {
        //                    whtAmtAp += wTaxAmt;
        //                    whtPnAp += pensionAmt;
        //                    whtCoAp += pensionAmt;
        //                }
        //            }
        //        }
        //        else if(oMatrix.Columns.Item("10000127").Cells.Item(i).Specific.Checked && oMatrix.Columns.Item("45").Cells.Item(i).Specific.Value == "204")
        //        {
        //            string docNum = oMatrix.Columns.Item("1").Cells.Item(i).Specific.Value;
        //            string WTCode = "";
        //            checkWTaxCodeFromMatrix(oMatrix, i, out WTCode, "ODPO", "DPO5", "204");

        //            var totalPayment = Convert.ToDecimal(FormsB1.cleanStringOfNonDigits(oMatrix.Columns.Item("24").Cells.Item(i).Specific.String));
        //            CalculateWTax(WTCode, totalPayment, out decimal wTaxAmt, out decimal pensionAmt);

        //            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //            string query = "select \"U_BDOSWhtAmt\", \"U_BDOSPnPhAm\", \"U_BDOSPnCoAm\" from DPO1 " + "\n"
        //            + "where \"DocEntry\" = " + "\n"
        //            + "(select \"DocEntry\" from ODPO " + "\n"
        //            + "where \"DocNum\" = '" + docNum + "')";

        //            oRecordSet.DoQuery(query);
        //            if (!oRecordSet.EoF)
        //            {
        //                if (!forAdditionalJE)
        //                {
        //                    whtAmtAp += wTaxAmt;
        //                    whtPnAp += pensionAmt;
        //                    whtCoAp += pensionAmt;
        //                } else if (forAdditionalJE && oRecordSet.Fields.Item("U_BDOSPnPhAm").Value != 0 && !isInvoiceType(WTCode))
        //                {
        //                    whtAmtAp += wTaxAmt;
        //                    whtPnAp += pensionAmt;
        //                    whtCoAp += pensionAmt;
        //                }
        //            }
        //        }
        //        else if (oMatrix.Columns.Item("10000127").Cells.Item(i).Specific.Checked && oMatrix.Columns.Item("45").Cells.Item(i).Specific.Value == "19")
        //        {
        //            string docNum = oMatrix.Columns.Item("1").Cells.Item(i).Specific.Value;
        //            string WTCode = "";
        //            checkWTaxCodeFromMatrix(oMatrix, i, out WTCode, "ORPC", "RPC5", "19");

        //            var totalPayment = Convert.ToDecimal(FormsB1.cleanStringOfNonDigits(oMatrix.Columns.Item("24").Cells.Item(i).Specific.String));
        //            CalculateWTax(WTCode, totalPayment, out decimal wTaxAmt, out decimal pensionAmt);

        //            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //            string query = "select \"U_BDOSWhtAmt\", \"U_BDOSPnPhAm\", \"U_BDOSPnCoAm\" from RPC1 " + "\n"
        //            + "where \"DocEntry\" = " + "\n"
        //            + "(select \"DocEntry\" from ORPC " + "\n"
        //            + "where \"DocNum\" = '" + docNum + "')";

        //            oRecordSet.DoQuery(query);
        //            if (!oRecordSet.EoF)
        //            {
        //                if (!forAdditionalJE)
        //                {
        //                    whtAmtAp += wTaxAmt;
        //                    whtPnAp += pensionAmt;
        //                    whtCoAp += pensionAmt;
        //                }
        //                else if (forAdditionalJE && oRecordSet.Fields.Item("U_BDOSPnPhAm").Value != 0 && !isInvoiceType(WTCode))
        //                {
        //                    whtAmtAp += wTaxAmt;
        //                    whtPnAp += pensionAmt;
        //                    whtCoAp += pensionAmt;
        //                }
        //            }
        //        }
        //    }

        //    //on Account
        //    decimal whtAmtOnAcct = 0;
        //    decimal whtPnOnAcct = 0;
        //    decimal whtCoOnAcct = 0;
        //    string totalOnAcct = oForm.Items.Item("13").Specific.Value;
        //    decimal totalOnAccount = getDecimal(totalOnAcct);
        //    if (isPension(wtCode, out rate))
        //    {
        //        whtPnOnAcct = totalOnAccount * 2 / 100;
        //        whtCoOnAcct = totalOnAccount * 2 / 100;
        //        whtAmtOnAcct = (totalOnAccount - whtPnOnAcct) * (decimal)rate / 100;
        //    }
        //    else
        //    {
        //        if (!forAdditionalJE)
        //        {
        //            whtPnOnAcct = 0;
        //            whtCoOnAcct = 0;
        //            whtAmtOnAcct = (totalOnAccount - whtPnOnAcct) * (decimal)getRate(wtCode) / 100;
        //        }
        //    }

        //    Wtax = Math.Round(whtAmtOnAcct + whtAmtAp, 2);
        //    WtPh = Math.Round(whtPnOnAcct + whtPnAp, 2);
        //    WtCo = Math.Round(whtCoOnAcct + whtCoAp, 2);

        //    oForm.Items.Item("BDOSPnPhAm").Specific.Value = WtPh;
        //    oForm.Items.Item("BDOSPnCoAm").Specific.Value = WtCo;
        //    oForm.Items.Item("BDOSWhtAmt").Specific.Value = Wtax;

        //    void CalculateWTax(string wTaxCode, decimal totalAmt, out decimal whtAmt, out decimal pension)
        //    {
        //        var wTaxrate = getRate(wTaxCode)/100;
        //        if (!isPension(wTaxCode, out _))
        //        {
        //            whtAmt = totalAmt / (1 - wTaxrate) * wTaxrate;
        //            pension = 0;
        //        }
        //        else
        //        {
        //            whtAmt = totalAmt / 0.98M / (1 - wTaxrate) * 0.98M * wTaxrate;
        //            pension = totalAmt / 0.98M / (1 - wTaxrate) * 0.02M;
        //        }
        //    }
        //}

        //public static bool isInvoiceType(string WTCode)
        //{
        //    string Category;
        //    SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //    string query = "select \"Category\" from OWHT " + "\n"
        //    + "where \"WTCode\" = '" + WTCode + "'";

        //    oRecordSet.DoQuery(query);
        //    if (!oRecordSet.EoF)
        //    {
        //        Category = oRecordSet.Fields.Item("Category").Value;
        //        if (Category == "I") return true;
        //    }
        //    return false;
        //}

        //--------------------------------------------INTERNET BANK-------------------------------------------->
        public static List<string> importPaymentOrderTBC(PaymentService oPaymentService, List<int> docEntryList, bool importBatchPaymentOrders, string batchName, out string errorText)
        {
            errorText = null;
            string info = null;
            List<string> infoList = new List<string>();

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string query = getQueryForImport(docEntryList, null, null, null, "TBC", true);
            string queryOnlyLocalisationAddOn = OutgoingPayment.getQueryForImportOnlyLocalisationAddOn(docEntryList, null, null, null, "TBC", true);
            try
            {
                oRecordSet.DoQuery(query);
            }
            catch
            {
                oRecordSet.DoQuery(queryOnlyLocalisationAddOn);
            }

            int count = oRecordSet.RecordCount;
            string accountNumber = null; //პაკეტურისთვის
            string accountCurrencyCode = null; //პაკეტურისთვის

            if (count > 0)
            {
                PaymentOrderIo[] paymentOrderArray = new PaymentOrderIo[count]; //TBC

                int i = 1;
                int j = 0;
                while (!oRecordSet.EoF)
                {
                    Dictionary<string, string> dataForTransferType = getDataForTransferType(oRecordSet);
                    string transferType = getTransferType(dataForTransferType, out errorText);
                    Dictionary<string, object> dataForImport = getDataForImport(oRecordSet, dataForTransferType, transferType);
                    string bankProgram = dataForTransferType["bankProgram"]; //ბანკის პროგრამა (გამგზავნი)

                    string transferProgram = bankProgram;

                    if (transferProgram == "TBC")
                    {
                        string status = oRecordSet.Fields.Item("U_status").Value.ToString();

                        if (status != "readyToLoad" && status != "resend" && status != "finishedWithErrors")
                        {
                            errorText = BDOSResources.getTranslate("DocumentSStatusShouldBeTheOneFromTheList") + " : \"" + BDOSResources.getTranslate("readyToLoad") + "\", \"" + BDOSResources.getTranslate("resend") + "\", \"" + BDOSResources.getTranslate("finishedWithErrors") + "\""; //დოკუმენტის სტატუსი უნდა იყოს ერთ-ერთი სიიდან
                            return null;
                        }
                        if (importBatchPaymentOrders)//პაკეტური გადარიცხვა
                        {
                            accountNumber = dataForImport["DebitAccount"].ToString();
                            accountCurrencyCode = dataForImport["DebitAccountCurrencyCode"].ToString();
                        }
                        if (transferType == "TransferToOwnAccountPaymentOrderIo" || transferType == "CurrencyExchangePaymentOrderIo") //გადარიცხვა საკუთარ ანგარიშზე/კონვერტაცია
                        {
                            TransferToOwnAccountPaymentOrderIo oTransferToOwnAccountPaymentOrderIo = new TransferToOwnAccountPaymentOrderIo();
                            MainPaymentService.createTransferToOwnAccountPaymentOrderIo(oTransferToOwnAccountPaymentOrderIo, dataForImport, i, importBatchPaymentOrders);
                            paymentOrderArray[j] = oTransferToOwnAccountPaymentOrderIo;
                        }
                        else if (transferType == "TransferWithinBankPaymentOrderIo") //გადარიცხვა თიბისი ბანკის ფილიალებში
                        {
                            TransferWithinBankPaymentOrderIo oTransferWithinBankPaymentOrderIo = new TransferWithinBankPaymentOrderIo();
                            MainPaymentService.createTransferWithinBankPaymentOrderIo(oTransferWithinBankPaymentOrderIo, dataForImport, i, importBatchPaymentOrders);
                            paymentOrderArray[j] = oTransferWithinBankPaymentOrderIo;
                        }
                        else if (transferType == "TransferToOtherBankNationalCurrencyPaymentOrderIo") //გადარიცხვა სხვა ბანკში (ეროვნული ვალუტა)
                        {
                            TransferToOtherBankNationalCurrencyPaymentOrderIo oTransferToOtherBankNationalCurrencyPaymentOrderIo = new TransferToOtherBankNationalCurrencyPaymentOrderIo();
                            MainPaymentService.createTransferToOtherBankNationalCurrencyPaymentOrderIo(oTransferToOtherBankNationalCurrencyPaymentOrderIo, dataForImport, i, importBatchPaymentOrders);
                            paymentOrderArray[j] = oTransferToOtherBankNationalCurrencyPaymentOrderIo;
                        }
                        else if (transferType == "TransferToOtherBankForeignCurrencyPaymentOrderIo") //გადარიცხვა სხვა ბანკში (უცხოური ვალუტა)
                        {
                            TransferToOtherBankForeignCurrencyPaymentOrderIo oTransferToOtherBankForeignCurrencyPaymentOrderIo = new TransferToOtherBankForeignCurrencyPaymentOrderIo();
                            MainPaymentService.createTransferToOtherBankForeignCurrencyPaymentOrderIo(oTransferToOtherBankForeignCurrencyPaymentOrderIo, dataForImport, i, importBatchPaymentOrders);
                            paymentOrderArray[j] = oTransferToOtherBankForeignCurrencyPaymentOrderIo;
                        }
                        else if (transferType == "TreasuryTransferPaymentOrderIo" || transferType == "TreasuryTransferPaymentOrderIoBP") //საბიუჯეტო გადარიცხვა
                        {
                            TreasuryTransferPaymentOrderIo oTreasuryTransferPaymentOrderIo = new TreasuryTransferPaymentOrderIo();
                            MainPaymentService.createTreasuryTransferPaymentOrderIo(oTreasuryTransferPaymentOrderIo, dataForImport, i, importBatchPaymentOrders);
                            paymentOrderArray[j] = oTreasuryTransferPaymentOrderIo;
                        }
                    }

                    i++;
                    j++;
                    oRecordSet.MoveNext();
                }
                if (!oPaymentService.Equals(null))
                {
                    if (importBatchPaymentOrders == false)//ინდივიდუალური გადარიცხვა
                    {
                        ImportSinglePaymentOrdersResponseIo orderResult = MainPaymentService.importSinglePaymentOrders(oPaymentService, paymentOrderArray, out errorText);

                        if (orderResult != null && string.IsNullOrEmpty(errorText) && orderResult.PaymentOrdersResults != null)
                        {
                            int length = paymentOrderArray.Length;
                            int docEntry;
                            string paymentID;

                            for (i = 0; i < length; i++)
                            {
                                docEntry = Convert.ToInt32(paymentOrderArray[i].documentNumber);

                                SAPbobsCOM.Payments oVendorPayments;
                                oVendorPayments = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments);
                                oVendorPayments.GetByKey(docEntry);
                                paymentID = orderResult.PaymentOrdersResults[i].paymentId.ToString();
                                oVendorPayments.UserFields.Fields.Item("U_paymentID").Value = paymentID;
                                oVendorPayments.UserFields.Fields.Item("U_bPaymentID").Value = "";
                                oVendorPayments.UserFields.Fields.Item("U_bStatus").Value = "empty";
                                oVendorPayments.UserFields.Fields.Item("U_posBPaymnt").Value = 0;

                                int returnCode = oVendorPayments.Update();
                                if (returnCode != 0)
                                {
                                    int errCode;
                                    string errMsg;
                                    Program.oCompany.GetLastError(out errCode, out errMsg);

                                    info = BDOSResources.getTranslate("TheRequestHasBeenSubmittedSuccessfully") + ", " + BDOSResources.getTranslate("ButTheDocumentFailedToUpdate") + "! " + errMsg + ", " + BDOSResources.getTranslate("Document") + " : " + docEntry; //თქვენი დავალება წარმატებით გაიგზავნა, მაგრამ ვერ მოხერხდა დოკუმენტის განახლება
                                    infoList.Add(info);
                                }
                                else
                                {
                                    errorText = null;
                                    info = BDOSResources.getTranslate("TheRequestHasBeenSubmittedSuccessfully") + "! " + BDOSResources.getTranslate("Document") + " : " + docEntry + ", " + BDOSResources.getTranslate("PaymentID") + " : " + paymentID; // თქვენი დავალება წარმატებით გაიგზავნა!
                                    List<int> docEntryListForUpdate = new List<int>();
                                    docEntryListForUpdate.Add(docEntry);

                                    List<string> updateInfo = updateStatusPaymentOrderTBC(oPaymentService, docEntryListForUpdate, out errorText);
                                    if (errorText != null)
                                    {
                                        info = info + "! " + errorText;
                                    }
                                    else
                                    {
                                        info = info + "! " + updateInfo[0];
                                    }
                                    infoList.Add(info);
                                }
                            }
                        }
                        else
                        {
                            info = BDOSResources.getTranslate("OperationCompletedUnSuccessfully");
                            infoList.Add(info);
                        }
                    }
                    else //პაკეტური გადარიცხვა
                    {
                        ImportBatchPaymentOrderResponseIo orderResult = MainPaymentService.importBatchPaymentOrders(oPaymentService, paymentOrderArray, accountNumber, accountCurrencyCode, batchName, out errorText);
                        {
                            if (orderResult != null && string.IsNullOrEmpty(errorText))
                            {
                                int length = paymentOrderArray.Length;
                                int docEntry;
                                string paymentID;

                                for (i = 0; i < length; i++)
                                {
                                    docEntry = Convert.ToInt32(paymentOrderArray[i].documentNumber);

                                    SAPbobsCOM.Payments oVendorPayments;
                                    oVendorPayments = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments);
                                    oVendorPayments.GetByKey(docEntry);
                                    paymentID = orderResult.mygeminiBatchId.ToString();
                                    oVendorPayments.UserFields.Fields.Item("U_paymentID").Value = "";
                                    oVendorPayments.UserFields.Fields.Item("U_bPaymentID").Value = paymentID;
                                    //oVendorPayments.UserFields.Fields.Item("U_bStatus").Value = "empty";
                                    oVendorPayments.UserFields.Fields.Item("U_posBPaymnt").Value = paymentOrderArray[i].position;

                                    int returnCode = oVendorPayments.Update();
                                    if (returnCode != 0)
                                    {
                                        int errCode;
                                        string errMsg;
                                        Program.oCompany.GetLastError(out errCode, out errMsg);
                                        info = BDOSResources.getTranslate("TheRequestHasBeenSubmittedSuccessfully") + ", " + BDOSResources.getTranslate("ButTheDocumentFailedToUpdate") + "! " + errMsg + ", " + BDOSResources.getTranslate("Document") + " : " + docEntry; //თქვენი დავალება წარმატებით გაიგზავნა, მაგრამ ვერ მოხერხდა დოკუმენტის განახლება
                                        infoList.Add(info);
                                    }
                                    else
                                    {
                                        errorText = null;
                                        info = BDOSResources.getTranslate("TheRequestHasBeenSubmittedSuccessfully") + "! " + BDOSResources.getTranslate("Document") + " : " + docEntry + ", " + BDOSResources.getTranslate("BatchPaymentID") + " : " + paymentID; // თქვენი დავალება წარმატებით გაიგზავნა!
                                        List<int> docEntryListForUpdate = new List<int>();
                                        docEntryListForUpdate.Add(docEntry);

                                        List<string> updateInfo = updateStatusPaymentOrderTBC(oPaymentService, docEntryListForUpdate, out errorText);
                                        if (errorText != null)
                                        {
                                            info = info + "! " + errorText;
                                        }
                                        else
                                        {
                                            info = info + "! " + updateInfo[0];
                                        }
                                        infoList.Add(info);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            else
            {
                errorText = BDOSResources.getTranslate("NoDocumentsForOperation") + "! " + BDOSResources.getTranslate("MakeSureToFillInTheDocuments") + "! "; //"ოპერაციისთვის დოკუმენტები არ არსებობს! გადაამოწმეთ დოკუმენტების შევსების სისწორე!";
            }
            return infoList;
        }

        public static List<string> importPaymentOrderBOG(HttpClient client, List<int> docEntryList, bool importBatchPaymentOrders, string batchName, out string errorText)
        {
            errorText = null;
            List<string> infoList = new List<string>();

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string query = getQueryForImport(docEntryList, null, null, null, "BOG", true);
            string queryOnlyLocalisationAddOn = OutgoingPayment.getQueryForImportOnlyLocalisationAddOn(docEntryList, null, null, null, "BOG", true);
            try
            {
                oRecordSet.DoQuery(query);
            }
            catch
            {
                oRecordSet.DoQuery(queryOnlyLocalisationAddOn);
            }

            int count = oRecordSet.RecordCount;

            List<DomesticPayment> domesticPaymentList = new List<DomesticPayment>();
            List<ForeignPayment> foreignPaymentList = new List<ForeignPayment>();
            List<ConversionPayment> conversionPaymentList = new List<ConversionPayment>();

            if (count > 0)
            {
                int i = 1;
                int j = 0;

                while (!oRecordSet.EoF)
                {
                    Dictionary<string, string> dataForTransferType = getDataForTransferType(oRecordSet);
                    string transferType = getTransferType(dataForTransferType, out errorText);
                    Dictionary<string, object> dataForImport = getDataForImport(oRecordSet, dataForTransferType, transferType);
                    string bankProgram = dataForTransferType["bankProgram"]; //ბანკის პროგრამა (გამგზავნი)

                    string transferProgram = bankProgram;

                    if (transferProgram == "BOG")
                    {
                        string status = oRecordSet.Fields.Item("U_status").Value.ToString();

                        if (status != "readyToLoad" && status != "resend" && status != "finishedWithErrors")
                        {
                            errorText = BDOSResources.getTranslate("DocumentSStatusShouldBeTheOneFromTheList") + " : \"" + BDOSResources.getTranslate("readyToLoad") + "\", \"" + BDOSResources.getTranslate("resend") + "\", \"" + BDOSResources.getTranslate("finishedWithErrors") + "\""; //დოკუმენტის სტატუსი უნდა იყოს ერთ-ერთი სიიდან
                            return null;
                        }
                        if (transferType == "TransferToNationalCurrencyPaymentOrderIo" || transferType == "TreasuryTransferPaymentOrderIo" || transferType == "TreasuryTransferPaymentOrderIoBP") //გადარიცხვა (ეროვნული ვალუტა) || სახაზინო
                        {
                            DomesticPayment oDomesticPaymentIo = new DomesticPayment();
                            MainPaymentServiceBOG.createDomesticPaymentOrderIo(oDomesticPaymentIo, dataForImport, importBatchPaymentOrders);
                            CommonFunctions.nullsToEmptyString(oDomesticPaymentIo);
                            domesticPaymentList.Add(oDomesticPaymentIo);
                        }
                        else if (transferType == "TransferToForeignCurrencyPaymentOrderIo") //გადარიცხვა (უცხოური ვალუტა)
                        {
                            ForeignPayment oForeignPaymentIo = new ForeignPayment();
                            MainPaymentServiceBOG.createForeignPaymentOrderIo(oForeignPaymentIo, dataForImport, importBatchPaymentOrders);
                            CommonFunctions.nullsToEmptyString(oForeignPaymentIo);
                            foreignPaymentList.Add(oForeignPaymentIo);
                        }
                        else if (transferType == "CurrencyExchangePaymentOrderIo" && importBatchPaymentOrders == false) //კონვერტაცია
                        {
                            ConversionPayment oConversionPaymentIo = new ConversionPayment();
                            MainPaymentServiceBOG.createConversionPaymentOrderIo(oConversionPaymentIo, dataForImport);
                            CommonFunctions.nullsToEmptyString(oConversionPaymentIo);
                            conversionPaymentList.Add(oConversionPaymentIo);
                        }
                    }

                    i++;
                    j++;
                    oRecordSet.MoveNext();
                }
                if (!client.Equals(null))
                {
                    Task<DocumentKey[]> orderResult = null;
                    Task<long> key;

                    if (domesticPaymentList.Count > 0)
                    {
                        if (importBatchPaymentOrders == false)//ინდივიდუალური გადარიცხვა
                        {
                            orderResult = MainPaymentServiceBOG.importDomesticPaymentOrders(client, domesticPaymentList);
                            if (orderResult != null)
                            {
                                DocumentKey[] orderResultFin = orderResult.Result;
                                resultProcessingBOG(client, orderResultFin, domesticPaymentList, importBatchPaymentOrders, 0, ref infoList);
                            }
                        }
                        else
                        {
                            key = MainPaymentServiceBOG.importBulkDomesticPaymentOrders(client, domesticPaymentList);
                            if (key != null)
                            {
                                resultProcessingBOG(client, null, domesticPaymentList, importBatchPaymentOrders, key.Result, ref infoList);
                            }
                        }
                    }
                    if (foreignPaymentList.Count > 0)
                    {
                        if (importBatchPaymentOrders == false)//ინდივიდუალური გადარიცხვა
                        {
                            orderResult = MainPaymentServiceBOG.importForeignPaymentOrders(client, foreignPaymentList);
                            if (orderResult != null)
                            {
                                DocumentKey[] orderResultFin = orderResult.Result;
                                resultProcessingBOG(client, orderResultFin, foreignPaymentList, importBatchPaymentOrders, 0, ref infoList);
                            }
                        }
                        else
                        {
                            key = MainPaymentServiceBOG.importBulkForeignPaymentOrders(client, foreignPaymentList);
                            if (key != null)
                            {
                                resultProcessingBOG(client, null, foreignPaymentList, importBatchPaymentOrders, key.Result, ref infoList);
                            }
                        }
                    }
                    if (conversionPaymentList.Count > 0)
                    {
                        orderResult = MainPaymentServiceBOG.importConversionPaymentOrders(client, conversionPaymentList);
                        if (orderResult != null)
                        {
                            DocumentKey[] orderResultFin = orderResult.Result;
                            resultProcessingBOG(client, orderResultFin, conversionPaymentList, ref infoList);
                        }
                    }
                }
            }
            else
            {
                errorText = BDOSResources.getTranslate("NoDocumentsForOperation") + "! " + BDOSResources.getTranslate("MakeSureToFillInTheDocuments") + "! "; //"ოპერაციისთვის დოკუმენტები არ არსებობს! გადაამოწმეთ დოკუმენტების შევსების სისწორე!";
            }
            return infoList;
        }

        public static void resultProcessingBOG(HttpClient client, DocumentKey[] orderResult, List<ConversionPayment> paymentList, ref List<string> infoList)
        {
            string errorText = null;
            int length = paymentList.Count();
            int docEntry;
            string paymentID;
            string uniqueID;
            string info = null;

            for (int i = 0; i < length; i++)
            {
                docEntry = Convert.ToInt32(paymentList[i].DocumentNo);

                if (orderResult[i].ResultCode != 0)
                {
                    info = BDOSResources.getTranslate("OperationCompletedUnSuccessfully") + "! \"" + getResultCode(orderResult[i].ResultCode) + "\", " + BDOSResources.getTranslate("Document") + " : " + docEntry;
                    infoList.Add(info);
                }
                else
                {
                    SAPbobsCOM.Payments oVendorPayments;
                    oVendorPayments = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments);
                    oVendorPayments.GetByKey(docEntry);

                    uniqueID = paymentList[i].UniqueId.ToString();
                    paymentID = orderResult[i].UniqueKey.ToString();

                    oVendorPayments.UserFields.Fields.Item("U_paymentID").Value = paymentID;
                    oVendorPayments.UserFields.Fields.Item("U_bPaymentID").Value = "";
                    oVendorPayments.UserFields.Fields.Item("U_bStatus").Value = "empty";
                    oVendorPayments.UserFields.Fields.Item("U_posBPaymnt").Value = 0;
                    oVendorPayments.UserFields.Fields.Item("U_uniqueID").Value = uniqueID;

                    int returnCode = oVendorPayments.Update();
                    if (returnCode != 0)
                    {
                        int errCode;
                        string errMsg;
                        Program.oCompany.GetLastError(out errCode, out errMsg);

                        info = BDOSResources.getTranslate("TheRequestHasBeenSubmittedSuccessfully") + ", " + BDOSResources.getTranslate("ButTheDocumentFailedToUpdate") + "! " + errMsg + ", " + BDOSResources.getTranslate("Document") + " : " + docEntry; //თქვენი დავალება წარმატებით გაიგზავნა, მაგრამ ვერ მოხერხდა დოკუმენტის განახლება
                        infoList.Add(info);
                    }
                    else
                    {
                        errorText = null;
                        info = BDOSResources.getTranslate("TheRequestHasBeenSubmittedSuccessfully") + "! " + BDOSResources.getTranslate("Document") + " : " + docEntry + ", " + BDOSResources.getTranslate("PaymentID") + " : " + paymentID; // თქვენი დავალება წარმატებით გაიგზავნა!
                        List<int> docEntryListForUpdate = new List<int>();
                        docEntryListForUpdate.Add(docEntry);

                        List<string> updateInfo = updateStatusPaymentOrderBOG(client, docEntryListForUpdate, out errorText);
                        if (errorText != null)
                        {
                            info = info + "! " + errorText;
                        }
                        else
                        {
                            info = info + "! " + updateInfo[0];
                        }
                        infoList.Add(info);
                    }
                }
            }
        }

        public static void resultProcessingBOG(HttpClient client, DocumentKey[] orderResult, List<DomesticPayment> paymentList, bool importBatchPaymentOrders, long key, ref List<string> infoList)
        {
            string errorText = null;
            int length = paymentList.Count(); //importBatchPaymentOrders ? 1 : orderResult.Length;
            int docEntry;
            string paymentID;
            string uniqueID;
            string info = null;

            for (int i = 0; i < length; i++)
            {
                docEntry = Convert.ToInt32(paymentList[i].DocumentNo);

                if (importBatchPaymentOrders == false && orderResult[i].ResultCode != 0) //ინდივიდუალური
                {
                    info = BDOSResources.getTranslate("OperationCompletedUnSuccessfully") + "! \"" + getResultCode(orderResult[i].ResultCode) + "\", " + BDOSResources.getTranslate("Document") + " : " + docEntry;
                    infoList.Add(info);
                }
                else if (importBatchPaymentOrders && key == 0) //პაკეტური
                {
                    info = BDOSResources.getTranslate("OperationCompletedUnSuccessfully") + "! " + BDOSResources.getTranslate("Document") + " : " + docEntry;
                    infoList.Add(info);
                }
                else
                {
                    SAPbobsCOM.Payments oVendorPayments;
                    oVendorPayments = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments);
                    oVendorPayments.GetByKey(docEntry);

                    uniqueID = paymentList[i].UniqueId.ToString();

                    if (importBatchPaymentOrders == false) //ინდივიდუალური
                    {
                        paymentID = orderResult[i].UniqueKey.ToString();
                        oVendorPayments.UserFields.Fields.Item("U_paymentID").Value = paymentID;
                        oVendorPayments.UserFields.Fields.Item("U_bPaymentID").Value = "";
                    }
                    else //პაკეტური
                    {
                        paymentID = key.ToString();
                        oVendorPayments.UserFields.Fields.Item("U_paymentID").Value = "";
                        oVendorPayments.UserFields.Fields.Item("U_bPaymentID").Value = paymentID;
                    }

                    oVendorPayments.UserFields.Fields.Item("U_bStatus").Value = "empty";
                    oVendorPayments.UserFields.Fields.Item("U_posBPaymnt").Value = 0;
                    oVendorPayments.UserFields.Fields.Item("U_uniqueID").Value = uniqueID;

                    int returnCode = oVendorPayments.Update();
                    if (returnCode != 0)
                    {
                        int errCode;
                        string errMsg;
                        Program.oCompany.GetLastError(out errCode, out errMsg);

                        info = BDOSResources.getTranslate("TheRequestHasBeenSubmittedSuccessfully") + ", " + BDOSResources.getTranslate("ButTheDocumentFailedToUpdate") + "! " + errMsg + ", " + BDOSResources.getTranslate("Document") + " : " + docEntry; //თქვენი დავალება წარმატებით გაიგზავნა, მაგრამ ვერ მოხერხდა დოკუმენტის განახლება
                        infoList.Add(info);
                    }
                    else
                    {
                        errorText = null;
                        info = BDOSResources.getTranslate("TheRequestHasBeenSubmittedSuccessfully") + "! " + BDOSResources.getTranslate("Document") + " : " + docEntry + ", " + BDOSResources.getTranslate("PaymentID") + " : " + paymentID; // თქვენი დავალება წარმატებით გაიგზავნა!
                        List<int> docEntryListForUpdate = new List<int>();
                        docEntryListForUpdate.Add(docEntry);

                        List<string> updateInfo = updateStatusPaymentOrderBOG(client, docEntryListForUpdate, out errorText);
                        if (errorText != null)
                        {
                            info = info + "! " + errorText;
                        }
                        else
                        {
                            info = info + "! " + updateInfo[0];
                        }
                        infoList.Add(info);
                    }
                }
            }
        }

        public static void resultProcessingBOG(HttpClient client, DocumentKey[] orderResult, List<ForeignPayment> paymentList, bool importBatchPaymentOrders, long key, ref List<string> infoList)
        {
            string errorText = null;
            int length = paymentList.Count();
            int docEntry;
            string paymentID;
            string uniqueID;
            string info = null;

            for (int i = 0; i < length; i++)
            {
                docEntry = Convert.ToInt32(paymentList[i].DocumentNo);

                if (importBatchPaymentOrders == false && orderResult[i].ResultCode != 0) //ინდივიდუალური
                {
                    info = BDOSResources.getTranslate("OperationCompletedUnSuccessfully") + "! \"" + getResultCode(orderResult[i].ResultCode) + "\", " + BDOSResources.getTranslate("Document") + " : " + docEntry;
                    infoList.Add(info);
                }
                else if (importBatchPaymentOrders && key == 0) //პაკეტური
                {
                    info = BDOSResources.getTranslate("OperationCompletedUnSuccessfully") + "! " + BDOSResources.getTranslate("Document") + " : " + docEntry;
                    infoList.Add(info);
                }
                else
                {
                    SAPbobsCOM.Payments oVendorPayments;
                    oVendorPayments = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments);
                    oVendorPayments.GetByKey(docEntry);

                    uniqueID = paymentList[i].UniqueId.ToString();

                    if (importBatchPaymentOrders == false) //ინდივიდუალური
                    {
                        paymentID = orderResult[i].UniqueKey.ToString();
                        oVendorPayments.UserFields.Fields.Item("U_paymentID").Value = paymentID;
                        oVendorPayments.UserFields.Fields.Item("U_bPaymentID").Value = "";
                    }
                    else //პაკეტური
                    {
                        paymentID = key.ToString();
                        oVendorPayments.UserFields.Fields.Item("U_paymentID").Value = "";
                        oVendorPayments.UserFields.Fields.Item("U_bPaymentID").Value = paymentID;
                    }

                    oVendorPayments.UserFields.Fields.Item("U_bStatus").Value = "empty";
                    oVendorPayments.UserFields.Fields.Item("U_posBPaymnt").Value = 0;
                    oVendorPayments.UserFields.Fields.Item("U_uniqueID").Value = uniqueID;

                    int returnCode = oVendorPayments.Update();
                    if (returnCode != 0)
                    {
                        int errCode;
                        string errMsg;
                        Program.oCompany.GetLastError(out errCode, out errMsg);

                        info = BDOSResources.getTranslate("TheRequestHasBeenSubmittedSuccessfully") + ", " + BDOSResources.getTranslate("ButTheDocumentFailedToUpdate") + "! " + errMsg + ", " + BDOSResources.getTranslate("Document") + " : " + docEntry; //თქვენი დავალება წარმატებით გაიგზავნა, მაგრამ ვერ მოხერხდა დოკუმენტის განახლება
                        infoList.Add(info);
                    }
                    else
                    {
                        errorText = null;
                        info = BDOSResources.getTranslate("TheRequestHasBeenSubmittedSuccessfully") + "! " + BDOSResources.getTranslate("Document") + " : " + docEntry + ", " + BDOSResources.getTranslate("PaymentID") + " : " + paymentID; // თქვენი დავალება წარმატებით გაიგზავნა!
                        List<int> docEntryListForUpdate = new List<int>();
                        docEntryListForUpdate.Add(docEntry);

                        List<string> updateInfo = updateStatusPaymentOrderBOG(client, docEntryListForUpdate, out errorText);
                        if (errorText != null)
                        {
                            info = info + "! " + errorText;
                        }
                        else
                        {
                            info = info + "! " + updateInfo[0];
                        }
                        infoList.Add(info);
                    }
                }
            }
        }

        private static string getResultCode(int? resultCode)
        {
            string result = "";

            switch (resultCode)
            {
                case 1:
                    result = "დოკუმენტი არ მოიძებნა";
                    break;
                case 2:
                    result = "დოკუმენტის ტიპი არ მოიძებნა";
                    break;
                case 4:
                    result = "კონვერტაციის კურსი არასწორია";
                    break;
                case 30:
                    result = "მიმღების ანგარიში და ბანკის კოდი არის ცარიელი";
                    break;
                case 31:
                    result = "მიმღები ბანკის კოდი არასწორია";
                    break;
                case 32:
                    result = "მოცემული ტიპის გადარიცხვისთვის შაბლონი არ არის განსაზღვრული";
                    break;
                case 33:
                    result = "გამგზავნის ანგარიში არასწორია";
                    break;
                case 34:
                    result = "გამგზავნის ანგარიშის ვალუტა არასწორია";
                    break;
                case 35:
                    result = "გამგზავნის ანგარიშის ტიპი არასწორია";
                    break;
                case 36:
                    result = "გამგზავნის საიდენტიფიკაციო კოდი არასწორია";
                    break;
                case 37:
                    result = "მიმღების ანგარიში ცარიელია";
                    break;
                case 38:
                    result = "მიმღების ანგარიშის სიგრძე დასშვებზე მეტია";
                    break;
                case 39:
                    result = "მიმღების ანგარიში არასწორია";
                    break;
                case 40:
                    result = "ხაზინაში გადარიცხვის პარამეტრები არასწორია";
                    break;
                case 41:
                    result = "მიმღების Inn არასწორია";
                    break;
                case 42:
                    result = "მიმღების დასახელება ცარიელია";
                    break;
                case 43:
                    result = "მიმღების დასახელების სიგრძე დასშვებზე მეტია";
                    break;
                case 44:
                    result = "მიმღების ანგარიში არასწორია";
                    break;
                case 45:
                    result = "მიმღების ანგარიშის ვალუტა არასწორია";
                    break;
                case 46:
                    result = "მიმღების ანგარიშის ტიპი არასწორია";
                    break;
                case 47:
                    result = "ვალუტირების თარიღი არასწორია";
                    break;
                case 48:
                    result = "ვალუტირების თარიღი არასწორია";
                    break;
                case 49:
                    result = "გადარიცხვის ვალუტა არასწორია";
                    break;
                case 50:
                    result = "თანხის ფორმატი არასწორია";
                    break;
                case 51:
                    result = "საკომისიო თანხის ფორმატი არასწორია";
                    break;
                case 52:
                    result = "გადამხდელის Inn არასწორია";
                    break;
                case 53:
                    result = "გადამხდელის დასახელების სიგრძე დასაშვებზე მეტია";
                    break;
                case 54:
                    result = "გადარიცხვის დანიშნულება ცარიელია";
                    break;
                case 55:
                    result = "გადარიცხვის დანიშნულების სიგრძე დასაშვებზე მეტია";
                    break;
                case 56:
                    result = "დამატებითი ველის სიგრძე დასაშვებზე მეტია";
                    break;
                case 57:
                    result = "დოკუმენტის ვალიდაციისას დაფიქსირდა შეცდომა";
                    break;
                case 58:
                    result = "დოკუმენტის შექმნისას დაფიქსირდა შეცდომა";
                    break;
                case 71:
                    result = "გამგზავნი და მიმღების ანგარიშების ვალუტა არასწორია";
                    break;
                case 72:
                    result = "მიმღების პარამეტრები არასწორია";
                    break;
                case 73:
                    result = "გამგზავნის ანგარიში არასწორია";
                    break;
                case 74:
                    result = "გამგზავნის დასახელება ცარიელია";
                    break;
                case 75:
                    result = "მიმღების დასახელება ცარიელია";
                    break;
                case 76:
                    result = "შუამავალის ინფორმაცია ცარიელია";
                    break;
                case 77:
                    result = "მიმღების ანგარიში ცარიელია";
                    break;
                case 78:
                    result = "მიმღების ანგარიში არასწორია";
                    break;
                case 79:
                    result = "გადარიცხვის დეტალების ფორმატი არასწორია";
                    break;
                case 80:
                    result = "გადარიცხვის დანიშნულების ფორმატი არასწორია";
                    break;
                case 81:
                    result = "ვალუტირების თარიღი არასწორია";
                    break;
                case 82:
                    result = "ვალუტა არასწორია";
                    break;
                case 83:
                    result = "არასაკმარისი თანხა";
                    break;
                case 84:
                    result = "სისტემური ფაზა (მიმდინარეობს დღის დახურვა)";
                    break;
                case 85:
                    result = "საკომისიოს პარამეტრების არასწორია";
                    break;
                case 86:
                    result = "საკომისიოს ანგარიში არასწორია";
                    break;
                case 87:
                    result = "მიმღების ანგარიში არასწორია";
                    break;
                case 90:
                    result = "მიმღების დასახელება გრძელია";
                    break;
                case 91:
                    result = "დანიშნულება გრძელია";
                    break;
                case 110:
                    result = "თანხის ფორმატი არასწორია";
                    break;
                case 111:
                    result = "გადარიცხვის ვალუტა არასწორია";
                    break;
                case 112:
                    result = "გამგზავნის ანგარიში არასწორია";
                    break;
                case 113:
                    result = "მიმღების ანგარიში არასწორია";
                    break;
                case 114:
                    result = "გამგზავნი და მიმღების ანგარიშები ერთიდაიგივეა";
                    break;
                case 115:
                    result = "გამგზავნი და მიმღების ანგარიშების ვალუტა ერთიდაიგივეა";
                    break;
                case 116:
                    result = "გამგზავნი და მიმღები სხვადასხვა კლიენტია";
                    break;
                case 117:
                    result = "გამგზავნი და მიმღები ფილიალი სხვადასხვაა";
                    break;
                case 118:
                    result = "დოკუმენტში პარამეტრები არასწორია";
                    break;
                case 119:
                    result = "დანიშნულება დასაშვებზე გრძელია";
                    break;
                case 120:
                    result = "ოპერაციის ვალუტისთვის ვერ მოიძებნა კურსი";
                    break;
                case 121:
                    result = "ოპერაციის ვალუტა არასწორია";
                    break;
                case 124:
                    result = "doc.result.code.unknown.doc.type";
                    break;
                case 147:
                    result = "document.result.code.wrong.iban.format";
                    break;
                case 150:
                    result = "აუცილებელი ველი";
                    break;
                case 151:
                    result = "ველის მნიშვნელობა უნდა იყოს რიცვითი";
                    break;
                case 152:
                    result = "ველის ფორმატი არასწორია";
                    break;
                case 153:
                    result = "ველის სიგრძე არასწორია";
                    break;
                case 154:
                    result = "ველის მნიშვნელობა არ მოიძებნა, აირჩიეთ ჩამონათვალიდან";
                    break;
                case 177:
                    result = "document.result.code.foreign.transfer.benef.bank.code";
                    break;
                case 333:
                    result = "სამუშაო ნაშთი არ არის საკმარისი. გსურთ საბუთის შექმნა?";
                    break;
                case 363:
                    result = "document.result.code.duplicate";
                    break;
                case 444:
                    result = "თანხა არ არის საკმარისი";
                    break;
                case 999:
                    result = "document.result.code.general.exception";
                    break;
                case 1001:
                    result = "გამგზავნის ანგარიში არასწორია";
                    break;
                case 1002:
                    result = "მიმღების ანგარიში ცარიელია";
                    break;
                case 9999:
                    result = "document.result.code.unknown.error";
                    break;
            }
            return result;
        }

        public static string getQueryForImport(List<int> docEntryList, string account, string startDate, string endDate, string program, bool allDocuments, string docType = "")
        {
            string query = @"SELECT
            ""OVPM"".""DocEntry"" AS ""DocEntry"",
            ""OVPM"".""PrjCode"" AS ""Project"",

            ""OVPM"".""DocNum"" AS ""DocNum"",
            ""OVPM"".""DocType"" AS ""DocType"",
            ""OVPM"".""DocDate"" AS ""DocDate"",
            ""OVPM"".""CardCode"" AS ""CardCode"",
            ""OVPM"".""CardName"" AS ""CardName"",
            ""OVPM"".""Address"" AS ""Address"",
            ""OVPM"".""TrsfrAcct"" AS ""TrsfrAcct"",
            ""OVPM"".""TrsfrSum"" AS ""TrsfrSum"",
            ""OVPM"".""TrsfrSumFC"" AS ""TrsfrSumFC"",
            ""OVPM"".""TrsfrDate"" AS ""TrsfrDate"",
            ""OVPM"".""DocRate"" AS ""DocRate"",
            ""OVPM"".""DiffCurr"" AS ""DiffCurr"",
            ""OVPM"".""DocCurr"" AS ""DocCurr"",
            ""OVPM"".""DocTotal"" AS ""DocTotal"",
            ""OVPM"".""PayToCode"" AS ""PayToCode"",
            ""OVPM"".""IsPaytoBnk"" AS ""IsPaytoBnk"",
            ""OVPM"".""PBnkCnt"" AS ""PBnkCnt"",
            ""OVPM"".""PBnkCode"" AS ""PBnkCode"",
            ""OVPM"".""PBnkAccnt"" AS ""PBnkAccnt"",
            ""OVPM"".""Comments"" AS ""Comments"",
            ""OVPM"".""Status"" AS ""Status"",
            ""OVPM"".""U_opType"" AS ""U_opType"",
            ""OVPM"".""U_status"" AS ""U_status"",
            ""OVPM"".""U_bStatus"" AS ""U_bStatus"",
            ""OVPM"".""U_addDescrpt"" AS ""U_addDescrpt"",
            ""OVPM"".""U_descrpt"" AS ""U_descrpt"",
            ""OVPM"".""U_chrgDtls"" AS ""U_chrgDtls"",
            ""OVPM"".""U_rprtCode"" AS ""U_rprtCode"",
            ""OVPM"".""U_paymentID"" AS ""U_paymentID"",
            ""OVPM"".""U_bPaymentID"" AS ""U_bPaymentID"",
            ""OVPM"".""U_creditAcct"" AS ""U_creditAcct"",
            ""OVPM"".""U_crdtActCur"" AS ""U_crdtActCur"",
            ""OVPM"".""U_tresrCode"" AS ""U_tresrCode"",
            ""OVPM"".""U_dsptchType"" AS ""U_dsptchType"",

            ""VPM4"".""U_employee"" AS ""U_employee"",
	        ""VPM4"".""U_employeeN"" AS ""U_employeeN"",
	        ""VPM4"".""U_creditAcct"" AS ""U_creditAcctEmp"",
	        ""VPM4"".""U_bankCode"" AS ""U_bankCodeEmp"",
	        ""VPM4"".""govID"" AS ""BeneficiaryTaxCodeEmp"",
	        ""VPM4"".""homeStreet"" AS ""BeneficiaryAddressEmp"",
            ""VPM4"".""homeCountr"" AS ""BeneficiaryRegistrationCountryCodeEmp"",
	        ""VPM4"".""BankName"" AS ""BeneficiaryBankNameEmp"",

            ""DSC1"".""BankCode"" AS ""DebitBankCode"",
            ""DSC1"".""Account"" AS ""DebitAccount"",
            ""DSC1"".""GLAccount"" AS ""GLAccount"",
            ""DSC1"".""IBAN"" AS ""IBAN"",
            ""DSC1"".""AcctName"" AS ""AcctName"",
            ""DSC1"".""U_program"" AS ""U_program"",
            ""OCRD"".""LicTradNum"" AS ""BeneficiaryTaxCode"",
            ""OCRD"".""Address"" AS ""BeneficiaryAddress"",
            ""OCRD"".""City"" AS ""RecipientCity"",
            ""OCRD"".""Country"" AS ""BeneficiaryRegistrationCountryCode"",

            CASE
            WHEN ""OACT"".""ActCurr"" = '##'
            THEN '" + Program.LocalCurrency + "' " +
            @"ELSE ""TempOCRN"".""DocCurrCod"" END AS ""DebitAccountCurrencyCode"",

            ""ODSC"".""BankName"" AS ""BeneficiaryBankName"",
            '' AS ""IntermediaryBankCode"",
            '' AS ""IntermediaryBankName"",
            '' AS ""TaxpayerCode"",
            '' AS ""TaxpayerName"",
            ""BDO_INTB"".""U_WSDL"" AS ""WSDL"",
            ""BDO_INTB"".""U_mode"" AS ""mode""

            FROM ""OVPM"" AS ""OVPM""
            INNER JOIN ""DSC1"" ON ""OVPM"".""TrsfrAcct"" = ""DSC1"".""GLAccount""
            LEFT JOIN ""ODSC"" ON ""OVPM"".""PBnkCode"" = ""ODSC"".""BankCode""
            LEFT JOIN ""OCRD"" ON ""OVPM"".""CardCode"" = ""OCRD"".""CardCode""
            INNER JOIN ""OACT"" ON ""OVPM"".""TrsfrAcct"" = ""OACT"".""AcctCode""
            INNER JOIN ""@BDO_INTB"" AS ""BDO_INTB"" ON ""DSC1"".""U_program"" = ""BDO_INTB"".""U_program""
            INNER JOIN ""OCRN"" ON ""OVPM"".""DocCurr"" = ""OCRN"".""CurrCode""

            LEFT JOIN (SELECT
            	 ""VPM4"".""DocNum"",
            	 ""VPM4"".""U_employee"",
            	 ""VPM4"".""U_employeeN"",
            	 ""VPM4"".""U_creditAcct"",
            	 ""VPM4"".""U_bankCode"",
            	 ""OHEM"".""govID"",
            	 ""OHEM"".""homeStreet"",
                 ""OHEM"".""homeCountr"",
            	 ""ODSC"".""BankName""
            	FROM ""VPM4""
            	INNER JOIN ""OHEM"" AS ""OHEM"" ON ""VPM4"".""U_employee"" = ""OHEM"".""empID""
            	INNER JOIN ""ODSC"" AS ""ODSC"" ON ""VPM4"".""U_bankCode"" = ""ODSC"".""BankCode""
            	WHERE ""VPM4"".""U_creditAcct"" IS NOT NULL
            	AND ""VPM4"".""U_creditAcct"" != ''
                GROUP BY ""VPM4"".""DocNum"",
            	 ""VPM4"".""U_employee"",
            	 ""VPM4"".""U_employeeN"",
            	 ""VPM4"".""U_creditAcct"",
            	 ""VPM4"".""U_bankCode"",
            	 ""OHEM"".""govID"",
            	 ""OHEM"".""homeStreet"",
                 ""OHEM"".""homeCountr"",
            	 ""ODSC"".""BankName"") AS ""VPM4"" ON (""OVPM"".""DocEntry"" = ""VPM4"".""DocNum""
            	AND ""OVPM"".""U_opType"" IN ('paymentToEmployee'))

            LEFT JOIN (SELECT ""CurrCode"",  ""DocCurrCod"" FROM ""OCRN"") AS ""TempOCRN"" ON ""OACT"".""ActCurr""  = ""TempOCRN"".""CurrCode""
            LEFT JOIN (SELECT ""CurrCode"",  ""DocCurrCod"" FROM ""OCRN"") AS ""TempOCRN1"" ON ""OVPM"".""U_crdtActCur""  = ""TempOCRN1"".""CurrCode""

            WHERE
            ""OVPM"".""Canceled"" = 'N'
            AND (""OVPM"".""IsPaytoBnk"" = 'Y' OR ""OVPM"".""DocType"" = 'A') AND ""OVPM"".""U_status"" != 'notToUpload' AND ""OVPM"".""U_status"" != 'empty'";

            if (allDocuments == false) //სტატუსის მიხედვით ფილტრი
            {
                query = query + @" AND (""OVPM"".""U_status"" = 'readyToLoad' OR ""OVPM"".""U_status"" = 'resend')";
            }

            if (!string.IsNullOrEmpty(docType)) //დოკუმენტის ტიპის მიხედვით ფილტრი
            {
                query = query + @" AND (""OVPM"".""U_opType"" = '" + docType + "')";
            }

            if (docEntryList != null && docEntryList.Count > 0) //DocEntry-ის მიხედვით ფილტრი
            {
                query = query + @" AND ""OVPM"".""DocEntry"" IN (" + string.Join(",", docEntryList) + ")";
            }
            if (string.IsNullOrEmpty(account) == false) //ანგ.ნომერის მიხედვით ფილტრი
            {
                query = query + @" AND ""DSC1"".""Account""  = '" + account + "'";
            }
            if (string.IsNullOrEmpty(startDate) == false && string.IsNullOrEmpty(endDate) == false) //თარიღის მიხედვით ფილტრი
            {
                query = query + @" AND ""OVPM"".""DocDate""  >= '" + startDate + @"' AND ""OVPM"".""DocDate""  <= '" + endDate + "'";
            }

            query = query + @" AND ""DSC1"".""U_program""  = '" + program + "'";

            query = query + @"ORDER BY ""OVPM"".""DocDate""";

            return query;
        }

        public static string getQueryForImportOnlyLocalisationAddOn(List<int> docEntryList, string account, string startDate, string endDate, string program, bool allDocuments)
        {
            string query = @"SELECT
            ""OVPM"".""DocEntry"" AS ""DocEntry"",
            ""OVPM"".""DocNum"" AS ""DocNum"",
            ""OVPM"".""DocType"" AS ""DocType"",
            ""OVPM"".""DocDate"" AS ""DocDate"",
            ""OVPM"".""CardCode"" AS ""CardCode"",
            ""OVPM"".""CardName"" AS ""CardName"",
            ""OVPM"".""Address"" AS ""Address"",
            ""OVPM"".""TrsfrAcct"" AS ""TrsfrAcct"",
            ""OVPM"".""TrsfrSum"" AS ""TrsfrSum"",
            ""OVPM"".""TrsfrSumFC"" AS ""TrsfrSumFC"",
            ""OVPM"".""TrsfrDate"" AS ""TrsfrDate"",
            ""OVPM"".""DocRate"" AS ""DocRate"",
            ""OVPM"".""DiffCurr"" AS ""DiffCurr"",
            ""OVPM"".""DocCurr"" AS ""DocCurr"",
            ""OVPM"".""DocTotal"" AS ""DocTotal"",
            ""OVPM"".""PayToCode"" AS ""PayToCode"",
            ""OVPM"".""IsPaytoBnk"" AS ""IsPaytoBnk"",
            ""OVPM"".""PBnkCnt"" AS ""PBnkCnt"",
            ""OVPM"".""PBnkCode"" AS ""PBnkCode"",
            ""OVPM"".""PBnkAccnt"" AS ""PBnkAccnt"",
            ""OVPM"".""Comments"" AS ""Comments"",
            ""OVPM"".""Status"" AS ""Status"",
            ""OVPM"".""U_opType"" AS ""U_opType"",
            ""OVPM"".""U_status"" AS ""U_status"",
            ""OVPM"".""U_bStatus"" AS ""U_bStatus"",
            ""OVPM"".""U_addDescrpt"" AS ""U_addDescrpt"",
            ""OVPM"".""U_descrpt"" AS ""U_descrpt"",
            ""OVPM"".""U_chrgDtls"" AS ""U_chrgDtls"",
            ""OVPM"".""U_rprtCode"" AS ""U_rprtCode"",
            ""OVPM"".""U_paymentID"" AS ""U_paymentID"",
            ""OVPM"".""U_bPaymentID"" AS ""U_bPaymentID"",
            ""OVPM"".""U_creditAcct"" AS ""U_creditAcct"",
            ""OVPM"".""U_crdtActCur"" AS ""U_crdtActCur"",
            ""OVPM"".""U_tresrCode"" AS ""U_tresrCode"",
            ""OVPM"".""U_dsptchType"" AS ""U_dsptchType"",

            '' AS ""U_employee"",
	        '' AS ""U_employeeN"",
	        '' AS ""U_creditAcctEmp"",
	        '' AS ""U_bankCodeEmp"",
	        '' AS ""BeneficiaryTaxCodeEmp"",
	        '' AS ""BeneficiaryAddressEmp"",
            '' AS ""BeneficiaryRegistrationCountryCodeEmp"",
	        '' AS ""BeneficiaryBankNameEmp"",

            ""DSC1"".""BankCode"" AS ""DebitBankCode"",
            ""DSC1"".""Account"" AS ""DebitAccount"",
            ""DSC1"".""GLAccount"" AS ""GLAccount"",
            ""DSC1"".""IBAN"" AS ""IBAN"",
            ""DSC1"".""AcctName"" AS ""AcctName"",
            ""DSC1"".""U_program"" AS ""U_program"",
            ""OCRD"".""LicTradNum"" AS ""BeneficiaryTaxCode"",
            ""OCRD"".""Address"" AS ""BeneficiaryAddress"",
            ""OCRD"".""City"" AS ""RecipientCity"",
            ""OCRD"".""Country"" AS ""BeneficiaryRegistrationCountryCode"",


            CASE
            WHEN ""OACT"".""ActCurr"" = '##'
            THEN '" + Program.LocalCurrency + "' " +
            @"ELSE ""TempOCRN"".""DocCurrCod"" END AS ""DebitAccountCurrencyCode"",

            ""ODSC"".""BankName"" AS ""BeneficiaryBankName"",
            '' AS ""IntermediaryBankCode"",
            '' AS ""IntermediaryBankName"",
            '' AS ""TaxpayerCode"",
            '' AS ""TaxpayerName"",
            ""BDO_INTB"".""U_WSDL"" AS ""WSDL"",
            ""BDO_INTB"".""U_mode"" AS ""mode""

            FROM ""OVPM"" AS ""OVPM""
            INNER JOIN ""DSC1"" ON ""OVPM"".""TrsfrAcct"" = ""DSC1"".""GLAccount""
            LEFT JOIN ""ODSC"" ON ""OVPM"".""PBnkCode"" = ""ODSC"".""BankCode""
            LEFT JOIN ""OCRD"" ON ""OVPM"".""CardCode"" = ""OCRD"".""CardCode""
            INNER JOIN ""OACT"" ON ""OVPM"".""TrsfrAcct"" = ""OACT"".""AcctCode""
            INNER JOIN ""@BDO_INTB"" AS ""BDO_INTB"" ON ""DSC1"".""U_program"" = ""BDO_INTB"".""U_program""
            INNER JOIN ""OCRN"" ON ""OVPM"".""DocCurr"" = ""OCRN"".""CurrCode""

            LEFT JOIN (SELECT ""CurrCode"",  ""DocCurrCod"" FROM ""OCRN"") AS ""TempOCRN"" ON ""OACT"".""ActCurr""  = ""TempOCRN"".""CurrCode""
            LEFT JOIN (SELECT ""CurrCode"",  ""DocCurrCod"" FROM ""OCRN"") AS ""TempOCRN1"" ON ""OVPM"".""U_crdtActCur""  = ""TempOCRN1"".""CurrCode""

            WHERE
            ""OVPM"".""Canceled"" = 'N'
            AND (""OVPM"".""IsPaytoBnk"" = 'Y' OR ""OVPM"".""DocType"" = 'A') AND ""OVPM"".""U_status"" != 'notToUpload' AND ""OVPM"".""U_status"" != 'empty'";

            if (allDocuments == false) //სტატუსის მიხედვით ფილტრი
            {
                query = query + @" AND (""OVPM"".""U_status"" = 'readyToLoad' OR ""OVPM"".""U_status"" = 'resend')";
            }
            if (docEntryList != null && docEntryList.Count > 0) //DocEntry-ის მიხედვით ფილტრი
            {
                query = query + @" AND ""OVPM"".""DocEntry"" IN (" + string.Join(",", docEntryList) + ")";
            }
            if (string.IsNullOrEmpty(account) == false) //ანგ.ნომერის მიხედვით ფილტრი
            {
                query = query + @" AND ""DSC1"".""Account""  = '" + account + "'";
            }
            if (string.IsNullOrEmpty(startDate) == false && string.IsNullOrEmpty(endDate) == false) //თარიღის მიხედვით ფილტრი
            {
                query = query + @" AND ""OVPM"".""DocDate""  >= '" + startDate + @"' AND ""OVPM"".""DocDate""  <= '" + endDate + "'";
            }

            query = query + @" AND ""DSC1"".""U_program""  = '" + program + "'";

            query = query + @"ORDER BY ""OVPM"".""DocDate""";

            return query;
        }

        public static List<string> updateStatusPaymentOrderTBC(PaymentService oPaymentService, List<int> docEntryList, out string errorText)
        {
            errorText = null;
            string info = null;
            List<string> infoList = new List<string>();

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string query = @"SELECT
            ""OVPM"".""DocEntry"" AS ""DocEntry"",
            ""OVPM"".""DocNum"" AS ""DocNum"",
            ""OVPM"".""U_paymentID"" AS ""PaymentID"",
            ""OVPM"".""U_bPaymentID"" AS ""BatchPaymentID"",
            ""OVPM"".""U_posBPaymnt"" AS ""PositionBatchPayment""

            FROM ""OVPM"" AS ""OVPM""

            WHERE ""OVPM"".""DocEntry"" IN (" + string.Join(",", docEntryList) + ") " +
            @"AND ((""OVPM"".""U_paymentID"" IS NOT NULL AND ""OVPM"".""U_paymentID"" != '')
            OR (""OVPM"".""U_bPaymentID"" IS NOT NULL AND ""OVPM"".""U_bPaymentID"" != ''))";

            oRecordSet.DoQuery(query);
            int recordCount = oRecordSet.RecordCount;
            if (recordCount > 0)
            {
                while (!oRecordSet.EoF)
                {
                    int docEntry = Convert.ToInt32(oRecordSet.Fields.Item("DocEntry").Value);
                    int docNum = Convert.ToInt32(oRecordSet.Fields.Item("DocNum").Value);

                    string paymentIDTXT = oRecordSet.Fields.Item("PaymentID").Value.ToString();
                    string batchPaymentIDTXT = oRecordSet.Fields.Item("batchPaymentID").Value.ToString();

                    long paymentID = string.IsNullOrEmpty(paymentIDTXT) ? 0 : Convert.ToInt64(paymentIDTXT);
                    long batchPaymentID = string.IsNullOrEmpty(batchPaymentIDTXT) ? 0 : Convert.ToInt64(batchPaymentIDTXT);
                    int positionBatchPayment = Convert.ToInt32(oRecordSet.Fields.Item("PositionBatchPayment").Value);

                    if (!oPaymentService.Equals(null))
                    {
                        if (paymentID != 0 && batchPaymentID == 0) //ინდივიდუალური
                        {
                            GetPaymentOrderStatusResponseIo orderResult = MainPaymentService.refreshSinglePaymentOrderStatus(oPaymentService, paymentID, true, out errorText);

                            if (orderResult != null && string.IsNullOrEmpty(errorText))
                            {
                                string statusInfo;
                                string status = getStatusTBC(orderResult.status, out statusInfo);

                                SAPbobsCOM.Payments oVendorPayments;
                                oVendorPayments = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments);
                                oVendorPayments.GetByKey(docEntry);

                                oVendorPayments.UserFields.Fields.Item("U_status").Value = status;

                                int returnCode = oVendorPayments.Update();
                                if (returnCode != 0)
                                {
                                    int errCode;
                                    string errMsg;
                                    Program.oCompany.GetLastError(out errCode, out errMsg);
                                    info = BDOSResources.getTranslate("TheDocumentSStatusHasBeenUpdatedSuccessfully") + ", " + BDOSResources.getTranslate("ButTheDocumentFailedToUpdate") + "! " + errMsg + BDOSResources.getTranslate("Document") + " : " + docEntry + ", " + BDOSResources.getTranslate("PaymentID") + " : " + paymentID;  //სტატუსების განახლება წარმატებით შესრულდა, მაგრამ ვერ მოხერხდა დოკუმენტის განახლება
                                }
                                else
                                {
                                    info = BDOSResources.getTranslate("TheDocumentSStatusHasBeenUpdatedSuccessfully") + "! " + BDOSResources.getTranslate("Document") + " : " + docEntry + ", " + BDOSResources.getTranslate("PaymentID") + " : " + paymentID; //სტატუსების განახლება წარმატებით შესრულდა
                                }
                            }
                            else
                            {
                                info = BDOSResources.getTranslate("TheDocumentSStatusHasNotBeenUpdatedSuccessfully") + "! " + BDOSResources.getTranslate("Document") + " : " + docEntry + ", " + BDOSResources.getTranslate("PaymentID") + " : " + paymentID; //სტატუსების განახლება წარუმატებლად შესრულდა
                            }
                            infoList.Add(info);
                        }
                        else if (batchPaymentID != 0) //პაკეტური
                        {
                            GetPaymentOrderStatusResponseIo orderResult = MainPaymentService.refreshBatchPaymentOrderStatus(oPaymentService, batchPaymentID, true, out errorText);

                            if (orderResult != null && string.IsNullOrEmpty(errorText)) //&& (orderResult.batchPaymentData != null || recordCount == 1))
                            {
                                SAPbobsCOM.Payments oVendorPayments;
                                oVendorPayments = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments);
                                oVendorPayments.GetByKey(docEntry);
                                string errorDetailEN = "";

                                string batchStatusInfo;
                                string batchStatus = getStatusTBC(orderResult.status, out batchStatusInfo);

                                if (orderResult.batchPaymentData != null)
                                {
                                    PaymentStatusDataIo[] batchPaymentDataResult = orderResult.batchPaymentData;
                                    PaymentStatusDataIo batchPaymentData = Array.Find(batchPaymentDataResult, item => item.position == positionBatchPayment);

                                    string statusInfo;
                                    string status = getStatusTBC(batchPaymentData.paymentStatus, out statusInfo);

                                    oVendorPayments.UserFields.Fields.Item("U_paymentID").Value = string.IsNullOrEmpty(batchPaymentData.paymentId) ? "" : batchPaymentData.paymentId;
                                    oVendorPayments.UserFields.Fields.Item("U_bStatus").Value = batchStatus;

                                    if (string.IsNullOrEmpty(batchPaymentData.errorDetailEN) == false)
                                    {
                                        errorDetailEN = batchPaymentData.errorDetailEN;
                                        oVendorPayments.UserFields.Fields.Item("U_status").Value = "finishedWithErrors";
                                    }
                                    else
                                    {
                                        oVendorPayments.UserFields.Fields.Item("U_status").Value = status;
                                    }
                                }
                                else if (batchStatus == "cancelled")
                                {
                                    oVendorPayments.UserFields.Fields.Item("U_status").Value = batchStatus;
                                    oVendorPayments.UserFields.Fields.Item("U_bStatus").Value = batchStatus;
                                }

                                int returnCode = oVendorPayments.Update();
                                if (returnCode != 0)
                                {
                                    int errCode;
                                    string errMsg;
                                    Program.oCompany.GetLastError(out errCode, out errMsg); //"სტატუსების განახლება წარმატებით შესრულდა
                                    info = BDOSResources.getTranslate("TheDocumentSStatusHasBeenUpdatedSuccessfully") + ", " + BDOSResources.getTranslate("ButTheDocumentFailedToUpdate") + "! " + errMsg + BDOSResources.getTranslate("Document") + " : " + docEntry + ", " + BDOSResources.getTranslate("BatchPaymentID") + " : " + batchPaymentID + "! " + errMsg;  //სტატუსების განახლება წარმატებით შესრულდა, მაგრამ ვერ მოხერხდა დოკუმენტის განახლება
                                }
                                else
                                {
                                    info = BDOSResources.getTranslate("TheDocumentSStatusHasBeenUpdatedSuccessfully") + "! " + BDOSResources.getTranslate("Document") + " : " + docEntry + ", " + BDOSResources.getTranslate("BatchPaymentID") + " : " + batchPaymentID + "! " + errorDetailEN; //სტატუსების განახლება წარმატებით შესრულდა
                                }
                            }
                            else
                            {
                                info = BDOSResources.getTranslate("TheDocumentSStatusHasNotBeenUpdatedSuccessfully") + "! " + BDOSResources.getTranslate("Document") + " : " + docEntry + ", " + BDOSResources.getTranslate("BatchPaymentID") + " : " + batchPaymentID; //სტატუსების განახლება წარუმატებლად შესრულდა
                            }
                            infoList.Add(info);
                        }
                    }
                    oRecordSet.MoveNext();
                }
            }
            else
            {
                errorText = BDOSResources.getTranslate("NoDocumentsForOperation") + "! " + BDOSResources.getTranslate("MakeSureToFillInTheDocuments") + "! "; //"ოპერაციისთვის დოკუმენტები არ არსებობს! გადაამოწმეთ დოკუმენტების შევსების სისწორე!";
            }
            return infoList;
        }

        public static List<string> updateStatusPaymentOrderBOG(HttpClient client, List<int> docEntryList, out string errorText)
        {
            errorText = null;
            string info = null;
            List<string> infoList = new List<string>();

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string query = @"SELECT
            ""OVPM"".""DocEntry"" AS ""DocEntry"",
            ""OVPM"".""DocNum"" AS ""DocNum"",
            ""OVPM"".""U_paymentID"" AS ""PaymentID"",
            ""OVPM"".""U_bPaymentID"" AS ""BatchPaymentID"",
            ""OVPM"".""U_posBPaymnt"" AS ""PositionBatchPayment"",
            ""OVPM"".""U_uniqueID"" AS ""UniqueID""

            FROM ""OVPM"" AS ""OVPM""

            WHERE ""OVPM"".""DocEntry"" IN (" + string.Join(",", docEntryList) + ") " +
            @"AND ((""OVPM"".""U_paymentID"" IS NOT NULL AND ""OVPM"".""U_paymentID"" != '')
            OR (""OVPM"".""U_bPaymentID"" IS NOT NULL AND ""OVPM"".""U_bPaymentID"" != ''))";

            oRecordSet.DoQuery(query);

            if (oRecordSet.RecordCount > 0)
            {
                while (!oRecordSet.EoF)
                {
                    int docEntry = Convert.ToInt32(oRecordSet.Fields.Item("DocEntry").Value);
                    int docNum = Convert.ToInt32(oRecordSet.Fields.Item("DocNum").Value);

                    string paymentIDTXT = oRecordSet.Fields.Item("PaymentID").Value.ToString();
                    string batchPaymentIDTXT = oRecordSet.Fields.Item("batchPaymentID").Value.ToString();

                    long paymentID = string.IsNullOrEmpty(paymentIDTXT) ? 0 : Convert.ToInt64(paymentIDTXT);
                    long batchPaymentID = string.IsNullOrEmpty(batchPaymentIDTXT) ? 0 : Convert.ToInt64(batchPaymentIDTXT);
                    int positionBatchPayment = Convert.ToInt32(oRecordSet.Fields.Item("PositionBatchPayment").Value);
                    string uniqueID = oRecordSet.Fields.Item("UniqueID").Value.ToString();

                    if (!client.Equals(null))
                    {
                        if (paymentID != 0 && batchPaymentID == 0) //ინდივიდუალური
                        {
                            bool batchPaymentOrders = false;

                            Task<List<DocumentStatus>> orderResult = MainPaymentServiceBOG.refreshSinglePaymentOrderStatus(client, paymentID);

                            if (orderResult != null)
                            {
                                List<DocumentStatus> orderResultFin = orderResult.Result;

                                string status = getStatusBOG(orderResultFin[0].Status, batchPaymentOrders);

                                string BulkLineStatus = orderResultFin[0].BulkLineStatus;
                                int? RejectCode = orderResultFin[0].RejectCode;
                                int? ResultCode = orderResultFin[0].ResultCode;
                                string UniqueId = orderResultFin[0].UniqueId.ToString();
                                long? UniqueKey = orderResultFin[0].UniqueKey;

                                if (ResultCode != 0 && ResultCode != null)
                                {
                                    info = BDOSResources.getTranslate("TheDocumentSStatusHasNotBeenUpdatedSuccessfully") + "! \"" + getResultCode(ResultCode) + "\", " + BDOSResources.getTranslate("Document") + " : " + docEntry + ", " + BDOSResources.getTranslate("PaymentID") + " : " + paymentID;
                                    infoList.Add(info);
                                }
                                else
                                {
                                    SAPbobsCOM.Payments oVendorPayments;
                                    oVendorPayments = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments);
                                    oVendorPayments.GetByKey(docEntry);

                                    oVendorPayments.UserFields.Fields.Item("U_status").Value = status;

                                    int returnCode = oVendorPayments.Update();
                                    if (returnCode != 0)
                                    {
                                        int errCode;
                                        string errMsg;
                                        Program.oCompany.GetLastError(out errCode, out errMsg);
                                        info = BDOSResources.getTranslate("TheDocumentSStatusHasBeenUpdatedSuccessfully") + ", " + BDOSResources.getTranslate("ButTheDocumentFailedToUpdate") + "! " + errMsg + BDOSResources.getTranslate("Document") + " : " + docEntry + ", " + BDOSResources.getTranslate("PaymentID") + " : " + paymentID;  //სტატუსების განახლება წარმატებით შესრულდა, მაგრამ ვერ მოხერხდა დოკუმენტის განახლება
                                    }
                                    else
                                    {
                                        info = BDOSResources.getTranslate("TheDocumentSStatusHasBeenUpdatedSuccessfully") + "! " + BDOSResources.getTranslate("Document") + " : " + docEntry + ", " + BDOSResources.getTranslate("PaymentID") + " : " + paymentID; //სტატუსების განახლება წარმატებით შესრულდა
                                    }
                                }
                            }
                            else
                            {
                                info = BDOSResources.getTranslate("TheDocumentSStatusHasNotBeenUpdatedSuccessfully") + "! " + BDOSResources.getTranslate("Document") + " : " + docEntry + ", " + BDOSResources.getTranslate("PaymentID") + " : " + paymentID; //სტატუსების განახლება წარუმატებლად შესრულდა
                            }

                            infoList.Add(info);
                        }
                        else if (string.IsNullOrEmpty(batchPaymentIDTXT) == false) //პაკეტური
                        {
                            bool batchPaymentOrders = true;

                            Task<BulkPaymentStatus> orderResult = MainPaymentServiceBOG.refreshBatchPaymentOrderStatus(client, batchPaymentIDTXT);

                            if (orderResult != null)
                            {
                                string bulkStatus = getStatusBOG(orderResult.Result.Status, batchPaymentOrders);

                                List<DocumentStatus> orderResultFin = orderResult.Result.DocumentStatuses;

                                DocumentStatus oDocumentStatus = orderResultFin.Find(x => x.UniqueId.ToString() == uniqueID);

                                if (!oDocumentStatus.Equals(null))
                                {
                                    string bulkLineStatus = getStatusBOG(oDocumentStatus.BulkLineStatus, batchPaymentOrders);
                                    int? RejectCode = oDocumentStatus.RejectCode;
                                    int? ResultCode = oDocumentStatus.ResultCode;
                                    string UniqueId = oDocumentStatus.UniqueId.ToString();
                                    long? UniqueKey = oDocumentStatus.UniqueKey;

                                    if (ResultCode != 0 && ResultCode != null)
                                    {
                                        info = BDOSResources.getTranslate("TheDocumentSStatusHasNotBeenUpdatedSuccessfully") + "! \"" + getResultCode(ResultCode) + "\", " + BDOSResources.getTranslate("Document") + " : " + docEntry + ", " + BDOSResources.getTranslate("PaymentID") + " : " + paymentID;
                                        infoList.Add(info);
                                    }
                                    else
                                    {
                                        SAPbobsCOM.Payments oVendorPayments;
                                        oVendorPayments = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments);
                                        oVendorPayments.GetByKey(docEntry);

                                        oVendorPayments.UserFields.Fields.Item("U_paymentID").Value = (UniqueKey == null || UniqueKey == 0) ? "" : UniqueKey.ToString();
                                        oVendorPayments.UserFields.Fields.Item("U_bStatus").Value = bulkStatus;
                                        oVendorPayments.UserFields.Fields.Item("U_status").Value = bulkLineStatus;

                                        int returnCode = oVendorPayments.Update();
                                        if (returnCode != 0)
                                        {
                                            int errCode;
                                            string errMsg;
                                            Program.oCompany.GetLastError(out errCode, out errMsg);
                                            info = BDOSResources.getTranslate("TheDocumentSStatusHasBeenUpdatedSuccessfully") + ", " + BDOSResources.getTranslate("ButTheDocumentFailedToUpdate") + "! " + errMsg + BDOSResources.getTranslate("Document") + " : " + docEntry + ", " + BDOSResources.getTranslate("BatchPaymentID") + " : " + batchPaymentIDTXT;  //სტატუსების განახლება წარმატებით შესრულდა, მაგრამ ვერ მოხერხდა დოკუმენტის განახლება
                                        }
                                        else
                                        {
                                            info = BDOSResources.getTranslate("TheDocumentSStatusHasBeenUpdatedSuccessfully") + "! " + BDOSResources.getTranslate("Document") + " : " + docEntry + ", " + BDOSResources.getTranslate("BatchPaymentID") + " : " + batchPaymentIDTXT; //სტატუსების განახლება წარმატებით შესრულდა
                                        }
                                    }
                                }
                            }
                            else
                            {
                                info = BDOSResources.getTranslate("TheDocumentSStatusHasNotBeenUpdatedSuccessfully") + "! " + BDOSResources.getTranslate("Document") + " : " + docEntry + ", " + BDOSResources.getTranslate("BatchPaymentID") + " : " + batchPaymentIDTXT; //სტატუსების განახლება წარუმატებლად შესრულდა
                            }

                            infoList.Add(info);
                        }
                    }
                    oRecordSet.MoveNext();
                }
            }
            else
            {
                errorText = BDOSResources.getTranslate("NoDocumentsForOperation") + "! " + BDOSResources.getTranslate("MakeSureToFillInTheDocuments") + "! "; //"ოპერაციისთვის დოკუმენტები არ არსებობს! გადაამოწმეთ დოკუმენტების შევსების სისწორე!";
            }
            return infoList;
        }

        private static string getStatusTBC(string status, out string info)
        {
            info = null;
            string statusValue;

            if (status == "I" || status == "Initial_state".ToUpper())
            {
                statusValue = "initialState"; //შექმნილი
                info = "Request is saved in database, might not be complete";
            }
            else if (status == "DR" || status == "Draft".ToUpper())
            {
                statusValue = "draft"; //დროებით შენახული
                info = "Transaction in the status Draft is visible only for creator of transaction";
            }
            else if (status == "G" || status == "Registered".ToUpper())
            {
                statusValue = "registered"; //დამატებული
                info = "Registered, complete request is stored in database";
            }
            else if (status == "D" || status == "Deleted".ToUpper())
            {
                statusValue = "deleted"; //წაშლილი
                info = "Set by myGemini if payment was deleted by user (simple delete, no cancelation request)";
            }
            else if (status == "WC" || status == "Waiting_for_certification".ToUpper())
            {
                statusValue = "waitingForCertification"; //ავტორიზაციის მოლოდინში
                info = "Attaching signature is required";
            }
            else if (status == "CERT" || status == "In_progress".ToUpper())
            {
                statusValue = "inProgress"; //დამუშავების პროცესში
                info = "Certified, request is fully certified and can be processed";
            }
            else if (status == "VERIF" || status == "VERIFIED")
            {
                statusValue = "inProgress"; //დამუშავების პროცესში
                info = "Verified, Request is ready for passing to bank system";
            }
            else if (status == "WS")
            {
                statusValue = "inProgress"; //დამუშავების პროცესში
                info = "Waiting for settlement, bank system accepted payment request or payment is status IN PROGRESS in UPS or CBS";
            }
            else if (status == "F" || status == "Finished".ToUpper())
            {
                statusValue = "finished"; //დასრულებული
                info = "Status  PERFORMED from UPS or CBS";
            }
            else if (status == "FL" || status == "Failed".ToUpper())
            {
                statusValue = "failed"; //უარყოფილი
                info = "Status REJECTED from UPS or CBS";
            }
            else if (status == "C" || status == "Cancelled".ToUpper())
            {
                statusValue = "cancelled"; //გაუქმებული
                info = "Set by myGemini if transaction was successfully cancelled by user (cancelation request successfully processed)";
            }
            else if (status == "INIT" || status == "For_Signing".ToUpper())
            {
                statusValue = "forSigning"; //ხელმოსაწერი
                info = "For Signing";
            }
            else if (status == "FINISHED_WITH_ERRORS" || status == "CPE" || status == "Error".ToUpper())
            {
                statusValue = "finishedWithErrors"; //დასრულებულია შეცდომებით
                info = "Finished with Errors";
            }
            else
            {
                statusValue = "readyToLoad"; //მომზადებულია გადასატვირთად
                info = "Status did not Changed";
            }

            return statusValue;
        }

        private static string getStatusBOG(string status, bool batchPaymentOrders)
        {
            string statusValue;

            if (batchPaymentOrders == false)
            {
                if (status == "C" || status == "D")
                {
                    statusValue = "cancelled"; //გაუქმებული
                }
                else if (status == "N")
                {
                    statusValue = "inProgress"; //დამუშავების პროცესში
                }
                else if (status == "P")
                {
                    statusValue = "finished"; //დასრულებული
                }
                else if (status == "A")
                {
                    statusValue = "forSigning"; //ხელმოსაწერი
                }
                else if (status == "S")
                {
                    statusValue = "signed"; //ხელმოწერილი
                }
                else if (status == "T")
                {
                    statusValue = "initialState"; //შექმნილი
                }
                else if (status == "R")
                {
                    statusValue = "failed"; //უარყოფილი
                }
                else if (status == "Z")
                {
                    statusValue = "draft"; //დროებით შენახული
                }
                else
                {
                    statusValue = "readyToLoad"; //მომზადებულია გადასატვირთად
                }
            }
            else
            {
                if (status == "C") //|| status == "D" ?
                {
                    statusValue = "cancelled"; //გაუქმებული
                }
                else if (status == "N" || status == "U")
                {
                    statusValue = "inProgress"; //დამუშავების პროცესში
                }
                else if (status == "P")
                {
                    statusValue = "finished"; //დასრულებული
                }
                else if (status == "S")
                {
                    statusValue = "signed"; //ხელმოწერილი
                }
                else if (status == "T")
                {
                    statusValue = "initialState"; //შექმნილი
                }
                else if (status == "D")
                {
                    statusValue = "finishedWithErrors"; //დასრულებულია შეცდომებით
                }
                else if (status == "Z")
                {
                    statusValue = "draft"; //დროებით შენახული
                }
                else
                {
                    statusValue = "readyToLoad"; //მომზადებულია გადასატვირთად
                }
            }
            return statusValue;
        }

        public static DataTable GetPaymentInvoices(int docEntry, PaymentType ptype, DateTime DocDate)
        {
            var dtPmtInvoices = new DataTable();
            dtPmtInvoices.Columns.Add("InvoiceDocEntry", typeof(int)); //"DocEntry" OrderedInvoicesWithOpenBalance
            dtPmtInvoices.Columns.Add("InstlmntID", typeof(int));
            dtPmtInvoices.Columns.Add("InvType", typeof(int));
            dtPmtInvoices.Columns.Add("DueDate", typeof(DateTime));
            dtPmtInvoices.Columns.Add("InsTotalFC", typeof(double));
            dtPmtInvoices.Columns.Add("InsTotal", typeof(double));
            dtPmtInvoices.Columns.Add("OpenAmountFC", typeof(double));
            dtPmtInvoices.Columns.Add("OpenAmount", typeof(double));
            dtPmtInvoices.Columns.Add("AppliedFC", typeof(double));
            dtPmtInvoices.Columns.Add("DocNum", typeof(int));
            dtPmtInvoices.Columns.Add("DocCur", typeof(string));
            dtPmtInvoices.Columns.Add("DocLine", typeof(int));

            //string PaymentTable = ptype == PaymentType.Draft ? "PDF2" : "VPM2";
            string betweenDays = "";
            if (Program.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
            {
                betweenDays = @"DAYS_BETWEEN(T0.""DueDate"",'" + DocDate.ToString("yyyy-MM-dd") + @"') AS ""OverdueDays"" ";
            }
            else
            {
                betweenDays = @"DATEDIFF(DAY, T0.""DueDate"", '" + DocDate.ToString("yyyy-MM-dd") + @"') AS ""OverdueDays"" ";
            }

            string query = @"SELECT * 
            FROM (SELECT           
	         T0.""DocEntry"" AS ""DocEntry"",
	         T0.""DocNum"" AS ""DocNum"",
	         T0.""DocCur"" AS ""DocCur"",
	         T0.""CardCode"" AS ""CardCode"",
	         T0.""CardName"" AS ""CardName"",
	         T0.""DocDate"" AS ""DocDate"",
	         T0.""DueDate"" AS ""DueDate"",
	         T0.""OpenAmount"" AS ""OpenAmount"",
	         T0.""InsTotal"" AS ""InsTotal"",
	         T0.""OpenAmountFC"" AS ""OpenAmountFC"",
	         T0.""InsTotalFC"" AS ""InsTotalFC"",
	         T0.""ObjType"" AS ""ObjType"",
	         T0.""Comments"" AS ""Comments"",
	         T0.""InstlmntID"" AS ""InstallmentID"",
	         T0.""LineID"" AS ""LineID""," + betweenDays + @"

	        FROM ( SELECT
	         TT0.""DocEntry"",
	         TT0.""DocNum"" AS ""DocNum"",
	         TT0.""DocCur"" AS ""DocCur"",
	         T3.""CardCode"" AS ""CardCode"",
	         T3.""CardName"" AS ""CardName"",
	         TT0.""DocDate"" AS ""DocDate"",
	         TT1.""DueDate"" AS ""DueDate"",
	         TT0.""ObjType"" AS ""ObjType"",
	         TT0.""Comments"" AS ""Comments"",
	         TT1.""InstlmntID"" AS ""InstlmntID"",
	         '0' AS ""LineID"",
	         SUM(TT1.""InsTotal"" - TT1.""PaidToDate""-TT1.""WTSum""+TT1.""WTApplied"") AS ""OpenAmount"",
	 	         SUM(TT1.""InsTotal"") AS ""InsTotal"",
	         SUM(TT1.""InsTotalFC"" - TT1.""PaidFC""-TT1.""WTSumFC""+TT1.""WTAppliedF"") AS ""OpenAmountFC"",
	         SUM(TT1.""InsTotalFC"") AS ""InsTotalFC""
		        FROM OPCH TT0
		        INNER JOIN PCH6 TT1 ON TT0.""DocEntry"" = TT1.""DocEntry""
		        INNER JOIN OCRD T3 ON TT0.""CardCode"" = T3.""CardCode""
		        GROUP BY TT0.""DocEntry"",
	         TT0.""DocNum"",
	         TT0.""DocCur"",
	         T3.""CardCode"",
	         T3.""CardName"",
	         TT0.""DocDate"",
	         TT1.""DueDate"",
	         TT0.""ObjType"",
	         TT0.""Comments"",
	         TT1.""InstlmntID""
		        UNION ALL SELECT
	         TT0.""DocEntry"",
	         TT0.""DocNum"" AS ""DocNum"",
	         TT0.""DocCur"" AS ""DocCur"",
	         T3.""CardCode"" AS ""CardCode"",
	         T3.""CardName"" AS ""CardName"",
	         TT0.""DocDate"" AS ""DocDate"",
	         TT1.""DueDate"" AS ""DueDate"",
	         TT0.""ObjType"" AS ""ObjType"",
	         TT0.""Comments"" AS ""Comments"",
	         TT1.""InstlmntID"" AS ""InstlmntID"",
	         '0' AS ""LineID"",
	         -SUM(TT1.""InsTotal"" - TT1.""PaidToDate""-TT1.""WTSum""+TT1.""WTApplied"")*-1 AS ""OpenAmount"",
	         -SUM(TT1.""InsTotal"")*-1 AS ""InsTotal"",
	         -SUM(TT1.""InsTotalFC"" - TT1.""PaidFC""-TT1.""WTSumFC""+TT1.""WTAppliedF"")*-1 AS ""OpenAmountFC"",
	         -SUM(TT1.""InsTotalFC"")*-1 AS ""InsTotalFC""
		        FROM OCPI TT0
		        INNER JOIN CPI6 TT1 ON TT0.""DocEntry"" = TT1.""DocEntry""
		        INNER JOIN OCRD T3 ON TT0.""CardCode"" = T3.""CardCode""
		        GROUP BY TT0.""DocEntry"",
	         TT0.""DocNum"",
	         TT0.""DocCur"",
	         T3.""CardCode"",
	         T3.""CardName"",
	         TT0.""DocDate"",
	         TT1.""DueDate"",
	         TT0.""ObjType"",
	         TT0.""Comments"",
	         TT1.""InstlmntID"" --

		        UNION ALL SELECT
	         TT0.""DocEntry"",
	         TT0.""DocNum"" AS ""DocNum"",
	         TT0.""DocCur"" AS ""DocCur"",
	         T3.""CardCode"" AS ""CardCode"",
	         T3.""CardName"" AS ""CardName"",
	         TT0.""DocDate"" AS ""DocDate"",
	         TT1.""DueDate"" AS ""DueDate"",
	         TT0.""ObjType"" AS ""ObjType"",
	         TT0.""Comments"" AS ""Comments"",
	         TT1.""InstlmntID"" AS ""InstlmntID"",
	         '0' AS ""LineID"",
	         SUM(TT1.""InsTotal"" - TT1.""PaidToDate""-TT1.""WTSum""+TT1.""WTApplied"")*-1 AS ""OpenAmount"",
	         SUM(TT1.""InsTotal"")*-1 AS ""InsTotal"",
	         SUM(TT1.""InsTotalFC"" - TT1.""PaidFC""-TT1.""WTSumFC""+TT1.""WTAppliedF"")*-1 AS ""OpenAmountFC"",
	         SUM(TT1.""InsTotalFC"")*-1 AS ""InsTotalFC""
		        FROM ORPC TT0
		        INNER JOIN RPC6 TT1 ON TT0.""DocEntry"" = TT1.""DocEntry""
		        INNER JOIN OCRD T3 ON TT0.""CardCode"" = T3.""CardCode""
		        GROUP BY TT0.""DocEntry"",
	         TT0.""DocNum"",
	         TT0.""DocCur"",
	         T3.""CardCode"",
	         T3.""CardName"",
	         TT0.""DocDate"",
	         TT1.""DueDate"",
	         TT0.""ObjType"",
	         TT0.""Comments"",
	         TT1.""InstlmntID"" --

		        UNION ALL SELECT
	         TT0.""DocEntry"",
	         TT0.""DocNum"" AS ""DocNum"",
	         TT0.""DocCur"" AS ""DocCur"",
	         T3.""CardCode"" AS ""CardCode"",
	         T3.""CardName"" AS ""CardName"",
	         TT0.""DocDate"" AS ""DocDate"",
	         TT1.""DueDate"" AS ""DueDate"",
	         TT0.""ObjType"" AS ""ObjType"",
	         TT0.""Comments"" AS ""Comments"",
	         TT1.""InstlmntID"" AS ""InstlmntID"",
	         '0' AS ""LineID"",
	         -SUM(TT1.""InsTotal"" - TT1.""PaidToDate""-TT1.""WTSum""+TT1.""WTApplied"")*-1 AS ""OpenAmount"",
	         -SUM(TT1.""InsTotal"")*-1 AS ""InsTotal"",
	         -SUM(TT1.""InsTotalFC"" - TT1.""PaidFC""-TT1.""WTSumFC""+TT1.""WTAppliedF"")*-1 AS ""OpenAmountFC"",
	         -SUM(TT1.""InsTotalFC"")*-1 AS ""InsTotalFC""
		        FROM ODPO TT0
		        INNER JOIN DPO6 TT1 ON TT0.""DocEntry"" = TT1.""DocEntry""
		        INNER JOIN OCRD T3 ON TT0.""CardCode"" = T3.""CardCode""
		        GROUP BY TT0.""DocEntry"",
	         TT0.""DocNum"",
	         TT0.""DocCur"",
	         T3.""CardCode"",
	         T3.""CardName"",
	         TT0.""DocDate"",
	         TT1.""DueDate"",
	         TT0.""ObjType"",
	         TT0.""Comments"",
	         TT1.""InstlmntID""
		        UNION ALL SELECT
	         TT0.""TransId"" AS ""DocEntry"",
	         TT0.""Number"" AS ""DocNum"",
	         TT0.""TransCurr"" AS ""DocCur"",
	         T3.""CardCode"" AS ""CardCode"",
	         T3.""CardName"" AS ""CardName"",
	         TT0.""RefDate"" AS ""DocDate"",
	         TT1.""DueDate"" AS ""DueDate"",
	         TT0.""ObjType"" AS ""ObjType"",
	         TT1.""LineMemo"" AS ""Comments"",
	         '1' AS ""InstlmntID"",
	         TT1.""Line_ID"" AS ""LineID"",
	         -SUM(TT1.""BalDueDeb"" - TT1.""BalDueCred"") AS ""OpenAmount"",
	         -SUM(TT1.""Debit"" - TT1.""Credit"" ) AS ""InsTotal"",
	         -SUM(TT1.""BalFcDeb"" - TT1.""BalFcCred"") AS ""OpenAmountFC"",
	         -SUM(TT1.""FCDebit"" - TT1.""FCCredit"" ) AS ""InsTotalFC""
		        FROM OJDT TT0
		        INNER JOIN JDT1 TT1 ON TT0.""TransId"" = TT1.""TransId""
		        INNER JOIN OCRD T3 ON TT1.""ShortName"" = T3.""CardCode""
		        --AND TT1.""ShortName"" = N'"" + cardCodeE + @""'
		        AND TT0.""TransType"" IN ('30',
	         '24',
	         '46')
		    AND TT0.""BtfStatus"" = 'O'
		    AND (""TT1"".""BalDueDeb"" > '0'
			    OR ""TT1"".""BalDueCred"" > '0')
		    AND ""TT1"".""DprId"" IS NULL
		     GROUP BY TT0.""TransId"",
	             	     TT0.""Number"",
                         TT0.""TransCurr"",
	             	     T3.""CardCode"",
	             	     T3.""CardName"",
	             	     TT0.""RefDate"",
	             	     TT1.""DueDate"",
	             	     TT0.""ObjType"",
	             	     TT1.""LineMemo"",
                         TT1.""Line_ID"") T0

	        ) AS ""OrderedInvoicesWithOpenBalance""
            INNER JOIN
            (SELECT
                ""PDF2"".""DocNum"" AS ""PaymentDocEntry"",
                ""PDF2"".""DocEntry"" AS ""InvoiceDocEntry"",
	            ""PDF2"".""InvType"",
                ""PDF2"".""AppliedFC"",
                ""PDF2"".""InstId"",
                ""PDF2"".""DocLine""
                FROM ""PDF2"" AS ""PDF2""
                WHERE ""PDF2"".""DocNum"" ='" + docEntry + @"'
            UNION ALL
                SELECT
                ""VPM2"".""DocNum"" AS ""PaymentDocEntry"",
                ""VPM2"".""DocEntry"" AS ""InvoiceDocEntry"",
	            ""VPM2"".""InvType"",
                ""VPM2"".""AppliedFC"",
                ""VPM2"".""InstId"",
                ""VPM2"".""DocLine""
                FROM ""VPM2"" AS ""VPM2""
                WHERE ""VPM2"".""DocNum"" ='" + docEntry + @"') AS ""OutgoingPaymentInvoices""
            ON (""ObjType"" = ""OutgoingPaymentInvoices"".""InvType""
            AND ""DocEntry"" = ""InvoiceDocEntry"" AND  ""InstallmentID"" = ""InstId"")
            WHERE ""PaymentDocEntry"" = " + docEntry +
            @" ORDER BY ""DueDate"", ""DocNum""";

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRecordSet.DoQuery(query);

            while (!oRecordSet.EoF)
            {
                var dtRow = dtPmtInvoices.NewRow();
                dtRow["InvoiceDocEntry"] = oRecordSet.Fields.Item("DocEntry").Value;
                dtRow["InstlmntID"] = oRecordSet.Fields.Item("InstallmentID").Value;
                dtRow["InvType"] = oRecordSet.Fields.Item("InvType").Value;
                dtRow["DueDate"] = oRecordSet.Fields.Item("DueDate").Value;
                dtRow["InsTotalFC"] = oRecordSet.Fields.Item("InsTotalFC").Value;
                dtRow["InsTotal"] = oRecordSet.Fields.Item("InsTotal").Value;
                dtRow["OpenAmount"] = oRecordSet.Fields.Item("OpenAmount").Value;
                dtRow["OpenAmountFC"] = oRecordSet.Fields.Item("OpenAmountFC").Value;
                dtRow["DocLine"] = oRecordSet.Fields.Item("DocLine").Value;
                dtRow["AppliedFC"] = oRecordSet.Fields.Item("AppliedFC").Value;
                dtRow["DocNum"] = oRecordSet.Fields.Item("PaymentDocEntry").Value;
                dtRow["DocCur"] = oRecordSet.Fields.Item("DocCur").Value;

                dtPmtInvoices.Rows.Add(dtRow);

                oRecordSet.MoveNext();
            }
            return dtPmtInvoices;
        }
        public static void FilterInvoiceMatrix(SAPbouiCOM.Form oForm, string agrNo, string prjCode)
        {
            try
            {
                oForm.Freeze(true);

                if (string.IsNullOrEmpty(prjCode) && string.IsNullOrEmpty(agrNo))
                    return;

                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("20").Specific;
                for (int i = 1; i <= oMatrix.RowCount; i++)
                {
                    var prjCodeMtr = oMatrix.Columns.Item("540000141").Cells.Item(i).Specific.Value;
                    var agrNoMtr = oMatrix.Columns.Item("234000060").Cells.Item(i).Specific.Value;

                    if (!string.IsNullOrEmpty(prjCode) && prjCode != prjCodeMtr)
                    {
                        oMatrix.DeleteRow(i);
                        i--;
                    }
                    else if (!string.IsNullOrEmpty(agrNo) && agrNo != agrNoMtr)
                    {
                        oMatrix.DeleteRow(i);
                        i--;
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

        public enum PaymentType
        {
            Draft = 1,
            Payment = 2
        }
        //<--------------------------------------------INTERNET BANK--------------------------------------------
    }
}