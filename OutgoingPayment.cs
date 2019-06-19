﻿using System;
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

            // use blanket agreement rate ranges
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "UseBlaAgRt");
            fieldskeysMap.Add("TableName", "OVPM");
            fieldskeysMap.Add("Description", "Use Blanket Agreement Rates");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 1);
            fieldskeysMap.Add("DefaultValue", "N");

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            GC.Collect();
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
            formItems.Add("Caption", BDOSResources.getTranslate("CreditAccount"));
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
            formItems.Add("Top", top+height+1);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", "Reporting Code");
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
            formItems.Add("Top", top+height+1);
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
            formItems.Add("Caption", BDOSResources.getTranslate("WithhTax"));
            formItems.Add("LinkTo", "BDOSWhtAmt");
            formItems.Add("Visible", true);

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
            formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", true);
            formItems.Add("Visible", true);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            //Fill
            formItems = new Dictionary<string, object>();
            itemName = "FillAmtPen";
            formItems.Add("Caption", BDOSResources.getTranslate("Fill"));
            formItems.Add("Size", 8);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            formItems.Add("Left", left_e + width_e + 5);
            formItems.Add("Width", 16);
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
            formItems.Add("Caption", BDOSResources.getTranslate("PhysEntityPension"));
            formItems.Add("LinkTo", "BDOSPnPhAm");
            formItems.Add("Visible", true);

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
            formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", true);
            formItems.Add("Visible", true);

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
            formItems.Add("Caption", BDOSResources.getTranslate("CompPension"));
            formItems.Add("LinkTo", "BDOSPnCoAm");
            formItems.Add("Visible", true);

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
            formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", true);
            formItems.Add("Visible", true);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            //---------------------- საპენსიო


            // -------------------- Use blanket agreement rates-----------------


            height = oForm.Items.Item("234000005").Height;
            top = oForm.Items.Item("234000005").Top;
            int left = oForm.Items.Item("FillAmtPen").Left - 3;

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
            formItems.Add("DisplayDesc", true);
            formItems.Add("SetAutoManaged", true);
            formItems.Add("FromPane", 2);
            formItems.Add("ToPane", 3);
            formItems.Add("Enabled", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }




            GC.Collect();
        }

        public static void CheckAccounts(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            string CardCode = null;
            bool isError = false;

            string TaxType = oForm.DataSources.DBDataSources.Item("OCRD").GetValue("U_BDO_TaxTyp", 0);

            if (TaxType.Trim() != "12")
            {
                return;
            }

            CardCode = oForm.Items.Item("5").Specific.Value.ToString();

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string query = "SELECT " +
                            "* " +
                            "FROM \"OCRD\" " +
                            "WHERE \"OCRD\".\"CardCode\"='" + CardCode + "'";

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
                    query = "SELECT " +
                                    "* " +
                                    "FROM \"OVTG\" " +
                                    "WHERE \"OVTG\".\"Code\"='" + vatGrp + "'";

                    oRecordSet.DoQuery(query);

                    if (!oRecordSet.EoF)
                    {
                        if (oRecordSet.Fields.Item("U_BDOSAccF").Value == "" || oRecordSet.Fields.Item("Account").Value == "")
                        {
                            errorText = BDOSResources.getTranslate("CheckVatGroupAccounts");
                        }
                    }
                }
            }

            if (isError == true)
            {
                errorText = BDOSResources.getTranslate("CheckVatGroupForBP");
            }
        }

        public static void taxes_OnClick(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            string LiablePrTx = oForm.DataSources.DBDataSources.Item("OVPM").GetValue("U_liablePrTx", 0).Trim();
            string PayNoDoc = oForm.DataSources.DBDataSources.Item("OVPM").GetValue("PayNoDoc", 0).Trim();

            if (LiablePrTx != "Y") //|| PayNoDoc != "Y"
            {
                SAPbouiCOM.EditText oEdit = oForm.Items.Item("PrBaseE").Specific;
                oEdit.Value = "";

                oEdit = oForm.Items.Item("PrBsDscr").Specific;
                oEdit.Value = "";

                oEdit = oForm.Items.Item("AmtPrTxE").Specific;
                oEdit.Value = "";

                oForm.Items.Item("26").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
            }

            //fillAmountTaxes( oForm, out errorText);

            setVisibleFormItems(oForm, out errorText);
        }

        public static void comboSelect(SAPbouiCOM.Form oForm, SAPbouiCOM.ItemEvent pVal, out string errorText)
        {
            errorText = null;

            try
            {
                if (pVal.BeforeAction == false)
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
                else if (pVal.BeforeAction == true)
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

        public static void formDataLoad(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            try
            {
                setVisibleFormItems(oForm, out errorText);

                fillAmountTaxes(oForm, out errorText);
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

        public static void setVisibleFormItems(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            SAPbouiCOM.Item oItem = null;
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

                bool ProfitTaxValuesVisible = (ProfitTaxTypeIsSharing == true && DocType == "S");

                bool docEntryIsEmpty = string.IsNullOrEmpty(docEntry);

                string liablePrTx = oForm.DataSources.DBDataSources.Item("OVPM").GetValue("U_liablePrTx", 0).Trim();
                oForm.Items.Item("liablePrTx").Enabled = (docEntryIsEmpty == true);
                oForm.Items.Item("PrBaseE").Enabled = (liablePrTx == "Y" && docEntryIsEmpty == true);

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
                oForm.Items.Item("ChngDcDt").Visible = (draftKey != "" && oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE);

                //საპენსიო
                bool PensionVisible = (DocType == "S");
                oForm.Items.Item("BDOSWhtS").Visible = PensionVisible;
                oForm.Items.Item("BDOSPnPhS").Visible = PensionVisible;
                oForm.Items.Item("BDOSPnCoS").Visible = PensionVisible;
                oForm.Items.Item("BDOSWhtAmt").Visible = PensionVisible;
                oForm.Items.Item("BDOSPnPhAm").Visible = PensionVisible;
                oForm.Items.Item("BDOSPnCoAm").Visible = PensionVisible;
                oForm.Items.Item("FillAmtPen").Visible = PensionVisible;
                oForm.Items.Item("BDOSWhtAmt").Enabled = PensionVisible && docEntryIsEmpty == true;
                oForm.Items.Item("BDOSPnPhAm").Enabled = PensionVisible && docEntryIsEmpty == true;
                oForm.Items.Item("BDOSPnCoAm").Enabled = PensionVisible && docEntryIsEmpty == true;
                oForm.Items.Item("FillAmtPen").Enabled = PensionVisible && docEntryIsEmpty == true;
                //საპენსიო

                if (docEntryIsEmpty == false)
                {
                    oItem = oForm.Items.Item("opTypeCB");
                    oItem.Enabled = false;
                }
                else
                {
                    oItem = oForm.Items.Item("opTypeCB");
                    oItem.Enabled = true;
                }

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

                oItem = oForm.Items.Item("234000005");
                SAPbouiCOM.EditText oEdit = oItem.Specific;
                oItem = oForm.Items.Item("UsBlaAgRtS");
                if (oEdit.Value != "")
                {
                    oItem.Enabled = true;
                }
                else oItem.Enabled = false;


                Dictionary<string, string> dataForTransferType = getDataForTransferType(oForm);
                string transferType = getTransferType(dataForTransferType, out errorText);

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
                else if (transferType == "TransferToNationalCurrencyPaymentOrderIo")
                {
                    oItem = oForm.Items.Item("chrgDtlsS");
                    oItem.Visible = false;
                    oItem = oForm.Items.Item("chrgDtlsCB");
                    oItem.Visible = false;
                    oItem = oForm.Items.Item("dsptTypeS");
                    oItem.Visible = true;
                    oItem = oForm.Items.Item("dsptTypeCB");
                    oItem.Visible = true;
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
                errorText = ex.Message;
            }
            finally
            {
                GC.Collect();
                oForm.Freeze(false);
                oForm.Update();
            }
        }

        public static void chooseFromList(SAPbouiCOM.Form oForm, SAPbouiCOM.IChooseFromListEvent oCFLEvento, string itemUID, bool beforeAction, out string errorText)
        {
            errorText = null;

            NumberFormatInfo Nfi = new NumberFormatInfo() { NumberDecimalSeparator = "." };

            try
            {
                string sCFL_ID = oCFLEvento.ChooseFromListUID;
                SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(sCFL_ID);

                if (beforeAction == true)
                {
                    if (sCFL_ID == "HouseBankAccount_CFL")
                    {

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
                    setVisibleFormItems(oForm, out errorText);
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

        private static Dictionary<string, string> getDataForTransferType(SAPbouiCOM.Form oForm)
        {
            try
            {
                //დოკუმენტის მონაცემები --->
                string docType = oForm.DataSources.DBDataSources.Item("OVPM").GetValue("DocType", 0).Trim();
                string opType = oForm.DataSources.DBDataSources.Item("OVPM").GetValue("U_opType", 0).Trim();
                string diffCurr = oForm.DataSources.DBDataSources.Item("OVPM").GetValue("DiffCurr", 0).Trim();
                string docCurr = oForm.DataSources.DBDataSources.Item("OVPM").GetValue("DocCurr", 0).Trim();
                docCurr = CommonFunctions.getCurrencyInternationalCode(docCurr);
                docCurr = diffCurr == "Y" ? Program.LocalCurrency : docCurr;
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
                string treasuryCode = oForm.DataSources.DBDataSources.Item("OVPM").GetValue("U_tresrCode", 0).Trim();
                //კონტრაგენტის მონაცემები (მიმღები) <---

                Dictionary<string, string> dataForTransferType = new Dictionary<string, string>();
                dataForTransferType.Add("docType", docType);
                dataForTransferType.Add("opType", opType);
                dataForTransferType.Add("diffCurr", diffCurr);
                dataForTransferType.Add("docCurr", docCurr);
                dataForTransferType.Add("description", description);
                dataForTransferType.Add("chargeDetails", chargeDetails);
                dataForTransferType.Add("reportCode", reportCode);
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
                string docCurr = oRecordSet.Fields.Item("DocCurr").Value.ToString();
                docCurr = CommonFunctions.getCurrencyInternationalCode(docCurr);
                docCurr = diffCurr == "Y" ? Program.LocalCurrency : docCurr;
                string description = oRecordSet.Fields.Item("U_descrpt").Value.ToString();
                string chargeDetails = oRecordSet.Fields.Item("U_chrgDtls").Value.ToString();
                string reportCode = oRecordSet.Fields.Item("U_rprtCode").Value.ToString();
                string RecipientCity = oRecordSet.Fields.Item("RecipientCity").Value.ToString();
                string BeneficiaryRegistrationCountryCode = oRecordSet.Fields.Item("BeneficiaryRegistrationCountryCode").Value.ToString();
                string BeneficiaryAddress = oRecordSet.Fields.Item("BeneficiaryAddress").Value.ToString();
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
                string treasuryCode = oRecordSet.Fields.Item("U_tresrCode").Value.ToString();
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
                string docCurr = dataForTransferType["docCurr"];
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
                        if (transferType == "TreasuryTransferPaymentOrderIo") //სახაზინო გადარიცხვა
                        {
                            CreditAccount = treasuryCode;
                            BeneficiaryBankCode = "TRESGE22"; //მიმღები ბანკის RTGS კოდი / სავალდებულო 
                            BeneficiaryName = "სახელმწიფო ხაზინა"; //მიმღების დასახელება  
                        }
                    }
                }

                if (docType == "A" && opType != "paymentToEmployee")
                {
                    CardCode = "";
                }

                decimal Amount;
                if (TransferCurrency == Program.LocalCurrency)
                    Amount = Convert.ToDecimal(oRecordSet.Fields.Item("TrsfrSum").Value);
                else
                    Amount = Convert.ToDecimal(oRecordSet.Fields.Item("TrsfrSumFC").Value);
                Amount = Math.Round(Amount, 2); //დამრგვალება აუცილებლად უნდა იყოს 2 ციფრამდე, ინტ.ბანკის გამო

                dataForImport.Add("DebitBankCode", bankCode);
                dataForImport.Add("CreditAccount", CreditAccount);
                dataForImport.Add("CreditAccountCurrencyCode", CreditAccountCurrencyCode);
                dataForImport.Add("Currency", TransferCurrency);
                dataForImport.Add("BeneficiaryName", BeneficiaryName);
                dataForImport.Add("Amount", Amount);
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
                        if (bankCode == "TBCBGE22")
                        {
                            if (creditBankCode == bankCode)
                            {
                                transferType = "TransferWithinBankPaymentOrderIo"; //გადარიცხვა თიბისი ბანკის ფილიალებში
                            }
                            else if (creditAcctCurrency == "GEL")
                            {
                                transferType = "TransferToOtherBankNationalCurrencyPaymentOrderIo"; //გადარიცხვა სხვა ბანკში (ეროვნული ვალუტა)
                            }
                            else if (creditAcctCurrency != "GEL")
                            {
                                transferType = "TransferToOtherBankForeignCurrencyPaymentOrderIo"; //გადარიცხვა სხვა ბანკში (უცხოური ვალუტა)
                            }
                        }
                        else if (bankCode == "BAGAGE22")
                        {
                            if (creditAcctCurrency == "GEL")
                            {
                                transferType = "TransferToNationalCurrencyPaymentOrderIo"; //გადარიცხვა (ეროვნული ვალუტა)
                            }
                            else if (creditAcctCurrency != "GEL")
                            {
                                transferType = "TransferToForeignCurrencyPaymentOrderIo"; //გადარიცხვა (უცხოური ვალუტა)
                            }
                        }
                    }
                }
            }
            else if (docType != "A")
            {
                if (bankCode == "TBCBGE22")
                {
                    if (bpBnkCode == "TBCBGE22")
                    {
                        transferType = "TransferWithinBankPaymentOrderIo"; //გადარიცხვა თიბისი ბანკის ფილიალებში
                    }
                    else if (string.IsNullOrEmpty(bpBnkCode) == false && bpBnkCode != "TBCBGE22") //pBnkCode != bankCode)
                    {
                        if (bpBAccountCurrency == "GEL")
                        {
                            transferType = "TransferToOtherBankNationalCurrencyPaymentOrderIo"; //გადარიცხვა სხვა ბანკში (ეროვნული ვალუტა)
                        }
                        else if (bpBAccountCurrency != "GEL")
                        {
                            transferType = "TransferToOtherBankForeignCurrencyPaymentOrderIo"; //გადარიცხვა სხვა ბანკში (უცხოური ვალუტა)
                        }
                    }
                }
                else if (bankCode == "BAGAGE22")
                {
                    if (string.IsNullOrEmpty(bpBnkCode) == false)
                    {
                        if (bpBAccountCurrency == "GEL")
                        {
                            transferType = "TransferToNationalCurrencyPaymentOrderIo"; //გადარიცხვა (ეროვნული ვალუტა)
                        }
                        else if (bpBAccountCurrency != "GEL")
                        {
                            transferType = "TransferToForeignCurrencyPaymentOrderIo"; //გადარიცხვა (უცხოური ვალუტა)
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

                    if (transferType == "TransferWithinBankPaymentOrderIo") //გადარიცხვა თიბისი ბანკის ფილიალებში
                    {
                        creditAcctTmp = bpBAccountCurrency;
                        if (string.IsNullOrEmpty(description))
                        {
                            errorText = BDOSResources.getTranslate("TheFollowingFieldIsMandatory") + " : \"" + oForm.Items.Item("descrptS").Specific.caption + "\""; //აუცილებელია შემდეგი ველების შევსება
                            return;
                        }
                    }
                    else if (transferType == "TreasuryTransferPaymentOrderIo") //საბიუჯეტო გადარიცხვა
                    {
                        creditAcctTmp = docCurr;
                        if (string.IsNullOrEmpty(treasuryCode))
                        {
                            errorText = BDOSResources.getTranslate("TheFollowingFieldIsMandatory") + " : \"" + oForm.Items.Item("tresrCodeS").Specific.caption + "\""; //აუცილებელია შემდეგი ველების შევსება
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
                        if (string.IsNullOrEmpty(chargeDetails) || string.IsNullOrEmpty(description))
                        {
                            errorText = BDOSResources.getTranslate("TheFollowingFieldsAreMandatory") + " : \"" + oForm.Items.Item("chrgDtlsS").Specific.caption + "\", \"" + oForm.Items.Item("descrptS").Specific.caption + "\""; //აუცილებელია შემდეგი ველების შევსება
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
                        if (string.IsNullOrEmpty(chargeDetails) || string.IsNullOrEmpty(description))
                        {
                            errorText = BDOSResources.getTranslate("TheFollowingFieldsAreMandatory") + " : \"" + oForm.Items.Item("chrgDtlsS").Specific.caption + "\", \"" + oForm.Items.Item("descrptS").Specific.caption + "\""; //აუცილებელია შემდეგი ველების შევსება
                            return;
                        }
                    }

                    if (string.IsNullOrEmpty(creditAcctTmp) == false && transferType != "CurrencyExchangePaymentOrderIo") //კონვერტაცია
                    {
                        if (docCurr != creditAcctTmp)
                        {
                            errorText = BDOSResources.getTranslate("DocumentSCurrencyAndTheCreditAccountSCurrencyIsDifferent") + "!"; //დოკუმენტის ვალუტა და მიმღები ანგარიშის ვალუტა განსხვავდება!";
                            return;
                        }
                    }
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
            string RecipientCity = dataForTransferType["RecipientCity"];
            string BeneficiaryAddress = dataForTransferType["BeneficiaryAddress"];
            string BeneficiaryRegistrationCountryCode = dataForTransferType["BeneficiaryRegistrationCountryCode"];
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
                    if (transferType == "TreasuryTransferPaymentOrderIo") //საბიუჯეტო გადარიცხვა
                    {
                        creditAcctTmp = docCurr;
                        if (string.IsNullOrEmpty(treasuryCode) || string.IsNullOrEmpty(description))
                        {
                            errorText = BDOSResources.getTranslate("TheFollowingFieldsAreMandatory") + " : \"" + oForm.Items.Item("descrptS").Specific.caption + "\", \"" + BDOSResources.getTranslate("tresrCodeS") + "\""; //აუცილებელია შემდეგი ველების შევსება
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
                        if (string.IsNullOrEmpty(chargeDetails) || string.IsNullOrEmpty(description) || string.IsNullOrEmpty(reportCode) || string.IsNullOrEmpty(BeneficiaryAddress) || string.IsNullOrEmpty(RecipientCity) || string.IsNullOrEmpty(BeneficiaryRegistrationCountryCode))
                        {
                            errorText = BDOSResources.getTranslate("TheFollowingFieldsAreMandatory") + " : \"" + oForm.Items.Item("chrgDtlsS").Specific.caption + "\", \"" + oForm.Items.Item("rprtCodeS").Specific.caption + "\", \"" + oForm.Items.Item("descrptS").Specific.caption + "\""; //აუცილებელია შემდეგი ველების შევსება
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
                        if (string.IsNullOrEmpty(chargeDetails) || string.IsNullOrEmpty(description))
                        {
                            errorText = BDOSResources.getTranslate("TheFollowingFieldsAreMandatory") + " : \"" + oForm.Items.Item("chrgDtlsS").Specific.caption + "\", \"" + oForm.Items.Item("rprtCodeS").Specific.caption + "\", \"" + oForm.Items.Item("descrptS").Specific.caption + "\""; //აუცილებელია შემდეგი ველების შევსება
                            return;
                        }
                    }

                    if (string.IsNullOrEmpty(creditAcctTmp) == false && transferType != "CurrencyExchangePaymentOrderIo") //კონვერტაცია
                    {
                        if (docCurr != creditAcctTmp)
                        {
                            errorText = BDOSResources.getTranslate("DocumentSCurrencyAndTheCreditAccountSCurrencyIsDifferent") + "!"; //დოკუმენტის ვალუტა და მიმღები ანგარიშის ვალუტა განსხვავდება!";
                            return;
                        }
                    }
                }
            }
        }

        public static void uiApp_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(BusinessObjectInfo.FormTypeEx, Program.currentFormCount);

            //შემოწმება
            if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD)
            {
                if (BusinessObjectInfo.BeforeAction == false & BusinessObjectInfo.ActionSuccess == false)
                {
                    BubbleEvent = false;
                }

                if (BusinessObjectInfo.BeforeAction == true)
                {
                    fillPhysicalEntityTaxes(oForm, out errorText);

                    // მოგების გადასახადი
                    if (ProfitTaxTypeIsSharing == true)
                    {
                        if (oForm.DataSources.DBDataSources.Item("OVPM").GetValue("DocType", 0) == "S")
                        {
                            bool TaxAccountsIsEmpty = ProfitTax.TaxAccountsIsEmpty();
                            if (TaxAccountsIsEmpty == true)
                            {
                                Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("TaxAccounts") + " " + BDOSResources.getTranslate("YouCantLeaveEmpty"));
                                Program.uiApp.MessageBox(BDOSResources.getTranslate("OperationUnsuccesfullSeeLog"));
                                BubbleEvent = false;
                            }
                        }

                        if (oForm.DataSources.DBDataSources.Item("OVPM").GetValue("U_liablePrTx", 0) == "Y")
                        {
                            if (oForm.DataSources.DBDataSources.Item("OVPM").GetValue("U_prBase", 0) == "")
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


                if (BusinessObjectInfo.ActionSuccess != BusinessObjectInfo.BeforeAction)
                {
                    //დოკუმენტის გატარების დროს გატარდეს ბუღლტრული გატარება
                    SAPbouiCOM.DBDataSource DocDBSourcePAYR = oForm.DataSources.DBDataSources.Item(0);
                    if (DocDBSourcePAYR.GetValue("CANCELED", 0) == "N")
                    {
                        string opType = DocDBSourcePAYR.GetValue("U_opType", 0).Trim();
                        if (opType != "salaryPayment" & opType != "paymentToEmployee")
                        {
                            string DocEntry = DocDBSourcePAYR.GetValue("DocEntry", 0);
                            string DocCurrency = DocDBSourcePAYR.GetValue("DocCurr", 0);
                            //decimal DocRate = Convert.ToDecimal( DocDBSourcePAYR.GetValue("DocRate", 0));
                            decimal DocRate = Convert.ToDecimal(FormsB1.cleanStringOfNonDigits(DocDBSourcePAYR.GetValue("DocRate", 0)));
                            string DocNum = DocDBSourcePAYR.GetValue("DocNum", 0);
                            DateTime DocDate = DateTime.ParseExact(DocDBSourcePAYR.GetValue("DocDate", 0), "yyyyMMdd", CultureInfo.InvariantCulture);

                            CommonFunctions.StartTransaction();

                            Program.JrnLinesGlobal = new DataTable();
                            DataTable reLines = null;
                            DataTable JrnLinesDT = createAdditionalEntries(oForm, null, null,null, DocCurrency, out reLines, DocRate);

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

                            //თუ დოკუმენტი გატარდა, მერე ვაკეთებს ბუღალტრულ გატარებას
                            if (BusinessObjectInfo.ActionSuccess == true & BusinessObjectInfo.BeforeAction == false)
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
            }

            if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE)
            {
                if (BusinessObjectInfo.BeforeAction == true)
                {
                    if (Program.cancellationTrans == true && Program.canceledDocEntry != 0)
                    {

                    }
                    else
                    {
                        checkFillDoc(oForm, out errorText);
                        if (errorText != null)
                        {
                            Program.uiApp.SetStatusBarMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, false);
                            BubbleEvent = false;
                        }
                    }
                }
                else if (BusinessObjectInfo.BeforeAction == false && BusinessObjectInfo.ActionSuccess == true)
                {
                    if (Program.cancellationTrans == true && Program.canceledDocEntry != 0)
                    {
                        cancellation(oForm, Program.canceledDocEntry, out errorText);
                        Program.canceledDocEntry = 0;
                    }
                }
            }

            if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD & BusinessObjectInfo.BeforeAction == false)
            {
                formDataLoad(oForm, out errorText);
            }
        }

        public static DataTable createAdditionalEntries(SAPbouiCOM.Form oForm, SAPbobsCOM.GeneralData oGeneralData, DataTable DTSource, DataTable DTSourceVPM2, string DocCurrency, out DataTable reLines, decimal DocRate)
        {
            DataTable jeLines = JournalEntry.JournalEntryTable();

            reLines = ProfitTax.ProfitTaxTable();
            DataRow reLinesRow = null;
            DataTable AccountTable = CommonFunctions.GetOACTTable();
            
            SAPbouiCOM.DBDataSource DBDataSourceTable = null;
            int JEcount = 0;

            SAPbouiCOM.DBDataSource docDBSource = null;
            SAPbouiCOM.DBDataSource BPDataSourceTable = null;

            if (oForm == null)
            {
                JEcount = DTSourceVPM2.Rows.Count;
                ProfitTaxTypeIsSharing = ProfitTax.ProfitTaxTypeIsSharing();
            }
            else
            {
                SAPbouiCOM.DBDataSources docDBSources = oForm.DataSources.DBDataSources;
                DBDataSourceTable = docDBSources.Item("VPM2");
                JEcount = DBDataSourceTable.Size;
                docDBSource = docDBSources.Item("OVPM");

                BPDataSourceTable = docDBSources.Item("OCRD");
            }

            string CardCode = CommonFunctions.getChildOrDbDataSourceValue(BPDataSourceTable, null, DTSource, "CardCode", 0).ToString();

            SAPbobsCOM.BusinessPartners oBP = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);
            string vatCode = "";
            string TaxType = "";
            if (oBP.GetByKey(CardCode) == true)
            {
                vatCode = oBP.VatGroup;
                TaxType = oBP.UserFields.Fields.Item("U_BDO_TaxTyp").Value;
            }

            DocCurrency = DocCurrency == CommonFunctions.getLocalCurrency() ? "" : DocCurrency;

            //დღგ-ს გატარება
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
                    SAPbobsCOM.Documents oInvoice;
                    oInvoice = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices);
                    oInvoice.GetByKey(InvoiceEntry);

                    decimal SumApplied = Convert.ToDecimal(CommonFunctions.getChildOrDbDataSourceValue(DBDataSourceTable, null, DTSourceVPM2, "AppliedFC", i));

                    if (SumApplied == 0)
                    {
                        SumApplied = Convert.ToDecimal(CommonFunctions.getChildOrDbDataSourceValue(DBDataSourceTable, null, DTSourceVPM2, "SumApplied", i));
                    }

                    SumApplied = SumApplied * Convert.ToDecimal(oInvoice.DocRate);

                    SAPbobsCOM.VatGroups oVatCode;
                    oVatCode = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVatGroups);
                    oVatCode.GetByKey(vatCode);

                    string DebitAccount = oVatCode.UserFields.Fields.Item("U_BDOSAccF").Value;
                    string CreditAccount = oVatCode.TaxAccount;

                    decimal vatRate = LandedCosts.GetVatGroupRate(vatCode);

                    decimal TaxAmount = SumApplied * vatRate / (100 + vatRate);
                    decimal TaxAmountFC = DocCurrency == "" ? 0 : TaxAmount / DocRate;
                    if (TaxAmount > 0)
                    {

                        JournalEntry.AddJournalEntryRow(AccountTable, jeLines, "Full", DebitAccount, CreditAccount, TaxAmount, TaxAmountFC, DocCurrency,
                                                            "", "", "", "", "", "", "", "");
                    }

                }
            }

            //მოგების გადასახადის გატარება
            if (ProfitTaxTypeIsSharing == true)
            {
                string U_liablePrTx = CommonFunctions.getChildOrDbDataSourceValue(docDBSource, null, DTSource, "U_liablePrTx", 0).ToString(); //docDBSource.GetValue("U_liablePrTx", 0).Trim();
                decimal NoDocSum = FormsB1.cleanStringOfNonDigits(CommonFunctions.getChildOrDbDataSourceValue(docDBSource, null, DTSource, "NoDocSum", 0).ToString()); ;

                string DebitAccount = CommonFunctions.getOADM("U_BDO_CapAcc").ToString();
                string CreditAccount = CommonFunctions.getOADM("U_BDO_TaxAcc").ToString();
                decimal U_BDO_PrTxRt = Convert.ToDecimal(CommonFunctions.getOADM("U_BDO_PrTxRt").ToString(),CultureInfo.InvariantCulture);

                if (U_liablePrTx == "Y" & NoDocSum > 0)
                {
                    string prBase = CommonFunctions.getChildOrDbDataSourceValue(docDBSource, null, DTSource, "U_prBase", 0).ToString().Trim();
                    decimal TaxAmount = NoDocSum * U_BDO_PrTxRt / (100 - U_BDO_PrTxRt);

                    decimal TaxAmountFC = DocCurrency == "" ? 0 : TaxAmount / DocRate;

                    JournalEntry.AddJournalEntryRow(AccountTable, jeLines, "Full", DebitAccount, CreditAccount, TaxAmount, TaxAmountFC, DocCurrency,
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
                        decimal SumApplied = FormsB1.cleanStringOfNonDigits( CommonFunctions.getChildOrDbDataSourceValue(DBDataSourceTable, null, DTSourceVPM2, "SumApplied", i).ToString());
                        decimal TaxAmount = SumApplied * U_BDO_PrTxRt / (100 - U_BDO_PrTxRt);

                        decimal TaxAmountFC = DocCurrency == "" ? 0 : TaxAmount / DocRate;

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

                            JournalEntry.AddJournalEntryRow(AccountTable, jeLines, "Full", DebitAccount, CreditAccount, TaxAmount, TaxAmountFC, DocCurrency,
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

            // პენსია            
            
            string wtCode = CommonFunctions.getChildOrDbDataSourceValue(BPDataSourceTable, null, DTSource, "WtCode", 0).ToString();

            string WTLiable = CommonFunctions.getChildOrDbDataSourceValue(BPDataSourceTable, null, DTSource, "WTLiable", 0).ToString();
            string U_BDOSPhisTx = CommonFunctions.getValue("OWHT", "U_BDOSPhisTx", "WTCode", wtCode).ToString();

            bool physicalEntityTax = (WTLiable == "Y" && U_BDOSPhisTx == "Y");
            

            if (physicalEntityTax)
            {
                string pensionCoWTCode = CommonFunctions.getOADM("U_BDOSPnCoP").ToString();
                string pensionPhWTCode = CommonFunctions.getOADM("U_BDOSPnPh").ToString();

                string DebitAccount;
                string CreditAccount;
                string DistrRule1 = "";
                string DistrRule2 = "";
                string DistrRule3 = "";
                string DistrRule4 = "";
                string DistrRule5 = "";

                string Project = CommonFunctions.getChildOrDbDataSourceValue(docDBSource, null, DTSource, "PrjCode", 0).ToString();

                decimal WhtAmount = FormsB1.cleanStringOfNonDigits(CommonFunctions.getChildOrDbDataSourceValue(docDBSource, null, DTSource, "U_BDOSWhtAmt", 0).ToString());
                decimal WhtAmountFC = DocCurrency == "" ? 0 : WhtAmount / DocRate;

                decimal PhysPensionAmount = FormsB1.cleanStringOfNonDigits(CommonFunctions.getChildOrDbDataSourceValue(docDBSource, null, DTSource, "U_BDOSPnPhAm", 0).ToString());
                decimal PhysPensionAmountFC = DocCurrency == "" ? 0 : PhysPensionAmount / DocRate;

                decimal CompanyPensionAmount = FormsB1.cleanStringOfNonDigits(CommonFunctions.getChildOrDbDataSourceValue(docDBSource, null, DTSource, "U_BDOSPnCoAm", 0).ToString());
                decimal CompanyPensionAmountFC = DocCurrency == "" ? 0 : CompanyPensionAmount / DocRate;

                if (WhtAmount != 0 && PhysPensionAmount != 0)
                {
                    DebitAccount = CommonFunctions.getValue("OWHT", "Account", "WTCode", wtCode).ToString(); //BP-ს ძირითადი WTCode-ს ანგარიში
                    JournalEntry.AddJournalEntryRow(AccountTable, jeLines, "OnlyDebit", DebitAccount, "", (WhtAmount + PhysPensionAmount), (WhtAmountFC + PhysPensionAmountFC), DocCurrency,
                                                        DistrRule1, DistrRule2, DistrRule3, DistrRule4, DistrRule5, Project, "", "");

                    CreditAccount = CommonFunctions.getValue("OWHT", "U_BdgtDbtAcc", "WTCode", wtCode).ToString(); //BP-ს ძირითადი WTCode-ს ვალდებულების ანგარიში
                    JournalEntry.AddJournalEntryRow(AccountTable, jeLines, "OnlyCredit", "", CreditAccount, WhtAmount, WhtAmountFC, DocCurrency,
                                                        DistrRule1, DistrRule2, DistrRule3, DistrRule4, DistrRule5, Project, "", "");

                    CreditAccount = CommonFunctions.getValue("OWHT", "Account", "WTCode", pensionPhWTCode).ToString(); //U_BdgtDbtAcc დასაქმებულის საპენსიოს ვალდებულების ანგარიში
                    JournalEntry.AddJournalEntryRow(AccountTable, jeLines, "OnlyCredit", "", CreditAccount, PhysPensionAmount, PhysPensionAmountFC, DocCurrency,
                                    DistrRule1, DistrRule2, DistrRule3, DistrRule4, DistrRule5, Project, "", "");
                }

                if (CompanyPensionAmount != 0)
                {
                    DebitAccount = CommonFunctions.getValue("OWHT", "Account", "WTCode", pensionCoWTCode).ToString(); // დამსაქმებლის საპენსიოს ანგარიში
                    CreditAccount = CommonFunctions.getValue("OWHT", "U_BdgtDbtAcc", "WTCode", pensionCoWTCode).ToString(); // დამსაქმებლის საპენსიოს ვალდებულების ანგარიში
                    JournalEntry.AddJournalEntryRow(AccountTable, jeLines, "Full", DebitAccount, CreditAccount, CompanyPensionAmount, CompanyPensionAmountFC, DocCurrency,
                                                        DistrRule1, DistrRule2, DistrRule3, DistrRule4, DistrRule5, Project, "", "");
                }
            }

            // პენსია

            return jeLines;
        }

        public static void JrnEntry(string DocEntry, string DocNum, DateTime DocDate, DataTable JrnLinesDT, DataTable reLines, out string errorText)
        {
            errorText = null;

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

        public static void cancellation(SAPbouiCOM.Form oForm, int docEntry, out string errorText)
        {
            errorText = null;

            try
            {
                JournalEntry.cancellation(oForm, docEntry, "46", out errorText);
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

        public static void uiApp_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            if (FormUID == "OutgoingPaymentNewDate")
            {
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK & pVal.BeforeAction == false)
                {
                    SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                    if (pVal.ItemUID == "1")
                    {
                        string newDate = oForm.Items.Item("newDate").Specific.Value;
                        changeDocDateRate(oForm, newDate);
                    }
                }
            }
            else
            {

                if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
                {
                    SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
                    {
                        if (pVal.ItemUID == "1" && pVal.BeforeAction == true)
                        {
                            CommonFunctions.fillDocRate(oForm, "OVPM", "OVPM");
                        }
                    }

                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD & pVal.BeforeAction == true)
                    {
                        OutgoingPayment.createFormItems(oForm, out errorText);
                        Program.FORM_LOAD_FOR_VISIBLE = true;
                        Program.FORM_LOAD_FOR_ACTIVATE = true;

                        formDataLoad(oForm, out errorText);
                    }

                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_COMBO_SELECT)
                    {
                        if ((pVal.ItemUID == "opTypeCB" || pVal.ItemUID == "18" || pVal.ItemUID == "107") && pVal.BeforeAction == false)
                        {
                            if (pVal.ItemUID == "107" && oForm.DataSources.DBDataSources.Item("OVPM").GetValue("IsPaytoBnk", 0).Trim() != "Y")
                            {
                                return;
                            }
                            setVisibleFormItems(oForm, out errorText);
                        }
                        oForm.Freeze(true);
                        comboSelect(oForm, pVal, out errorText);
                        oForm.Freeze(false);
                    }

                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE & pVal.BeforeAction == false)
                    {
                        if (Program.FORM_LOAD_FOR_ACTIVATE == true)
                        {
                            setVisibleFormItems(oForm, out errorText); ;
                            Program.FORM_LOAD_FOR_ACTIVATE = false;
                        }
                    }

                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
                    {
                        if ((pVal.ItemUID == "57" || pVal.ItemUID == "56" || pVal.ItemUID == "58") && pVal.BeforeAction == false)
                        {
                            setVisibleFormItems(oForm, out errorText); ;
                        }

                        if (pVal.ItemUID == "ChngDcDt" && pVal.BeforeAction == false)
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
                    }

                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
                    {
                        if ((pVal.ItemUID == "liablePrTx" || pVal.ItemUID == "37") && pVal.BeforeAction == false)
                        {
                            oForm.Freeze(true);
                            taxes_OnClick(oForm, out errorText);
                            oForm.Freeze(false);
                        }
                    }

                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_VALIDATE)
                    {
                        if ((pVal.ItemUID == "5") && pVal.BeforeAction == false)
                        {
                            setVisibleFormItems(oForm, out errorText);
                        }
                        
                    }
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_VALIDATE & pVal.BeforeAction == false)
                    {
                        if (pVal.ItemUID == "234000005")
                        {
                            setVisibleFormItems(oForm, out errorText);
                        }
                    }

                    if(pVal.ItemUID == "UsBlaAgRtS" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
                    {
                        SAPbouiCOM.CheckBox oCheckBox = (SAPbouiCOM.CheckBox)oForm.Items.Item("UsBlaAgRtS").Specific;
                        if (oCheckBox.Checked == true)
                        {
                            CommonFunctions.fillDocRate(oForm, "OVPM", "OVPM");
                        }
                    }

                    if (pVal.ItemUID == "creditActE" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                    {
                        SAPbouiCOM.IChooseFromListEvent oCFLEvento = null;
                        oCFLEvento = ((SAPbouiCOM.IChooseFromListEvent)(pVal));

                        chooseFromList(oForm, oCFLEvento, pVal.ItemUID, pVal.BeforeAction, out errorText);
                    }

                    if (pVal.ItemUID == "PrBaseE" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST & pVal.BeforeAction == false)
                    {
                        SAPbouiCOM.IChooseFromListEvent oCFLEvento = null;
                        oCFLEvento = ((SAPbouiCOM.IChooseFromListEvent)(pVal));

                        chooseFromList(oForm, oCFLEvento, pVal.ItemUID, pVal.BeforeAction, out errorText);
                    }

                    if (pVal.ItemUID == "FillAmtTxs" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK & pVal.BeforeAction == false & pVal.InnerEvent == false)
                    {
                        oForm.Freeze(true);
                        try
                        {
                            fillAmountTaxes(oForm, out errorText);
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
                            oForm.Freeze(false);
                            GC.Collect();
                        }
                    }

                    if (pVal.ItemUID == "FillAmtPen" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.BeforeAction == false && pVal.InnerEvent == false)
                    {
                        fillPhysicalEntityTaxes(oForm, out errorText);
                    }

                    if (pVal.ItemUID == "1" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.BeforeAction == true)
                    {
                        fillPhysicalEntityTaxes(oForm, out errorText);
                    }

                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_RESIZE && pVal.BeforeAction == false)
                    {
                        oForm.Freeze(true);
                        resizeForm(oForm, out errorText);
                        oForm.Freeze(false);
                    }

                    if (Program.openPaymentMeans == true && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE && pVal.BeforeAction == false)
                    {
                        Program.openPaymentMeans = false;
                        setVisibleFormItems(oForm, out errorText);
                    }
                }
            }
        }

        private static void changeDocDateRate(SAPbouiCOM.Form oFormDate, string newDate)
        {
            //SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm("426", Program.currentFormCount);
            SAPbouiCOM.Form oForm = CurrentForm;
            DateTime DocDate = Convert.ToDateTime(DateTime.ParseExact(newDate, "yyyyMMdd", CultureInfo.InvariantCulture));
            string DocCurr = oForm.DataSources.DBDataSources.Item(0).GetValue("DocCurr", 0);
            string DocEntry = oForm.DataSources.DBDataSources.Item(0).GetValue("DocEntry", 0);
            decimal TrsfrSum = Convert.ToDecimal(oForm.DataSources.DBDataSources.Item(0).GetValue("TrsfrSum", 0), CultureInfo.InvariantCulture);
            decimal NoDocSumFC = Convert.ToDecimal(oForm.DataSources.DBDataSources.Item(0).GetValue("NoDocSumFC", 0), CultureInfo.InvariantCulture);
            string DiffCurr = oForm.DataSources.DBDataSources.Item(0).GetValue("DiffCurr", 0);

            string bpCurrency = oForm.DataSources.DBDataSources.Item(1).GetValue("Currency", 0);

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset oRecordSetPDF2 = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            if (bpCurrency == "##" || String.IsNullOrEmpty(bpCurrency))
            {
                var invoicesDT = GetPaymentInvoices(Convert.ToInt32(DocEntry), PaymentType.Draft, DocDate);
                var currencies = invoicesDT.AsEnumerable().Select(x => x["DocCur"]);
                string firstcurrency = (string)currencies.FirstOrDefault();
                int otherCurrenciesCount = invoicesDT.AsEnumerable().Where(x => (string)x["DocCur"] != firstcurrency).Count();
                int firstCurrencyCount = invoicesDT.AsEnumerable().Where(x => (string)x["DocCur"] == firstcurrency).Count();

                if (otherCurrenciesCount > 0)
                {
                    Program.uiApp.SetStatusBarMessage(BDOSResources.getTranslate("InvoicesDifferentCurrenciesError"), SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    return;
                }
                else
                {
                    DiffCurr = (DocCurr != firstcurrency) ? "Y" : "N";
                    DocCurr = firstcurrency;
                }
            }

            string errorText = null;
            decimal DocRate = 0;
            decimal TrsfrSumFC = 0;

            SAPbobsCOM.SBObob oSBOBob = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);

            if (DocCurr == CurrencyB1.getMainCurrency(out errorText))
            {
                string queryOPDF = @"update OPDF 
                                    set 
                                    ""DocDate"" = '" + newDate + @"',
                                    ""TaxDate"" = '" + newDate + @"', 
                                    ""VatDate"" = '" + newDate + @"', 
                                    ""DocDueDate"" = '" + newDate + @"'
                                    
                                    where ""DocEntry"" = " + DocEntry + @"";
                oRecordSet.DoQuery(queryOPDF);
            }
            else if (DiffCurr == "N")
            {
                SAPbobsCOM.Recordset RateRecordset = oSBOBob.GetCurrencyRate(DocCurr, DocDate);

                while (!RateRecordset.EoF)
                {
                    DocRate = Convert.ToDecimal(RateRecordset.Fields.Item("CurrencyRate").Value);
                    RateRecordset.MoveNext();
                }

                string queryOPDF = @"update OPDF 
                                    set 
                                    ""DocDate"" = '" + newDate + @"',
                                    ""TaxDate"" = '" + newDate + @"', 
                                    ""VatDate"" = '" + newDate + @"', 
                                    ""DocDueDate"" = '" + newDate + @"',
                                    ""DocRate"" = " + DocRate.ToString(CultureInfo.InvariantCulture) + @"
                                    where ""DocEntry"" = " + DocEntry + @"";
                oRecordSet.DoQuery(queryOPDF);
            }
            else
            {
                SAPbobsCOM.Recordset RateRecordset = oSBOBob.GetCurrencyRate(DocCurr, DocDate);

                while (!RateRecordset.EoF)
                {
                    DocRate = Convert.ToDecimal(RateRecordset.Fields.Item("CurrencyRate").Value);
                    RateRecordset.MoveNext();
                }

                TrsfrSumFC = CommonFunctions.roundAmountByGeneralSettings(TrsfrSum / DocRate, "Sum");
                decimal TrsfrSumFCOld = TrsfrSumFC;

                string query = @"select * from PDF2
                       where ""DocNum"" = " + DocEntry + @"";
                oRecordSet.DoQuery(query);

                while (!oRecordSet.EoF)
                {
                    /////////////////////
                    decimal AppliedFC = Convert.ToDecimal(oRecordSet.Fields.Item("AppliedFC").Value);
                    string DocNum = oRecordSet.Fields.Item("DocNum").Value.ToString();
                    string InvoiceDocEntry = oRecordSet.Fields.Item("DocEntry").Value.ToString();
                    string InstId = oRecordSet.Fields.Item("InstId").Value.ToString();

                    if (TrsfrSumFC == 0)
                    {
                        break;
                    }

                    if (AppliedFC > TrsfrSumFC)
                    {
                        AppliedFC = TrsfrSumFC;
                    }

                    TrsfrSumFC = TrsfrSumFC - AppliedFC;

                    string queryPDF2 = @"update PDF2
                    set 
                    ""DocRate"" = " + DocRate.ToString(CultureInfo.InvariantCulture) + @",
                    ""AppliedFC"" =" + AppliedFC.ToString(CultureInfo.InvariantCulture) + @",
                    ""BfDcntSumF"" = " + AppliedFC.ToString(CultureInfo.InvariantCulture) + @",
                    ""BfNetDcntF"" = " + AppliedFC.ToString(CultureInfo.InvariantCulture) + @"
                    where ""DocNum"" = " + DocNum + @"
                    and ""DocEntry"" =  " + InvoiceDocEntry + @"
                    and ""InstId"" = " + InstId + "";

                    oRecordSetPDF2.DoQuery(queryPDF2);
                    /////////////////////

                    oRecordSet.MoveNext();
                }

                query = @"update OPDF 
                                    set 
                                    ""DocDate"" = '" + newDate + @"',
                                    ""TaxDate"" = '" + newDate + @"', 
                                    ""VatDate"" = '" + newDate + @"', 
                                    ""DocDueDate"" = '" + newDate + @"'" +
                                    ((bpCurrency == "##") ? "" :
                                    @",""DocRate"" = " + DocRate.ToString(CultureInfo.InvariantCulture) + @",
                                    ""TrsfrSumFC"" = " + TrsfrSumFCOld.ToString(CultureInfo.InvariantCulture) + @",
                                    ""DocTotalFC"" = " + TrsfrSumFCOld.ToString(CultureInfo.InvariantCulture)) +

                                    @"where ""DocEntry"" = " + DocEntry + @"";
                oRecordSet.DoQuery(query);

                //დარჩენილი თანხით ინვოისების ჩახურვა
                if (TrsfrSumFC != 0)
                {
                    DataTable Invoices = GetPaymentInvoices(Convert.ToInt32(DocEntry), PaymentType.Draft, DocDate);

                    foreach (DataRow InvoicesRow in Invoices.Rows)
                    {
                        Decimal BalanceDue = Convert.ToDecimal(InvoicesRow["OpenAmountFC"]);
                        Decimal AppliedFC = Convert.ToDecimal(InvoicesRow["AppliedFC"]);
                        string DocNum = InvoicesRow["DocNum"].ToString();
                        string InvoiceDocEntry = InvoicesRow["InvoiceDocEntry"].ToString();
                        string InstlmntID = InvoicesRow["InstlmntID"].ToString();

                        if (BalanceDue > AppliedFC)
                        {
                            decimal AppliedFCNEW = Math.Min(BalanceDue, AppliedFC + TrsfrSumFC);

                            string queryPDF2 = @"update PDF2 set
                                ""AppliedFC"" = " + AppliedFCNEW.ToString(CultureInfo.InvariantCulture) + @",
                                ""BfDcntSumF"" = " + AppliedFCNEW.ToString(CultureInfo.InvariantCulture) + @",
                                ""BfNetDcntF"" = " + AppliedFCNEW.ToString(CultureInfo.InvariantCulture) + @"
                                where ""DocNum"" = " + DocNum + @"
                                and ""DocEntry"" = " + InvoiceDocEntry + @"
                                and ""InstId"" = " + InstlmntID;

                            oRecordSetPDF2.DoQuery(queryPDF2);
                            TrsfrSumFC = TrsfrSumFC - AppliedFCNEW + AppliedFC;
                        }
                    }
                }

                string noDocSumField = "NoDocSumFC";

                decimal onAccountSum = TrsfrSumFC;
                if ((bpCurrency == "##" || String.IsNullOrEmpty(bpCurrency)) && DocRate > 0)
                {
                    noDocSumField = "NoDocSum";
                    onAccountSum = onAccountSum * DocRate;
                }

                //დარჩენილი თანხის OnAccount-ზე გაშვება
                if (onAccountSum != 0)
                {
                    query = @"update OPDF 
                        set 
                        " +
                        $"\"{noDocSumField}\" = " + onAccountSum.ToString(CultureInfo.InvariantCulture) + @",
                        ""PayNoDoc"" = 'Y'
                        where ""DocEntry"" = " + DocEntry + @"";

                    oRecordSet.DoQuery(query);
                }
                else
                {
                    query = @"update OPDF 
                        set 
                            " +
                        $"\"{noDocSumField}\" = " + "0" + @",
                        ""PayNoDoc"" = 'N'
                        where ""DocEntry"" = " + DocEntry + @"";

                    oRecordSet.DoQuery(query);
                }
            }

            Marshal.ReleaseComObject(oRecordSet);
            oRecordSet = null;

            Marshal.ReleaseComObject(oRecordSetPDF2);
            oRecordSetPDF2 = null;

            oFormDate.Close();
            oForm.Close();

            SAPbouiCOM.Form oJournalForm = Program.uiApp.OpenForm((SAPbouiCOM.BoFormObjectEnum)140, "", DocEntry);

        }

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

            if (formExist == true)
            {
                if (newForm == true)
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

        public static void fillAmountTaxes(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            oForm.Freeze(true);
            NumberFormatInfo Nfi = new NumberFormatInfo() { NumberDecimalSeparator = "." };
            try
            {
                double AmountPr = 0;

                SAPbouiCOM.Item oItemPrTx = oForm.Items.Item("AmtPrTxE");
                oItemPrTx.Enabled = true;

                string liablePrTx = oForm.DataSources.DBDataSources.Item("OVPM").GetValue("U_liablePrTx", 0).Trim();

                double profitTaxRate = Convert.ToDouble(ProfitTax.GetProfitTaxRate());

                SAPbouiCOM.Item oItemNoDocSum = oForm.Items.Item("13");
                SAPbouiCOM.EditText oEditNoDocSum = ((SAPbouiCOM.EditText)(oItemNoDocSum.Specific));
                string noDocSumTx = oEditNoDocSum.Value;

                double noDocSum = Convert.ToDouble(noDocSumTx, CultureInfo.InvariantCulture);
                if (liablePrTx == "Y" && noDocSum > 0)
                {
                    AmountPr = AmountPr + Convert.ToDouble(CommonFunctions.roundAmountByGeneralSettings(Convert.ToDecimal(noDocSum / (100 - profitTaxRate) * profitTaxRate), "Sum"));
                }

                SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("20").Specific));
                for (int i = 1; i < oMatrix.RowCount + 1; i++)
                {
                    string Payment = FormsB1.cleanStringOfNonDigits(oMatrix.Columns.Item("24").Cells.Item(i).Specific.Value).ToString();
                    string DocType = oMatrix.Columns.Item("45").Cells.Item(i).Specific.Value.ToString();
                    string DocNum = oMatrix.Columns.Item("1").Cells.Item(i).Specific.Value.ToString();
                    string Selected = oMatrix.Columns.Item("10000127").Cells.Item(i).Specific.Caption;

                    if (Selected == "Y" && (DocType == "204" || DocType == "18"))
                    {
                        bool prTx = GetNonEconExpAP(Convert.ToInt16(DocNum), DocType);
                        if (prTx)
                        {
                            AmountPr = AmountPr + Convert.ToDouble(CommonFunctions.roundAmountByGeneralSettings(Convert.ToDecimal(Payment) / Convert.ToDecimal(100 - profitTaxRate) * Convert.ToDecimal(profitTaxRate), "Sum"));
                        }
                    }

                }                

                SAPbouiCOM.EditText oEditAmtPrTx = ((SAPbouiCOM.EditText)(oItemPrTx.Specific));

                oEditAmtPrTx.Value = FormsB1.ConvertDecimalToString(Convert.ToDecimal(AmountPr));

                oForm.Items.Item("26").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                oItemPrTx.Enabled = false;
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
            finally
            {
                GC.Collect();
                oForm.Update();
                oForm.Freeze(false);
            }
        }

        public static void fillPhysicalEntityTaxes(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            oForm.Freeze(true);

            try
            {
                SAPbouiCOM.DBDataSources docDBSources = oForm.DataSources.DBDataSources;

                string wtCode = oForm.DataSources.DBDataSources.Item("OCRD").GetValue("WtCode", 0);

                bool physicalEntityTax = (oForm.DataSources.DBDataSources.Item("OCRD").GetValue("WTLiable", 0) == "Y" &&
                                CommonFunctions.getValue("OWHT", "U_BDOSPhisTx", "WTCode", wtCode).ToString() == "Y");

                string errorTextCheck;
                string docDatestr = docDBSources.Item("OVPM").GetValue("DocDate", 0);
                if (physicalEntityTax && string.IsNullOrEmpty(docDatestr))
                {
                    errorText = BDOSResources.getTranslate("DocDate") + " " + BDOSResources.getTranslate("YouCantLeaveEmpty");
                    return;
                }

                DateTime DocDate = DateTime.ParseExact(docDatestr, "yyyyMMdd", CultureInfo.InvariantCulture);
                Dictionary<string, decimal> PhysicalEntityPensionRates = WithholdingTax.getPhysicalEntityPensionRates(DocDate, wtCode, out errorTextCheck);

                if (physicalEntityTax && !string.IsNullOrEmpty(errorTextCheck))
                {
                    errorText = errorTextCheck;
                    return;
                }

                decimal TotalPensPhAm = 0;
                decimal TotalWhtAmt = 0;
                decimal TotalPensCoAm = 0;

                decimal PensPhAm = 0;
                decimal WhtAmt = 0;
                decimal PensCoAm = 0;
                decimal GrossAmount = 0;
                decimal GrossAmountFC = 0;

                SAPbouiCOM.Item oItemTxVal = oForm.Items.Item("13");
                SAPbouiCOM.EditText oEditTxVal = ((SAPbouiCOM.EditText)(oItemTxVal.Specific));
                string amtTxVal = oEditTxVal.Value;

                if (physicalEntityTax)
                {
                    if (physicalEntityTax && amtTxVal != "" && docDBSources.Item("OVPM").GetValue("WtCode", 0).ToString().Trim() != "")
                    {
                        bool frgn = docDBSources.Item("OVPM").GetValue("DocCurr", 0).Trim() != CommonFunctions.getLocalCurrency();

                        GrossAmount = Convert.ToDecimal(amtTxVal, CultureInfo.InvariantCulture);

                        PensPhAm = CommonFunctions.roundAmountByGeneralSettings(GrossAmount * PhysicalEntityPensionRates["PensionWTaxRate"] / 100, "Sum");
                        WhtAmt = CommonFunctions.roundAmountByGeneralSettings((GrossAmount - PensPhAm) * PhysicalEntityPensionRates["WTRate"] / 100, "Sum");
                        PensCoAm = CommonFunctions.roundAmountByGeneralSettings(GrossAmount * PhysicalEntityPensionRates["PensionCoWTaxRate"] / 100, "Sum");

                        SAPbouiCOM.EditText oEditWTax = oForm.Items.Item("111").Specific;
                        oEditWTax.Value = FormsB1.ConvertDecimalToStringForEditboxStrings(PensPhAm + WhtAmt);

                        if (frgn)
                        {
                            GrossAmountFC = GrossAmount * Convert.ToDecimal(CommonFunctions.getChildOrDbDataSourceValue(docDBSources.Item("OVPM"), null, null, "DocRate", 0), CultureInfo.InvariantCulture);

                            TotalPensPhAm = CommonFunctions.roundAmountByGeneralSettings(GrossAmountFC * PhysicalEntityPensionRates["PensionWTaxRate"] / 100, "Sum");
                            TotalWhtAmt = CommonFunctions.roundAmountByGeneralSettings((GrossAmountFC - TotalPensPhAm) * PhysicalEntityPensionRates["WTRate"] / 100, "Sum");
                            TotalPensCoAm = CommonFunctions.roundAmountByGeneralSettings(GrossAmountFC * PhysicalEntityPensionRates["PensionCoWTaxRate"] / 100, "Sum");
                        }
                        else
                        {
                            TotalPensPhAm = PensPhAm;
                            TotalWhtAmt = WhtAmt;
                            TotalPensCoAm = PensCoAm;
                        }
                    }

                    SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("20").Specific));
                    SAPbouiCOM.Columns oColumns = oMatrix.Columns;
                    string DocType;
                    decimal wtaxInv;

                    for (int i = 1; i < oMatrix.RowCount + 1; i++)
                    {
                        if (oColumns.Item("10000127").Cells.Item(i).Specific.Caption == "Y")
                        {
                            docDatestr = oColumns.Item("21").Cells.Item(i).Specific.Value.ToString();
                            DocDate = DateTime.ParseExact(docDatestr, "yyyyMMdd", CultureInfo.InvariantCulture);
                            PhysicalEntityPensionRates = WithholdingTax.getPhysicalEntityPensionRates(DocDate, wtCode, out errorTextCheck);

                            DocType = oColumns.Item("45").Cells.Item(i).Specific.Value.ToString().Trim();
                            GrossAmount = Convert.ToDecimal(FormsB1.cleanStringOfNonDigits(oColumns.Item("24").Cells.Item(i).Specific.Value), CultureInfo.InvariantCulture);
                            wtaxInv = Convert.ToDecimal(FormsB1.cleanStringOfNonDigits(oColumns.Item("72").Cells.Item(i).Specific.Value), CultureInfo.InvariantCulture);

                            decimal CurRate = Convert.ToDecimal(FormsB1.cleanStringOfNonDigits(oColumns.Item("41").Cells.Item(i).Specific.Value), CultureInfo.InvariantCulture);
                            if (CurRate != 0)
                            {
                                GrossAmount = GrossAmount * CurRate;
                                wtaxInv = wtaxInv * CurRate;
                            }

                            if (DocType == "19")
                            {
                                GrossAmount = GrossAmount * (-1);
                                wtaxInv = wtaxInv * (-1);
                            }

                            if ((DocType == "18" || DocType == "19" || DocType == "204") && wtaxInv != 0 && PhysicalEntityPensionRates["PensionWTaxRate"] != 0)
                            {
                                PensPhAm = CommonFunctions.roundAmountByGeneralSettings(wtaxInv * 100 * PhysicalEntityPensionRates["PensionWTaxRate"] / (100 * PhysicalEntityPensionRates["PensionWTaxRate"] + PhysicalEntityPensionRates["WTRate"] * (100 - PhysicalEntityPensionRates["PensionWTaxRate"])), "Sum");
                                TotalPensPhAm = TotalPensPhAm + PensPhAm;
                                if (PensPhAm != 0)
                                {
                                    TotalWhtAmt = TotalWhtAmt + (Convert.ToDecimal(wtaxInv, CultureInfo.InvariantCulture) - PensPhAm);
                                }
                                TotalPensCoAm = TotalPensCoAm + CommonFunctions.roundAmountByGeneralSettings((GrossAmount + wtaxInv) * PhysicalEntityPensionRates["PensionCoWTaxRate"] / 100, "Sum");
                            }
                        }
                    }
                }
                SAPbouiCOM.EditText oEdit = oForm.Items.Item("BDOSWhtAmt").Specific;
                oEdit.Value = FormsB1.ConvertDecimalToString(TotalWhtAmt);

                oEdit = oForm.Items.Item("BDOSPnPhAm").Specific;
                oEdit.Value = FormsB1.ConvertDecimalToString(TotalPensPhAm);

                oEdit = oForm.Items.Item("BDOSPnCoAm").Specific;
                oEdit.Value = FormsB1.ConvertDecimalToString(TotalPensCoAm);

                oForm.Items.Item("26").Click(SAPbouiCOM.BoCellClickType.ct_Regular);

            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
            finally
            {
                oForm.Freeze(false);
                oForm.Update();
                GC.Collect();
            }
        }


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
                if (cashFlowRelevant == true && string.IsNullOrEmpty(cashFlowLineItemID))
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

                if (cashFlowRelevant == true)
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
                    if (newDoc == true)
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

        public static string createDocumentTransferToBPType(SAPbouiCOM.DataTable oDataTable, int i, out int docEntry, out int docNum, out string errorText)
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
                if (cashFlowRelevant == true && string.IsNullOrEmpty(cashFlowLineItemID))
                    errorText = errorText + BDOSResources.getTranslate("TheFollowingFieldIsMandatory") + " : \"" + BDOSResources.getTranslate("CashFlowLineItemID") + "\"! ";
                string partnerTaxCode = oDataTable.GetValue("PartnerTaxCode", i);
                SAPbobsCOM.Recordset oRecordSet = CommonFunctions.getBPBankInfo(partnerAccountNumber + partnerCurrency, partnerTaxCode, "S");
                if (oRecordSet == null)
                {
                    errorText = BDOSResources.getTranslate("CouldNotFindBusinessPartner") + "! " + BDOSResources.getTranslate("Account") + " \"" + partnerAccountNumber + currency + "\"";
                    if (string.IsNullOrEmpty(partnerTaxCode) == false)
                    {
                        errorText = errorText + ", " + BDOSResources.getTranslate("Tin") + " \"" + partnerTaxCode + "\"! ";
                    }
                    else errorText = errorText + "! ";
                }

                if (string.IsNullOrEmpty(errorText) == false)
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
                oPayments.DocTypte = SAPbobsCOM.BoRcptTypes.rSupplier;
                oPayments.DocDate = docDate;
                oPayments.TaxDate = docDate;

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
                            oPayments.DocRate = oSBOBob.GetCurrencyRate(BPCurrency, docDate).Fields.Item("CurrencyRate").Value;
                            docRate = Convert.ToDecimal(oPayments.DocRate);
                            transferSumLC = amount;
                            transferSumFC = amount / docRate;
                        }
                        else if (partnerCurrencySapCode == BPCurrency)
                        {
                            oPayments.DocCurrency = BPCurrency;
                            oPayments.LocalCurrency = SAPbobsCOM.BoYesNoEnum.tNO;
                            oPayments.DocRate = oSBOBob.GetCurrencyRate(partnerCurrencySapCode, docDate).Fields.Item("CurrencyRate").Value;
                            docRate = Convert.ToDecimal(oPayments.DocRate);
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
                            oPayments.DocRate = oSBOBob.GetCurrencyRate(partnerCurrencySapCode, docDate).Fields.Item("CurrencyRate").Value;
                            docRate = Convert.ToDecimal(oPayments.DocRate);
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

                if (cashFlowRelevant == true)
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
                    if (newDoc == true)
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
                if (cashFlowRelevant == true && string.IsNullOrEmpty(cashFlowLineItemID))
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

                if (cashFlowRelevant == true)
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
                    if (newDoc == true)
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
                if (cashFlowRelevant == true && string.IsNullOrEmpty(cashFlowLineItemID))
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

                if (cashFlowRelevant == true)
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
                    if (newDoc == true)
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
                if (cashFlowRelevant == true && string.IsNullOrEmpty(cashFlowLineItemID))
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

                if (cashFlowRelevant == true)
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
                    if (newDoc == true)
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


        //--------------------------------------------INTERNET BANK-------------------------------------------->
        /// <summary>იმპორტი ინტერნეტ ბანკში</summary>
        /// <param name="Program.oCompany"></param>
        /// <param name="oPaymentService"></param>
        /// <param name="docEntryList"></param>
        /// <param name="importBatchPaymentOrders"></param>
        /// <param name="batchName"></param>
        /// <param name="errorText"></param>
        /// <returns></returns>
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
                        if (importBatchPaymentOrders == true)//პაკეტური გადარიცხვა
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
                        else if (transferType == "TreasuryTransferPaymentOrderIo") //საბიუჯეტო გადარიცხვა
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
                        if (transferType == "TransferToNationalCurrencyPaymentOrderIo" || transferType == "TreasuryTransferPaymentOrderIo") //გადარიცხვა (ეროვნული ვალუტა) || სახაზინო
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
            int length = paymentList.Count(); //importBatchPaymentOrders == true ? 1 : orderResult.Length;
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
                else if (importBatchPaymentOrders == true && key == 0) //პაკეტური
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
                else if (importBatchPaymentOrders == true && key == 0) //პაკეტური
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

        public static string getQueryForImport(List<int> docEntryList, string account, string startDate, string endDate, string program, bool allDocuments)
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

        /// <summary>ინტერნეტ ბანკის სტატუსების განახლება</summary>
        /// <param name="Program.oCompany"></param>
        /// <param name="docEntry"></param>
        /// <param name="errorText"></param>
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

            string PaymentTable = ptype == PaymentType.Draft ? "PDF2" : "VPM2";
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
                ""VPM2"".""AppliedFC"",
	            ""VPM2"".""InvType"",
	            ""VPM2"".""DocEntry"" AS ""InvoiceEntry"",
                ""VPM2"".""InstId"",

	            ""VPM2"".""DocNum"" AS ""PaymentDocEntry"",
                 ""VPM2"".""DocLine""
	            FROM " + PaymentTable + @" AS ""VPM2"") AS ""OutgoingPaymentInvoices"" ON (""ObjType"" = ""OutgoingPaymentInvoices"".""InvType"" 
            AND ""DocEntry"" = ""InvoiceEntry"" AND  ""InstallmentID"" = ""InstId"")
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


        public enum PaymentType
        {
            Draft = 1,
            Payment = 2
        }

        //<--------------------------------------------INTERNET BANK--------------------------------------------
    }
}