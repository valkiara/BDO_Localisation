using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace BDO_Localisation_AddOn
{
    static partial class IncomingPayment
    {
        private static bool changeU_OutDoc = false;

        public static void createUserFields(out string errorText)
        {
            errorText = null;

            Dictionary<string, object> fieldskeysMap;

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "UseBlaAgRt");
            fieldskeysMap.Add("TableName", "ORCT");
            fieldskeysMap.Add("Description", "Use Blanket Agreement Rates");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 1);
            fieldskeysMap.Add("DefaultValue", "N");

            UDO.addUserTableFields(fieldskeysMap, out errorText);
            GC.Collect();
            /*Dictionary<string, object> fieldskeysMap;
            Dictionary<string, string> listValidValuesDict;

            fieldskeysMap = new Dictionary<string, object>();
            listValidValuesDict = new Dictionary<string, string>();
            listValidValuesDict.Add("empty", ""); //ცარიელი
            listValidValuesDict.Add("downloadedFromTheBank", "Downloaded From The Bank"); //ჩამოტვირთულია ბანკიდან

            fieldskeysMap.Add("Name", "status"); //სტატუსი
            fieldskeysMap.Add("TableName", "ORCT");
            fieldskeysMap.Add("Description", "Status");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);
            fieldskeysMap.Add("ValidValues", listValidValuesDict);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "opType"); //ოპერაციის ტიპი
            fieldskeysMap.Add("TableName", "ORCT");
            fieldskeysMap.Add("Description", "Operation Type");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //მიმღების ანგარიში
            fieldskeysMap.Add("Name", "creditAcct");
            fieldskeysMap.Add("TableName", "ORCT");
            fieldskeysMap.Add("Description", "Credit Account");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //მიმღების ანგარიშის ვალუტა
            fieldskeysMap.Add("Name", "crdtActCur");
            fieldskeysMap.Add("TableName", "ORCT");
            fieldskeysMap.Add("Description", "Credit Account Currency");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //ტრანზაქციის ID
            fieldskeysMap.Add("Name", "paymentID");
            fieldskeysMap.Add("TableName", "ORCT");
            fieldskeysMap.Add("Description", "Payment ID");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            listValidValuesDict = new Dictionary<string, string>();
            listValidValuesDict.Add("SHA", "SHA"); //მიმღები მიიღებს შუამავალი ბანკის საკომისიოთი ნაკლებ თანხას (SHA)
            listValidValuesDict.Add("OUR", "OUR"); //მიმღები მიიღებს სრულ თანხას, გადარიცხვის საკომისიოს დაემატება 20USD/30EUR (OUR)

            fieldskeysMap.Add("Name", "chrgDtls"); //ხარჯი
            fieldskeysMap.Add("TableName", "ORCT");
            fieldskeysMap.Add("Description", "Charge Details");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);
            fieldskeysMap.Add("ValidValues", listValidValuesDict);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            listValidValuesDict = new Dictionary<string, string>();
            listValidValuesDict.Add("BULK", "BULK"); //BULK - სტანდარტული გადარიცხვა
            listValidValuesDict.Add("MT103", "MT103"); //MT103 ინდივიდუალური გადარიცხვა

            fieldskeysMap.Add("Name", "dsptchType"); //გადარიცხვის მეთოდი
            fieldskeysMap.Add("TableName", "ORCT");
            fieldskeysMap.Add("Description", "Dispatch Type");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);
            fieldskeysMap.Add("ValidValues", listValidValuesDict);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //დანიშნულება
            fieldskeysMap.Add("Name", "descrpt");
            fieldskeysMap.Add("TableName", "ORCT");
            fieldskeysMap.Add("Description", "Description");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 254);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //დამატებითი დანიშნულება
            fieldskeysMap.Add("Name", "addDescrpt");
            fieldskeysMap.Add("TableName", "ORCT");
            fieldskeysMap.Add("Description", "Additional Description");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 254);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //დოკუმენტის ნომერი ინტ. ბანკში
            fieldskeysMap.Add("Name", "docNumber");
            fieldskeysMap.Add("TableName", "ORCT");
            fieldskeysMap.Add("Description", "Document Number");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //ინტ. ბანკის ოპერაციის კოდი
            fieldskeysMap.Add("Name", "transCode");
            fieldskeysMap.Add("TableName", "ORCT");
            fieldskeysMap.Add("Description", "Transaction Code");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //ტრანზაქციის ID 2
            fieldskeysMap.Add("Name", "ePaymentID");
            fieldskeysMap.Add("TableName", "ORCT");
            fieldskeysMap.Add("Description", "External Payment ID");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //ინტ. ბანკის ოპერაციის კოდი 2
            fieldskeysMap.Add("Name", "opCode");
            fieldskeysMap.Add("TableName", "ORCT");
            fieldskeysMap.Add("Description", "Operation Code");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //Outgoing DocEntry //OVPM
            fieldskeysMap.Add("Name", "outDoc");
            fieldskeysMap.Add("TableName", "ORCT");
            fieldskeysMap.Add("Description", "Outgoing DocEntry");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 11);

            UDO.addUserTableFields( fieldskeysMap, out errorText);*/
        }

        public static void createFormItems(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

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

            SAPbobsCOM.Payments oIncomingPayments = null;
            SAPbobsCOM.ValidValues oValidValues = null;
            SAPbobsCOM.Fields oFields = null;

            oIncomingPayments = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments);
            oFields = oIncomingPayments.UserFields.Fields;

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
            formItems.Add("TableName", "ORCT");
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
                //oComboBox.ValidValues.Add("treasuryTransfer", BDOSResources.getTranslate("treasuryTransfer")); //სახაზინო გადარიცხვა
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
                //listValidValuesDict.Add("treasuryTransfer", BDOSResources.getTranslate("treasuryTransfer")); //სახაზინო გადარიცხვა
                listValidValuesDict.Add("other", BDOSResources.getTranslate("other")); //სხვა              

                formItems = new Dictionary<string, object>();
                itemName = "opTypeCB"; //10 characters
                formItems.Add("isDataSource", true);
                formItems.Add("DataSource", "DBDataSources");
                formItems.Add("TableName", "ORCT");
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
            formItems.Add("TableName", "ORCT");
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

            listValidValuesDict = CommonFunctions.getCurrencyListForValidValues();

            formItems = new Dictionary<string, object>();
            itemName = "crdActCuCB"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "ORCT");
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
            formItems.Add("TableName", "ORCT");
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
            formItems.Add("TableName", "ORCT");
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
            formItems.Add("TableName", "ORCT");
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
            formItems.Add("TableName", "ORCT");
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
            formItems.Add("TableName", "ORCT");
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

            top = top + height + 1;

            multiSelection = false;
            objectType = "46"; // oVendorPayments
            string uniqueID_lf_OutgoingPaymentCFL = "OutgoingPayment_CFL";
            FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_lf_OutgoingPaymentCFL);

            formItems = new Dictionary<string, object>();
            itemName = "outDocS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("OutgoingReference"));
            formItems.Add("LinkTo", "outDocE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "outDocE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "ORCT");
            formItems.Add("Alias", "U_outDoc");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e + width_e / 2 + 20);
            formItems.Add("Width", width_e / 2 - 20);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("ChooseFromListUID", uniqueID_lf_OutgoingPaymentCFL);
            formItems.Add("ChooseFromListAlias", "DocEntry");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "outDocLB"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            formItems.Add("Left", left_e + width_e / 2);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("LinkTo", "outDocE");
            formItems.Add("LinkedObjectType", objectType);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }


            // ------------------- use blanket agreement rate ranges

            // -------------------- Use blanket agreement rates-----------------

            int left = oForm.Items.Item("234000004").Left;
            height = oForm.Items.Item("234000004").Height;
            top = oForm.Items.Item("234000004").Top + height + 5;

            formItems = new Dictionary<string, object>();
            itemName = "UsBlaAgRtS"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "ORCT");
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
            formItems.Add("Enabled", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }
        }

        public static void comboSelect(SAPbouiCOM.Form oForm, SAPbouiCOM.ItemEvent pVal)
        {
            try
            {
                oForm.Freeze(true);

                if (!pVal.BeforeAction)
                {
                    if (pVal.ItemUID == "opTypeCB")
                    {
                        string opType = oForm.DataSources.DBDataSources.Item("ORCT").GetValue("U_opType", 0).Trim();
                        SAPbouiCOM.EditText oEditText;
                        SAPbouiCOM.ComboBox oComboBox;

                        if (opType == "transferToOwnAccount") //გადარიცხვა პირად ანგარიშზე
                        {
                            //oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("tresrCodeE").Specific;
                            //oEditText.Value = "";
                            oComboBox = (SAPbouiCOM.ComboBox)oForm.Items.Item("crdActCuCB").Specific;
                            oComboBox.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue);
                            oComboBox = (SAPbouiCOM.ComboBox)oForm.Items.Item("chrgDtlsCB").Specific;
                            oComboBox.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue);
                            oComboBox = (SAPbouiCOM.ComboBox)oForm.Items.Item("dsptTypeCB").Specific;
                            oComboBox.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue);
                        }
                        else if (opType == "currencyExchange") //კონვერტაცია
                        {
                            //oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("tresrCodeE").Specific;
                            //oEditText.Value = "";
                            oComboBox = (SAPbouiCOM.ComboBox)oForm.Items.Item("chrgDtlsCB").Specific;
                            oComboBox.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue);
                            oComboBox = (SAPbouiCOM.ComboBox)oForm.Items.Item("dsptTypeCB").Specific;
                            oComboBox.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue);
                        }
                    }
                }
            }
            catch
            {
                throw;
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        public static DataTable createAdditionalEntries(Dictionary<string, object> oDictionary, string DocCurrency, decimal DocRate, DateTime DocDate, out string errorText)
        {
            errorText = null;
            SAPbobsCOM.Payments oPayments = null;
            oPayments = (SAPbobsCOM.Payments)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments);
            DataTable jeLines = null;
            DataRow jeLinesRow = null;

            try
            {
                jeLines = JournalEntry.JournalEntryTable();

                string docType;
                string accountCodeIN;
                decimal trsfrSumIN;
                decimal trsfrSumINFC;
                string accountCodeOUT;
                decimal trsfrSumOUT;
                string outDocIN;
                decimal trsfrSum;
                decimal DocRateOUT;
                string DocCurrencyOUT;
                string U_PrjCode;

                accountCodeIN = oDictionary["CardCode"].ToString();
                trsfrSumIN = Convert.ToDecimal(oDictionary["TrsfrSum"], NumberFormatInfo.InvariantInfo);
                outDocIN = oDictionary["U_outDoc"].ToString();
                docType = oDictionary["DocType"].ToString();
                string U_crdtActCur = oDictionary["U_crdtActCur"].ToString();
                U_PrjCode = oDictionary["PrjCode"].ToString();

                if (docType == "A" && !string.IsNullOrEmpty(outDocIN))
                {
                    if (oPayments.GetByKey(Convert.ToInt32(outDocIN)))
                    {
                        accountCodeOUT = oPayments.CardCode;
                        DocCurrencyOUT = oPayments.DocCurrency;
                        string DocCurrencyForConvert = oPayments.DocCurrency;
                        DocRateOUT = (decimal)oPayments.DocRate;
                        trsfrSumOUT = Convert.ToDecimal(oPayments.TransferSum);

                        trsfrSum = trsfrSumIN - trsfrSumOUT;

                        string localcurrency = CommonFunctions.getLocalCurrency();
                        string year = DocDate.Year.ToString();

                        string accountCodeGain = CommonFunctions.getPeriodsCategory("GLGainXdif", year);
                        string accountCodeLoss = CommonFunctions.getPeriodsCategory("GLLossXdif", year);

                        if (string.IsNullOrEmpty(accountCodeLoss) || string.IsNullOrEmpty(accountCodeLoss))
                            return jeLines;

                        int J = 0;

                        string DocCurrencyIN = DocCurrency == localcurrency ? "" : DocCurrency;
                        trsfrSumINFC = DocCurrencyIN == "" ? 0 : trsfrSumIN / DocRate;

                        jeLinesRow = jeLines.Rows.Add(J);
                        jeLinesRow["AccountCode"] = accountCodeIN; //Debit
                        jeLinesRow["ShortName"] = accountCodeIN;
                        jeLinesRow["ContraAccount"] = accountCodeIN;
                        jeLinesRow["Credit"] = 0;
                        jeLinesRow["Debit"] = Convert.ToDouble(trsfrSumIN);
                        jeLinesRow["FCDebit"] = Convert.ToDouble(trsfrSumINFC);
                        jeLinesRow["FCCurrency"] = DocCurrencyIN;
                        jeLinesRow["ProjectCode"] = U_PrjCode;
                        J++;

                        decimal trsfrSumOUTFC = 0;
                        if (DocCurrency == localcurrency && DocCurrencyOUT != localcurrency)
                            trsfrSumOUTFC = DocRateOUT == 0 ? 0 : trsfrSumOUT / DocRateOUT;
                        else
                            DocCurrencyOUT = "";

                        jeLinesRow = jeLines.Rows.Add(J);
                        jeLinesRow["AccountCode"] = accountCodeOUT; //Credit
                        jeLinesRow["ShortName"] = accountCodeOUT;
                        jeLinesRow["ContraAccount"] = accountCodeOUT;
                        jeLinesRow["Credit"] = Convert.ToDouble(trsfrSumOUT);
                        jeLinesRow["Debit"] = 0;
                        jeLinesRow["FCCredit"] = Convert.ToDouble(trsfrSumOUTFC);
                        jeLinesRow["FCCurrency"] = DocCurrencyOUT;
                        jeLinesRow["ProjectCode"] = U_PrjCode;
                        J++;

                        if (trsfrSum > 0)
                        {
                            jeLinesRow = jeLines.Rows.Add(J);
                            jeLinesRow["AccountCode"] = accountCodeGain; //Credit
                            jeLinesRow["ShortName"] = accountCodeGain;
                            jeLinesRow["ContraAccount"] = accountCodeGain;
                            jeLinesRow["Credit"] = Convert.ToDouble(trsfrSum);
                            jeLinesRow["Debit"] = 0;
                            jeLinesRow["ProjectCode"] = U_PrjCode;

                            J++;
                        }
                        else if (trsfrSum < 0)
                        {
                            jeLinesRow = jeLines.Rows.Add(J);
                            jeLinesRow["AccountCode"] = accountCodeLoss; //Debit
                            jeLinesRow["ShortName"] = accountCodeLoss;
                            jeLinesRow["ContraAccount"] = accountCodeLoss;
                            jeLinesRow["Credit"] = 0;
                            jeLinesRow["Debit"] = Convert.ToDouble(trsfrSum * (-1));
                            jeLinesRow["ProjectCode"] = U_PrjCode;

                            J++;
                        }



                        if (U_crdtActCur != DocCurrencyIN && DocCurrencyIN != "" && U_crdtActCur != "GEL")
                        {
                            trsfrSumOUTFC = DocRateOUT == 0 ? 0 : trsfrSumOUT / DocRateOUT;

                            jeLinesRow = jeLines.Rows.Add(J);
                            jeLinesRow["AccountCode"] = accountCodeOUT; //Credit
                            jeLinesRow["ShortName"] = accountCodeOUT;
                            jeLinesRow["ContraAccount"] = accountCodeOUT;
                            jeLinesRow["Credit"] = 0;
                            jeLinesRow["Debit"] = Convert.ToDouble(trsfrSumOUT);
                            jeLinesRow["ProjectCode"] = U_PrjCode;
                            J++;

                            jeLinesRow = jeLines.Rows.Add(J);
                            jeLinesRow["AccountCode"] = accountCodeOUT; //Credit
                            jeLinesRow["ShortName"] = accountCodeOUT;
                            jeLinesRow["ContraAccount"] = accountCodeOUT;
                            jeLinesRow["Credit"] = Convert.ToDouble(trsfrSumOUT); ;
                            jeLinesRow["FCCredit"] = Convert.ToDouble(trsfrSumOUTFC);
                            jeLinesRow["Debit"] = 0;
                            jeLinesRow["FCDebit"] = 0;
                            jeLinesRow["FCCurrency"] = DocCurrencyForConvert;
                            jeLinesRow["ProjectCode"] = U_PrjCode;
                            J++;

                        }
                    }
                }
                return jeLines;
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
                return jeLines;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oPayments);
                oPayments = null;
            }
        }

        public static void JrnEntry(string DocEntry, string DocNum, DateTime DocDate, DataTable JrnLinesDT, out string errorText)
        {
            errorText = null;

            try
            {
                JournalEntry.JrnEntry(DocEntry, "24", "Incoming Payment: " + DocNum, DocDate, JrnLinesDT, out errorText);
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
        }

        public static void createJrnEntry(SAPbouiCOM.Form oForm, string DocEntry, SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool bubbleEvent, out string errorText)
        {
            bubbleEvent = true;
            errorText = null;

            Dictionary<string, object> oDictionary = new Dictionary<string, object>();
            DateTime DocDate = new DateTime();

            string canceled;

            if (oForm != null)
            {
                SAPbouiCOM.DBDataSource oDBDataSource = oForm.DataSources.DBDataSources.Item("ORCT");
                canceled = oDBDataSource.GetValue("CANCELED", 0).Trim();

                if (oDBDataSource.GetValue("DocType", 0).Trim() == "A")
                {
                    oDictionary.Add("CANCELED", canceled);
                    oDictionary.Add("DocEntry", oDBDataSource.GetValue("DocEntry", 0).Trim());
                    oDictionary.Add("DocNum", oDBDataSource.GetValue("DocNum", 0).Trim());
                    oDictionary.Add("DocRate", oDBDataSource.GetValue("DocRate", 0).Trim());
                    oDictionary.Add("DocCurr", oDBDataSource.GetValue("DocCurr", 0).Trim());
                    oDictionary.Add("DocDate", oDBDataSource.GetValue("DocDate", 0));
                    oDictionary.Add("CardCode", oForm.DataSources.DBDataSources.Item("RCT4").GetValue("AcctCode", 0).Trim());
                    oDictionary.Add("TrsfrSum", oDBDataSource.GetValue("TrsfrSum", 0));
                    oDictionary.Add("U_outDoc", oDBDataSource.GetValue("U_outDoc", 0).Trim());
                    oDictionary.Add("U_crdtActCur", oDBDataSource.GetValue("U_crdtActCur", 0).Trim());
                    oDictionary.Add("DocType", oDBDataSource.GetValue("DocType", 0).Trim());
                    oDictionary.Add("PrjCode", oDBDataSource.GetValue("PrjCode", 0).Trim());
                    DocDate = DateTime.ParseExact(oDictionary["DocDate"].ToString(), "yyyyMMdd", CultureInfo.InvariantCulture);
                }
            }
            else
            {
                SAPbobsCOM.Payments oIncomingPayment = null;
                oIncomingPayment = (SAPbobsCOM.Payments)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments);
                if (oIncomingPayment.GetByKey(Convert.ToInt32(DocEntry)))
                {
                    canceled = oIncomingPayment.Cancelled == SAPbobsCOM.BoYesNoEnum.tNO ? "N" : "Y";
                    if (oIncomingPayment.DocType == SAPbobsCOM.BoRcptTypes.rAccount)
                    {
                        oDictionary.Add("CANCELED", canceled);
                        oDictionary.Add("DocEntry", oIncomingPayment.DocEntry);
                        oDictionary.Add("DocNum", oIncomingPayment.DocNum);
                        oDictionary.Add("DocRate", oIncomingPayment.DocRate);
                        oDictionary.Add("DocDate", oIncomingPayment.DocDate);
                        oDictionary.Add("DocCurr", oIncomingPayment.DocCurrency);
                        oDictionary.Add("CardCode", oIncomingPayment.CardCode);
                        oDictionary.Add("TrsfrSum", oIncomingPayment.TransferSum);
                        oDictionary.Add("U_outDoc", oIncomingPayment.UserFields.Fields.Item("U_outDoc").Value.Trim());
                        oDictionary.Add("U_crdtActCur", oIncomingPayment.UserFields.Fields.Item("U_crdtActCur").Value.Trim());
                        oDictionary.Add("DocType", "A");
                        oDictionary.Add("PrjCode", oIncomingPayment.ProjectCode);
                        DocDate = Convert.ToDateTime(oDictionary["DocDate"]);
                    }
                }
                else
                {
                    Marshal.FinalReleaseComObject(oIncomingPayment);
                    bubbleEvent = false;
                    return;
                }
                Marshal.FinalReleaseComObject(oIncomingPayment);
            }

            //დოკუმენტის გატარების დროს გატარდეს ბუღლტრული გატარება
            if (canceled == "N" && oDictionary.Count() > 0)
            {
                DocEntry = oDictionary["DocEntry"].ToString();
                string DocCurrency = oDictionary["DocCurr"].ToString();
                decimal DocRate = Convert.ToDecimal(FormsB1.cleanStringOfNonDigits(oDictionary["DocRate"].ToString()));
                string DocNum = oDictionary["DocNum"].ToString();

                CommonFunctions.StartTransaction();

                Program.JrnLinesGlobal = new DataTable();
                DataTable JrnLinesDT = createAdditionalEntries(oDictionary, DocCurrency, DocRate, DocDate, out errorText);
                if (errorText != null)
                {
                    bubbleEvent = false;
                    return;
                }
                if (JrnLinesDT == null)
                    return;
                JrnEntry(DocEntry, DocNum, DocDate, JrnLinesDT, out errorText);
                if (errorText != null)
                {
                    bubbleEvent = false;
                    return;
                }
                else
                {
                    if (BusinessObjectInfo == null || !BusinessObjectInfo.ActionSuccess)
                        Program.JrnLinesGlobal = JrnLinesDT;
                }

                //თუ დოკუმენტი გატარდა, მერე ვაკეთებს ბუღალტრულ გატარებას
                if (BusinessObjectInfo == null || (BusinessObjectInfo.ActionSuccess && !BusinessObjectInfo.BeforeAction))
                    CommonFunctions.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                else
                    CommonFunctions.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
            }
        }

        public static void setVisibleFormItems(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.Item oItem;
            oForm.Freeze(true);

            try
            {
                oForm.Items.Item("26").Click();

                oForm.Items.Item("statusCB").Enabled = false;
                string docEntry = oForm.DataSources.DBDataSources.Item("ORCT").GetValue("DocEntry", 0).Trim();
                string opType = oForm.DataSources.DBDataSources.Item("ORCT").GetValue("U_opType", 0).Trim();
                string docType = oForm.DataSources.DBDataSources.Item("ORCT").GetValue("DocType", 0).Trim();
                string outDoc = oForm.DataSources.DBDataSources.Item("ORCT").GetValue("U_outDoc", 0).Trim();

                //string PayNoDoc = oForm.DataSources.DBDataSources.Item("ORCT").GetValue("PayNoDoc", 0).Trim();
                //string CardCode = oForm.DataSources.DBDataSources.Item("ORCT").GetValue("CardCode", 0).Trim();

                oForm.Items.Item("opTypeCB").Enabled = string.IsNullOrEmpty(docEntry);

                if (string.IsNullOrEmpty(opType) || opType == "transferToOwnAccount" || opType == "currencyExchange" || opType == "treasuryTransfer")
                {
                    try
                    {
                        SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("71").Specific;

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

                //Dictionary<string, string> dataForTransferType = getDataForTransferType( oForm);
                //string transferType = getTransferType( dataForTransferType, out errorText);

                oForm.Items.Item("outDocS").Visible = oForm.Items.Item("outDocE").Visible = opType == "currencyExchange" && docType == "A";

                if (docType == "A")
                {
                    oItem = oForm.Items.Item("opTypeS");
                    oItem.Visible = true;
                    oItem = oForm.Items.Item("opTypeCB");
                    oItem.Visible = true;

                    if (!string.IsNullOrEmpty(outDoc))
                    {
                        oItem = oForm.Items.Item("outDocE");
                        oItem.Enabled = false;
                    }
                    else
                    {
                        oItem = oForm.Items.Item("outDocE");
                        oItem.Enabled = true;
                        SAPbouiCOM.EditText oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("outDocE").Specific;
                        oEditText.ChooseFromListUID = "OutgoingPayment_CFL";
                        oEditText.ChooseFromListAlias = "DocEntry";
                    }
                    //oItem = oForm.Items.Item("chrgDtlsS");
                    //oItem.Visible = false;
                    //oItem = oForm.Items.Item("chrgDtlsCB");
                    //oItem.Visible = false;
                    //oItem = oForm.Items.Item("dsptTypeS");
                    //oItem.Visible = false;
                    //oItem = oForm.Items.Item("dsptTypeCB");
                    //oItem.Visible = false;

                    if (opType == "transferToOwnAccount") //გადარიცხვა პირად ანგარიშზე
                    {
                        oItem = oForm.Items.Item("creditActS");
                        oItem.Visible = true;
                        oItem = oForm.Items.Item("creditActE");
                        oItem.Visible = true;
                        oItem = oForm.Items.Item("crdActCuCB");
                        oItem.Visible = false;

                        //if (transferType == "TransferToOtherBankForeignCurrencyPaymentOrderIo" || transferType == "TransferToForeignCurrencyPaymentOrderIo")
                        //{
                        //    oItem = oForm.Items.Item("chrgDtlsS");
                        //    oItem.Visible = true;
                        //    oItem = oForm.Items.Item("chrgDtlsCB");
                        //    oItem.Visible = true;
                        //}
                        //else if (transferType == "TransferToNationalCurrencyPaymentOrderIo")
                        //{
                        //    oItem = oForm.Items.Item("dsptTypeS");
                        //    oItem.Visible = true;
                        //    oItem = oForm.Items.Item("dsptTypeCB");
                        //    oItem.Visible = true;
                        //}
                    }
                    else if (opType == "currencyExchange") //კონვერტაცია
                    {
                        oItem = oForm.Items.Item("creditActS");
                        oItem.Visible = true;
                        oItem = oForm.Items.Item("creditActE");
                        oItem.Visible = true;
                        oItem = oForm.Items.Item("crdActCuCB");
                        oItem.Visible = true;
                        //oItem = oForm.Items.Item("tresrCodeS");
                        //oItem.Visible = false;
                        //oItem = oForm.Items.Item("tresrCodeE");
                        //oItem.Visible = false;

                        //if (transferType == "TransferToOtherBankForeignCurrencyPaymentOrderIo" || transferType == "TransferToForeignCurrencyPaymentOrderIo")
                        //{
                        //    oItem = oForm.Items.Item("chrgDtlsS");
                        //    oItem.Visible = true;
                        //    oItem = oForm.Items.Item("chrgDtlsCB");
                        //    oItem.Visible = true;
                        //    oItem = oForm.Items.Item("dsptTypeS");
                        //    oItem.Visible = false;
                        //    oItem = oForm.Items.Item("dsptTypeCB");
                        //    oItem.Visible = false;
                        //}
                        //else if (transferType == "TransferToNationalCurrencyPaymentOrderIo")
                        //{
                        //    oItem = oForm.Items.Item("chrgDtlsS");
                        //    oItem.Visible = false;
                        //    oItem = oForm.Items.Item("chrgDtlsCB");
                        //    oItem.Visible = false;
                        //    oItem = oForm.Items.Item("dsptTypeS");
                        //    oItem.Visible = true;
                        //    oItem = oForm.Items.Item("dsptTypeCB");
                        //    oItem.Visible = true;
                        //}
                    }
                    else
                    {
                        oItem = oForm.Items.Item("creditActS");
                        oItem.Visible = false;
                        oItem = oForm.Items.Item("creditActE");
                        oItem.Visible = false;
                        oItem = oForm.Items.Item("crdActCuCB");
                        oItem.Visible = false;
                        //oItem = oForm.Items.Item("chrgDtlsS");
                        //oItem.Visible = false;
                        //oItem = oForm.Items.Item("chrgDtlsCB");
                        //oItem.Visible = false;
                        //oItem = oForm.Items.Item("dsptTypeS");
                        //oItem.Visible = false;
                        //oItem = oForm.Items.Item("dsptTypeCB");
                        //oItem.Visible = false;
                    }
                }
                //else if (transferType == "TransferToOtherBankForeignCurrencyPaymentOrderIo" || transferType == "TransferToForeignCurrencyPaymentOrderIo")
                //{
                //    oItem = oForm.Items.Item("chrgDtlsS");
                //    oItem.Visible = true;
                //    oItem = oForm.Items.Item("chrgDtlsCB");
                //    oItem.Visible = true;
                //    oItem = oForm.Items.Item("dsptTypeS");
                //    oItem.Visible = false;
                //    oItem = oForm.Items.Item("dsptTypeCB");
                //    oItem.Visible = false;
                //}
                //else if (transferType == "TransferToNationalCurrencyPaymentOrderIo")
                //{
                //    oItem = oForm.Items.Item("chrgDtlsS");
                //    oItem.Visible = false;
                //    oItem = oForm.Items.Item("chrgDtlsCB");
                //    oItem.Visible = false;
                //    oItem = oForm.Items.Item("dsptTypeS");
                //    oItem.Visible = true;
                //    oItem = oForm.Items.Item("dsptTypeCB");
                //    oItem.Visible = true;
                //}
                else
                {
                    oItem = oForm.Items.Item("opTypeS");
                    oItem.Visible = false;
                    oItem = oForm.Items.Item("opTypeCB");
                    oItem.Visible = false;
                    oItem = oForm.Items.Item("outDocS");
                    oItem.Visible = false;
                    oItem = oForm.Items.Item("outDocE");
                    oItem.Visible = false;
                    oItem = oForm.Items.Item("creditActS");
                    oItem.Visible = false;
                    oItem = oForm.Items.Item("creditActE");
                    oItem.Visible = false;
                    oItem = oForm.Items.Item("crdActCuCB");
                    oItem.Visible = false;
                    //oItem = oForm.Items.Item("chrgDtlsS");
                    //oItem.Visible = false;
                    //oItem = oForm.Items.Item("chrgDtlsCB");
                    //oItem.Visible = false;
                    //oItem = oForm.Items.Item("dsptTypeS");
                    //oItem.Visible = false;
                    //oItem = oForm.Items.Item("dsptTypeCB");
                    //oItem.Visible = false;
                }
            }
            catch
            {
                throw;
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        public static void chooseFromList(SAPbouiCOM.Form oForm, SAPbouiCOM.ItemEvent pVal, SAPbouiCOM.IChooseFromListEvent oCFLEvento, ref bool bubbleEvent)
        {
            NumberFormatInfo Nfi = new NumberFormatInfo() { NumberDecimalSeparator = "." };

            try
            {
                oForm.Freeze(true);

                if (pVal.BeforeAction)
                {
                    SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID);
                    if (oCFLEvento.ChooseFromListUID == "HouseBankAccount_CFL")
                    {

                    }
                    else if (oCFLEvento.ChooseFromListUID == "OutgoingPayment_CFL")
                    {
                        string opType = oForm.DataSources.DBDataSources.Item("ORCT").GetValue("U_opType", 0).Trim();

                        SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();
                        SAPbouiCOM.Condition oCon;

                        List<string> outDocList = getOutgoingPaymentsDocumentList(opType);
                        int docCount = outDocList.Count;
                        for (int i = 0; i < docCount; i++)
                        {
                            oCon = oCons.Add();
                            oCon.Alias = "DocEntry";
                            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                            oCon.CondVal = outDocList[i];
                            oCon.Relationship = (i == docCount - 1) ? SAPbouiCOM.BoConditionRelationship.cr_NONE : SAPbouiCOM.BoConditionRelationship.cr_OR;
                        }

                        oCFL.SetConditions(oCons);
                    }
                    else if (oCFLEvento.ChooseFromListUID == "1") //Blanket Agreement
                    {
                        string project = oForm.DataSources.DBDataSources.Item("ORCT").GetValue("PrjCode", 0);
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
                    SAPbouiCOM.DataTable oDataTable = oCFLEvento.SelectedObjects;

                    if (oDataTable != null)
                    {
                        if (oCFLEvento.ChooseFromListUID == "HouseBankAccount_CFL")
                        {
                            string account = Convert.ToString(oDataTable.GetValue("Account", 0));
                            string currency;
                            LanguageUtils.IgnoreErrors<string>(() => oForm.Items.Item("creditActE").Specific.Value = account);
                            try
                            {
                                CommonFunctions.accountParse(account, out currency);
                                oForm.Items.Item("crdActCuCB").Specific.Select(currency, SAPbouiCOM.BoSearchKey.psk_ByValue);
                            }
                            catch { }
                        }
                        if (oCFLEvento.ChooseFromListUID == "OutgoingPayment_CFL")
                        {
                            string docEntry = Convert.ToString(oDataTable.GetValue("DocEntry", 0));
                            if (!string.IsNullOrEmpty(docEntry))
                            {
                                try
                                {
                                    LanguageUtils.IgnoreErrors<string>(() => oForm.Items.Item("outDocE").Specific.Value = docEntry);
                                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                                    {
                                        changeU_OutDoc = true;
                                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                                    }
                                }
                                catch
                                {
                                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                                    {
                                        changeU_OutDoc = true;
                                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                                    }
                                }
                            }
                        }
                        else if (oCFLEvento.ChooseFromListUID == "1") //Blanket Agreement
                        {
                            var agrNo = Convert.ToString(oDataTable.GetValue("AbsID", 0));
                            var prjCode = Convert.ToString(oDataTable.GetValue("Project", 0));
                            var bpCurr = Convert.ToString(oDataTable.GetValue("BPCurr", 0));

                            LanguageUtils.IgnoreErrors<string>(() => oForm.Items.Item("95").Specific.Value = prjCode);

                            FilterInvoiceMatrix(oForm, agrNo, prjCode);

                            oForm.Items.Item("26").Click(); //Remark

                            oForm.Items.Item("95").Enabled = string.IsNullOrEmpty(prjCode);
                            oForm.Items.Item("234000005").Enabled = string.IsNullOrEmpty(agrNo);

                            SetUsBlaAgRtSAvailability(oForm, !string.IsNullOrEmpty(agrNo) && bpCurr != Program.LocalCurrency);
                        }

                        else if (oCFLEvento.ChooseFromListUID == "23") //Project
                        {
                            var prjCode = Convert.ToString(oDataTable.GetValue("PrjCode", 0));
                            
                            oForm.Items.Item("234000005").Enabled = true;
                            LanguageUtils.IgnoreErrors<string>(() => oForm.Items.Item("234000005").Specific.Value = string.Empty);

                            FilterInvoiceMatrix(oForm, null, prjCode);

                            oForm.Items.Item("26").Click(); //Remark
                            oForm.Items.Item("95").Enabled = string.IsNullOrEmpty(prjCode);

                            SetUsBlaAgRtSAvailability(oForm);
                        }
                    }
                    setVisibleFormItems(oForm);
                }
            }
            catch
            {
                throw;
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        public static List<string> getOutgoingPaymentsDocumentList(string opType)
        {
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            List<string> baseDocList = new List<string>();

            try
            {
                string query = @"SELECT
	                 ""PMNT"".""DocEntry"" 
                FROM ""OVPM"" AS ""PMNT"" 
                WHERE ""PMNT"".""Canceled"" = 'N' 
                AND ""PMNT"".""DocType"" = 'A' 
                AND ""PMNT"".""U_opType"" = '" + opType + @"' 
                AND CAST(""PMNT"".""DocEntry""AS NVARCHAR) NOT IN (SELECT
                	 ""ORCT"".""U_outDoc"" 
                	FROM ""ORCT"" AS ""ORCT"" 
                	WHERE ""ORCT"".""Canceled"" = 'N' 
                	AND ""ORCT"".""U_outDoc"" <> '' 
                	AND ""ORCT"".""U_outDoc"" IS NOT NULL)";

                oRecordSet.DoQuery(query);

                while (!oRecordSet.EoF)
                {
                    baseDocList.Add(oRecordSet.Fields.Item("DocEntry").Value.ToString());
                    oRecordSet.MoveNext();
                }
            }
            catch
            {
                return baseDocList;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                oRecordSet = null;
            }
            return baseDocList;
        }

        public static void resizeForm(SAPbouiCOM.Form oForm)
        {
            try
            {
                oForm.Freeze(true);

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

                top = top + height + 1;

                oItem = oForm.Items.Item("outDocS");
                oItem.Top = top;
                oItem = oForm.Items.Item("outDocE");
                oItem.Top = top;
                oItem = oForm.Items.Item("outDocLB");
                oItem.Top = top;
            }
            catch
            {
                throw;
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        public static void uiApp_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(BusinessObjectInfo.FormTypeEx, Program.currentFormCount);

            if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD)
            {
                if (!BusinessObjectInfo.BeforeAction && !BusinessObjectInfo.ActionSuccess)
                {
                    BubbleEvent = false;
                }

                //if (BusinessObjectInfo.ActionSuccess && !BusinessObjectInfo.BeforeAction)
                //{
                //    SAPbouiCOM.DBDataSource DocDBSourceOCRD = oForm.DataSources.DBDataSources.Item(0);

                //    string DocEntry = DocDBSourceOCRD.GetValue("DocEntry", 0).Trim();
                //    JrnEntry( DocEntry, out errorText);                  
                //    if (errorText != null)
                //    {
                //        Program.uiApp.MessageBox(errorText);
                //        BubbleEvent = false;
                //    }
                //}

                if (BusinessObjectInfo.ActionSuccess != BusinessObjectInfo.BeforeAction)
                {
                    createJrnEntry(oForm, null, BusinessObjectInfo, out BubbleEvent, out errorText);
                    if (!string.IsNullOrEmpty(errorText))
                    {
                        Program.uiApp.MessageBox(errorText);
                        BubbleEvent = false;
                    }
                }
            }

            else if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE)
            {
                if (BusinessObjectInfo.BeforeAction)
                {
                    if (Program.cancellationTrans && Program.canceledDocEntry != 0)
                    {
                        SAPbouiCOM.DBDataSource oDBDataSource = oForm.DataSources.DBDataSources.Item("ORCT");
                        string docEntry = oDBDataSource.GetValue("DocEntry", 0).Trim();
                        if (isARDownPaymentVATAccrual(Convert.ToInt32(docEntry)))
                        {
                            Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("YouCantCancelDocumentBecauseOfARDownPaymentVATAccrual"), SAPbouiCOM.BoMessageTime.bmt_Short);
                            BubbleEvent = false;
                            return;
                        }
                    }
                }

                if (BusinessObjectInfo.ActionSuccess && !BusinessObjectInfo.BeforeAction)
                {
                    if (changeU_OutDoc)
                    {
                        createJrnEntry(oForm, null, BusinessObjectInfo, out BubbleEvent, out errorText);
                        if (!string.IsNullOrEmpty(errorText))
                        {
                            Program.uiApp.MessageBox(errorText);
                            BubbleEvent = false;
                        }
                        else
                            changeU_OutDoc = false;
                    }
                }
            }

            //შემოწმება
            //if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD)
            //{
            //    if (BusinessObjectInfo.BeforeAction )
            //    {
            //        checkFillDoc( oForm, out errorText);
            //        if (errorText != null)
            //        {
            //            Program.uiApp.SetStatusBarMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, false);
            //            BubbleEvent = false;
            //        }
            //    }
            //}

            //if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE)
            //{
            //    if (BusinessObjectInfo.BeforeAction )
            //    {
            //        checkFillDoc( oForm, out errorText);
            //        if (errorText != null)
            //        {
            //            Program.uiApp.SetStatusBarMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, false);
            //            //Program.uiApp.MessageBox(errorText);
            //            BubbleEvent = false;
            //        }
            //    }
            //}

            if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE && !BusinessObjectInfo.BeforeAction && BusinessObjectInfo.ActionSuccess)
            {
                if (Program.cancellationTrans && Program.canceledDocEntry != 0)
                {
                    cancellation(oForm, Program.canceledDocEntry);
                    Program.canceledDocEntry = 0;
                }
            }

            else if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD && !BusinessObjectInfo.BeforeAction)
            {
                if (BusinessObjectInfo.FormTypeEx == "170")
                    setVisibleFormItems(oForm);
                changeU_OutDoc = false;
            }
        }

        public static void uiApp_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText;

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD && pVal.BeforeAction)
                {
                    createFormItems(oForm, out errorText);
                    Program.FORM_LOAD_FOR_VISIBLE = true;
                    Program.FORM_LOAD_FOR_ACTIVATE = true;
                }

                else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE && !pVal.BeforeAction)
                {
                    if (Program.FORM_LOAD_FOR_VISIBLE)
                        Program.FORM_LOAD_FOR_VISIBLE = false;
                }

                else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_RESIZE && !pVal.BeforeAction)
                {
                    resizeForm(oForm);
                }

                else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE && !pVal.BeforeAction)
                {
                    if (Program.FORM_LOAD_FOR_ACTIVATE)
                    {
                        setVisibleFormItems(oForm);
                        Program.FORM_LOAD_FOR_ACTIVATE = false;
                    }
                    if (Program.openPaymentMeans)
                    {
                        Program.openPaymentMeans = false;
                        setVisibleFormItems(oForm);
                    }
                }

                else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                {
                    if (pVal.ItemUID == "creditActE" || pVal.ItemUID == "outDocE" || pVal.ItemUID == "95" || pVal.ItemUID == "234000005" || pVal.ItemUID == "5")
                    {
                        SAPbouiCOM.IChooseFromListEvent oCFLEvento = (SAPbouiCOM.IChooseFromListEvent)pVal;
                        chooseFromList(oForm, pVal, oCFLEvento, ref BubbleEvent);
                    }
                }
                else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_COMBO_SELECT)
                {
                    comboSelect(oForm, pVal);
                    if (!pVal.BeforeAction)
                    {
                        if (pVal.ItemUID == "opTypeCB" || pVal.ItemUID == "18" || pVal.ItemUID == "107")
                        {
                            if (pVal.ItemUID == "107" && oForm.DataSources.DBDataSources.Item("ORCT").GetValue("IsPaytoBnk", 0).Trim() != "Y")
                                return;
                            setVisibleFormItems(oForm);
                        }
                    }
                }

                else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
                {
                    if (pVal.BeforeAction)
                    {
                        if (pVal.ItemUID == "1")
                            CommonFunctions.fillDocRate(oForm, "ORCT", true);
                    }
                    else
                    {
                        if (pVal.ItemUID == "UsBlaAgRtS" && !pVal.InnerEvent)
                        {
                            CommonFunctions.fillDocRate(oForm, "ORCT");
                        }
                        else if (pVal.ItemUID == "57" || pVal.ItemUID == "56" || pVal.ItemUID == "58")
                        {
                            setVisibleFormItems(oForm);

                            oForm.Items.Item("95").Enabled = true;
                            oForm.Items.Item("234000005").Enabled = true;

                            SetUsBlaAgRtSAvailability(oForm);
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
                                var prjCode = oForm.DataSources.DBDataSources.Item("ORCT").GetValue("PrjCode", 0);
                                var agrNo = oForm.DataSources.DBDataSources.Item("ORCT").GetValue("AgrNo", 0);
                                oForm.Items.Item("95").Enabled = string.IsNullOrEmpty(prjCode);
                                oForm.Items.Item("95").Specific.Value = "";
                                
                                var item = oForm.Items.Item("234000005");
                                if (item.Visible)
                                {
                                    item.Enabled = string.IsNullOrEmpty(agrNo);
                                    item.Specific.Value = "";
                                }

                                SetUsBlaAgRtSAvailability(oForm);
                            }
                        }
                    }
                }
                if (pVal.ItemChanged)
                {
                    if (pVal.ItemUID == "10") //Item - Posting Date
                    {
                        string docEntry = oForm.DataSources.DBDataSources.Item("ORCT").GetValue("DocEntry", 0);
                        if (!string.IsNullOrEmpty(docEntry))
                        {

                        }
                        else
                        {
                            oForm.Items.Item("95").Enabled = true;
                            oForm.Items.Item("95").Specific.Value = "";

                            var item = oForm.Items.Item("234000005");
                            if (item.Visible)
                            {
                                item.Enabled = true;
                                item.Specific.Value = "";
                            }                            
                            SetUsBlaAgRtSAvailability(oForm);
                        }
                    }
                }
            }
        }

       public static void SetUsBlaAgRtSAvailability(SAPbouiCOM.Form oForm, bool enabled = false)
        {
            try
            {
                oForm.Freeze(true);

                var item = oForm.Items.Item("UsBlaAgRtS");
                SAPbouiCOM.CheckBox oCheckBox = (SAPbouiCOM.CheckBox)item.Specific;
                if (item.Enabled && oCheckBox.Checked)
                {
                    oCheckBox.Checked = false;
                    oForm.Items.Item("26").Click(); //Remark
                }
                item.Enabled = enabled;
            }
            catch
            {
                throw;
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        public static void attachOutgoingPayments(string paymentID, string documentNumber, string ePaymentID, string outDoc, string opType)
        {
            string incDoc;
            string errorText;
            try
            {
                BDOSInternetBanking.getPairPaymentsDocument("ORCT", paymentID, documentNumber, ePaymentID, opType, out incDoc);
                if (!string.IsNullOrEmpty(incDoc))
                {
                    SAPbobsCOM.Payments oIncomingPayment = null;
                    oIncomingPayment = (SAPbobsCOM.Payments)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments);
                    if (oIncomingPayment.GetByKey(Convert.ToInt32(incDoc)))
                    {
                        oIncomingPayment.UserFields.Fields.Item("U_outDoc").Value = outDoc;
                        int returnCode = oIncomingPayment.Update();
                        if (returnCode == 0)
                        {
                            bool bubbleEvent;
                            createJrnEntry(null, incDoc.ToString(), null, out bubbleEvent, out errorText);
                        }
                    }
                    Marshal.FinalReleaseComObject(oIncomingPayment);
                    oIncomingPayment = null;
                }
            }
            catch { }
        }

        public static string createDocumentTransferToOwnAccountType(SAPbouiCOM.DataTable oDataTable, int i, out int docEntry, out int docNum, out string errorText)
        {
            errorText = null;
            docEntry = 0;
            docNum = 0;
            SAPbobsCOM.Payments oPayments = null;
            oPayments = (SAPbobsCOM.Payments)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments);
            SAPbobsCOM.Payments oPaymentsNew = null;
            oPaymentsNew = (SAPbobsCOM.Payments)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments);
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
                if (!CommonFunctions.isAccountInHouseBankAccount(partnerAccountNumber + partnerCurrency))
                    errorText = errorText + BDOSResources.getTranslate("CouldNotFindHouseBankAccount") + " \"" + partnerAccountNumber + partnerCurrency + "\"! ";

                if (!string.IsNullOrEmpty(errorText))
                {
                    errorText = BDOSResources.getTranslate("OperationCompletedUnSuccessfully") + "! " + errorText + BDOSResources.getTranslate("TableRow") + " : " + (i + 1);
                    return null;
                }

                oPayments.ProjectCode = projectCod;
                oPayments.DocObjectCode = SAPbobsCOM.BoPaymentsObjectType.bopot_IncomingPayments;
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
                oPayments.AccountPayments.ProjectCode = projectCod;
                oPayments.AccountPayments.AccountCode = GLAccountCode;
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

                //outgoing - ის დოკუმენტის მოძებნა და მიბმა --->
                //string outDoc;
                //BDOSInternetBanking.getPairPaymentsDocument( "OVPM", oDataTable.GetValue("PaymentID", i), "transferToOwnAccount", out outDoc);
                //if (!string.IsNullOrEmpty(outDoc))
                //{
                //    oPayments.UserFields.Fields.Item("U_outDoc").Value = outDoc;
                //}
                //outgoing - ის დოკუმენტის მოძებნა და მიბმა <---

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

        public static string createDocumentTransferFromBPType(SAPbouiCOM.DataTable oDataTable, DataRow oDataTableDetail, bool onAccount, SAPbouiCOM.Form oForm, int i, out int docEntry, out int docNum, out string errorText, string transactionType = null)
        {
            errorText = null;
            docEntry = 0;
            docNum = 0;
            SAPbobsCOM.Payments oPayments = null;
            oPayments = (SAPbobsCOM.Payments)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments);
            SAPbobsCOM.Payments oPaymentsNew = null;
            oPaymentsNew = (SAPbobsCOM.Payments)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments);
            SAPbobsCOM.SBObob oSBOBob = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);

            bool automaticPaymentInternetBanking = (oForm.DataSources.UserDataSources.Item("autoPay").ValueEx == "Y");

            string dpTxt = "";

            try
            {
                string localCurrency = CommonFunctions.getLocalCurrency();

                DateTime docDate = oDataTable.GetValue("DocumentDate", i);
                DateTime valueDate = oDataTable.GetValue("ValueDate", i);
                string GLAccountCode = oDataTable.GetValue("GLAccountCode", i);
                string projectCod = oDataTableDetail == null ? oDataTable.GetValue("Project", i) : oDataTableDetail["Project"].ToString();

                if (string.IsNullOrEmpty(GLAccountCode))
                    errorText = BDOSResources.getTranslate("TheFollowingFieldIsMandatory") + " : \"" + BDOSResources.getTranslate("GLAccountCode") + "\"! ";
                string cashFlowLineItemName = oDataTable.GetValue("CashFlowLineItemName", i);
                string accountNumber = oDataTable.GetValue("AccountNumber", i);
                string currency = oDataTable.GetValue("Currency", i);
                string InvCurrency = oDataTableDetail == null ? oDataTable.GetValue("Currency", i) : oDataTableDetail["Currency"].ToString();

                string currencySapCode = CommonFunctions.getCurrencySapCode(currency);
                if (string.IsNullOrEmpty(currencySapCode))
                    errorText = errorText + BDOSResources.getTranslate("CouldNotFindCurrency") + " \"" + currency + "\"! ";
                string partnerAccountNumber = oDataTable.GetValue("PartnerAccountNumber", i);
                string partnerCurrency = oDataTable.GetValue("PartnerCurrency", i);
                string partnerCurrencySapCode = CommonFunctions.getCurrencySapCode(partnerCurrency);

                if (oDataTableDetail != null)
                {
                    partnerCurrencySapCode = InvCurrency;
                }

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

                string blnkAgr = oDataTableDetail == null ? oDataTable.GetValue("BlnkAgr", i) : oDataTableDetail["BlnkAgr"].ToString();

                string useBlaAgRt = oDataTableDetail == null ? oDataTable.GetValue("UseBlaAgRt", i) : oDataTableDetail["UseBlaAgRt"].ToString();

                SAPbobsCOM.Recordset oRecordSet = null;
                if (transactionType == OperationTypeFromIntBank.ReturnFromSupplier.ToString())
                {
                    oRecordSet = CommonFunctions.getBPBankInfo(partnerAccountNumber + partnerCurrency, partnerTaxCode, "S");
                }
                else
                {
                    oRecordSet = CommonFunctions.getBPBankInfo(partnerAccountNumber + partnerCurrency, partnerTaxCode, "C");
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

                string cardCode = oRecordSet.Fields.Item("CardCode").Value;
                oPayments.CardCode = cardCode;
                oPayments.CardName = oRecordSet.Fields.Item("CardName").Value;
                string BPCurrency = oRecordSet.Fields.Item("Currency").Value;
                //oPayments.PayToBankCountry = oRecordSet.Fields.Item("Country").Value;
                //oPayments.PayToBankCode = oRecordSet.Fields.Item("BankCode").Value;
                //oPayments.PayToBankAccountNo = oRecordSet.Fields.Item("Account").Value;
                //oPayments.IsPayToBank = SAPbobsCOM.BoYesNoEnum.tYES;
                oPayments.ProjectCode = projectCod;
                oPayments.ControlAccount = GLAccountCode;

                oPayments.ProjectCode = projectCod;
                oPayments.DocObjectCode = SAPbobsCOM.BoPaymentsObjectType.bopot_IncomingPayments;

                if (transactionType == OperationTypeFromIntBank.ReturnFromSupplier.ToString())
                {
                    oPayments.DocTypte = SAPbobsCOM.BoRcptTypes.rSupplier;
                }
                else
                {
                    oPayments.DocTypte = SAPbobsCOM.BoRcptTypes.rCustomer;
                }

                oPayments.DocDate = docDate;
                oPayments.TaxDate = docDate;

                decimal docRate = Convert.ToDecimal(oDataTable.GetValue("DocRateIN", i), NumberFormatInfo.InvariantInfo);

                double docRateByBlnktAgr = 0;
                double docRateINByBlnktAgr = 0;
                if (!string.IsNullOrEmpty(blnkAgr))
                {
                    oPayments.BlanketAgreement = Convert.ToInt32(blnkAgr);
                    oPayments.UserFields.Fields.Item("U_UseBlaAgRt").Value = useBlaAgRt;
                    string docCur;
                    if (useBlaAgRt == "Y")
                    {
                        docRateByBlnktAgr = Convert.ToDouble(BlanketAgreement.GetBlAgremeentCurrencyRate(Convert.ToInt32(blnkAgr), out docCur, docDate), NumberFormatInfo.InvariantInfo);
                        if (docRate != 0)
                            docRateINByBlnktAgr = Convert.ToDouble(BlanketAgreement.GetBlAgremeentCurrencyRate(Convert.ToInt32(blnkAgr), out docCur, null, docRate), NumberFormatInfo.InvariantInfo);
                    }
                }

                //docRate = useBlaAgRt == "Y" ? docRateByBlnktAgr : oSBOBob.GetCurrencyRate(BPCurrency, docDate).Fields.Item("CurrencyRate").Value;

                decimal transferSumLC = 0;
                decimal transferSumFC = 0;

                decimal amount = oDataTableDetail == null ? Convert.ToDecimal(oDataTable.GetValue("Amount", i), NumberFormatInfo.InvariantInfo) : Convert.ToDecimal(oDataTableDetail["Amount"], NumberFormatInfo.InvariantInfo);

                if (onAccount)
                {
                    amount = Convert.ToDecimal(oDataTable.GetValue("PaymentOnAccount", i), NumberFormatInfo.InvariantInfo);
                }


                decimal addDPAmt = 0;
                decimal addDPamtLocal = 0;
                if (automaticPaymentInternetBanking && oDataTableDetail == null)
                {
                    addDPAmt = Convert.ToDecimal(oDataTable.GetValue("AddDownPaymentAmount", i), NumberFormatInfo.InvariantInfo);

                }
                if (amount == 0)
                {
                    amount = addDPAmt;
                }

                if (currencySapCode == partnerCurrencySapCode || oDataTableDetail != null)
                {
                    if (BPCurrency != "##" && BPCurrency != localCurrency)
                    {
                        if (partnerCurrencySapCode == localCurrency)
                        {
                            oPayments.DocCurrency = BPCurrency;
                            oPayments.LocalCurrency = SAPbobsCOM.BoYesNoEnum.tYES;
                            if (docRate != 0)
                                oPayments.DocRate = useBlaAgRt == "Y" ? docRateINByBlnktAgr : Convert.ToDouble(docRate, NumberFormatInfo.InvariantInfo);
                            else
                                oPayments.DocRate = useBlaAgRt == "Y" ? docRateByBlnktAgr : oSBOBob.GetCurrencyRate(BPCurrency, docDate).Fields.Item("CurrencyRate").Value;
                            docRate = Convert.ToDecimal(oPayments.DocRate);
                            transferSumLC = amount;
                            transferSumFC = amount / docRate;
                        }
                        else if (partnerCurrencySapCode == BPCurrency)
                        {
                            oPayments.DocCurrency = BPCurrency;
                            oPayments.LocalCurrency = SAPbobsCOM.BoYesNoEnum.tNO;
                            if (docRate != 0)
                                oPayments.DocRate = useBlaAgRt == "Y" ? docRateINByBlnktAgr : Convert.ToDouble(docRate, NumberFormatInfo.InvariantInfo);
                            else
                                oPayments.DocRate = useBlaAgRt == "Y" ? docRateByBlnktAgr : oSBOBob.GetCurrencyRate(BPCurrency, docDate).Fields.Item("CurrencyRate").Value;
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
                            if (docRate != 0)
                                oPayments.DocRate = useBlaAgRt == "Y" ? docRateINByBlnktAgr : Convert.ToDouble(docRate, NumberFormatInfo.InvariantInfo);
                            else
                                oPayments.DocRate = useBlaAgRt == "Y" ? docRateByBlnktAgr : oSBOBob.GetCurrencyRate(partnerCurrencySapCode, docDate).Fields.Item("CurrencyRate").Value;
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

                transferSumLC = CommonFunctions.roundAmountByGeneralSettings(transferSumLC, "Sum");
                transferSumFC = CommonFunctions.roundAmountByGeneralSettings(transferSumFC, "Sum");

                //როდესაც სავალუტო თანხა 0-ია(სიმცირის გამო) დოკუმენტი არ შექმნას
                if (oPayments.DocRate > 0 && transferSumFC == 0)
                {
                    errorText = null;
                    return "";
                }

                amount = CommonFunctions.roundAmountByGeneralSettings(amount, "Sum");

                oPayments.TransferAccount = transferAccount;
                oPayments.TransferDate = docDate;
                oPayments.TransferSum = Convert.ToDouble(amount, NumberFormatInfo.InvariantInfo);
                if (onAccount == false)
                {
                    string expression = "LineNumExportMTR = '" + oDataTable.GetValue("LineNum", i) + "'";

                    if (oDataTableDetail != null)
                    {

                        expression = expression + " and Project = '" + oDataTableDetail["Project"] + "'"
                        + "and BlnkAgr = '" + oDataTableDetail["BlnkAgr"] + "'"
                        + "and useBlaAgRt = '" + oDataTableDetail["useBlaAgRt"] + "'"
                        + "and Currency = '" + oDataTableDetail["Currency"] + "'";

                    }

                    DataRow[] foundRows = BDOSInternetBanking.TableExportMTRForDetail.Select(expression);
                    string docType;
                    if (foundRows.Count() > 0)
                    {
                        for (int j = 0; j < foundRows.Count(); j++)
                        {
                            docType = Convert.ToString(foundRows[j]["DocType"]);
                            oPayments.Invoices.DocEntry = Convert.ToInt32(foundRows[j]["DocEntry"]);
                            oPayments.Invoices.InstallmentId = Convert.ToInt32(foundRows[j]["InstallmentID"]);

                            if (docType == "13")
                                oPayments.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_Invoice;
                            else if (docType == "14")
                                oPayments.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_CredItnote;
                            else if (docType == "203")
                                oPayments.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_DownPayment;
                            else if (docType == "163")
                                oPayments.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_APCorrectionInvoice;
                            else if (docType == "30")
                            {
                                oPayments.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_JournalEntry;
                                oPayments.Invoices.DocLine = Convert.ToInt32(foundRows[j]["LineID"]);
                            }

                            oPayments.Invoices.SumApplied = Convert.ToDouble(foundRows[j]["TotalPaymentLocal"], NumberFormatInfo.InvariantInfo);
                            if (foundRows[j]["Currency"].ToString() != localCurrency)
                                oPayments.Invoices.AppliedFC = Convert.ToDouble(foundRows[j]["TotalPayment"], NumberFormatInfo.InvariantInfo);
                            oPayments.Invoices.Add();
                        }
                    }


                    if (automaticPaymentInternetBanking)
                    {
                        //Tu Tanxa darcha iqmneba avansi da emateba cxrilshi invoisebtan ertad
                        int dpdocEntry;
                        int dpdocNum;

                        if (addDPAmt > 0)
                        {
                            oDataTable.SetValue("AddDownPaymentAmount", i, Convert.ToDouble(addDPAmt));
                            dpTxt = ARDownPaymentRequest.createDocumentTransferFromBPType(oDataTable, oForm, i, "", "", out dpdocEntry, out dpdocNum, out errorText);

                            oPayments.Invoices.DocEntry = dpdocEntry;
                            oPayments.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_DownPayment;

                            oPayments.Invoices.SumApplied = Convert.ToDouble(addDPAmt);
                        }

                        if (!string.IsNullOrEmpty(errorText))
                        {
                            return null;
                        }
                    }
                }
                else
                {
                    if (addDPAmt > 0)
                    {
                        int dpdocEntry;
                        int dpdocNum;

                        oDataTable.SetValue("AddDownPaymentAmount", i, Convert.ToDouble(addDPAmt));
                        dpTxt = ARDownPaymentRequest.createDocumentTransferFromBPType(oDataTable, oForm, i, "", "", out dpdocEntry, out dpdocNum, out errorText);

                        oPayments.Invoices.DocEntry = dpdocEntry;
                        oPayments.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_DownPayment;

                        oPayments.Invoices.SumApplied = Convert.ToDouble(addDPAmt);
                    }

                    if (!string.IsNullOrEmpty(errorText))
                    {
                        return null;
                    }
                }


                Marshal.FinalReleaseComObject(oRecordSet);
                oRecordSet = null;

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
                        oDataTable.SetValue("InDetail", i, "");
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

        public static string createDocumentOtherIncomesType(SAPbouiCOM.DataTable oDataTable, int i, out int docEntry, out int docNum, out string errorText)
        {
            errorText = null;
            docEntry = 0;
            docNum = 0;
            SAPbobsCOM.Payments oPayments = null;
            oPayments = (SAPbobsCOM.Payments)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments);
            SAPbobsCOM.Payments oPaymentsNew = null;
            oPaymentsNew = (SAPbobsCOM.Payments)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments);
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

                if (!string.IsNullOrEmpty(errorText))
                {
                    errorText = BDOSResources.getTranslate("OperationCompletedUnSuccessfully") + "! " + errorText + BDOSResources.getTranslate("TableRow") + " : " + (i + 1);
                    return null;
                }

                oPayments.ProjectCode = projectCod;
                oPayments.DocObjectCode = SAPbobsCOM.BoPaymentsObjectType.bopot_IncomingPayments;
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
                oPayments.AccountPayments.ProjectCode = projectCod;
                oPayments.AccountPayments.AccountCode = GLAccountCode;
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
            oPayments = (SAPbobsCOM.Payments)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments);
            SAPbobsCOM.Payments oPaymentsNew = null;
            oPaymentsNew = (SAPbobsCOM.Payments)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments);
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
                if (string.IsNullOrEmpty(partnerCurrencySapCode))
                    errorText = errorText + BDOSResources.getTranslate("CouldNotFindCurrency") + " \"" + partnerCurrencySapCode + "\"! ";
                if (string.IsNullOrEmpty(currencyExchangeSapCode))
                    errorText = errorText + BDOSResources.getTranslate("CouldNotFindCurrency") + " \"" + currencyExchangeSapCode + "\"! ";
                if (!CommonFunctions.isAccountInHouseBankAccount(partnerAccountNumber + partnerCurrency))
                    errorText = errorText + BDOSResources.getTranslate("CouldNotFindHouseBankAccount") + " \"" + partnerAccountNumber + partnerCurrency + "\"! ";
                string transferAccount = CommonFunctions.getTransferAccount(accountNumber + currency);
                if (string.IsNullOrEmpty(transferAccount))
                    errorText = errorText + BDOSResources.getTranslate("CheckGLAccountForHouseBankAccount") + " \"" + accountNumber + currency + "\"! ";
                bool cashFlowRelevant = CommonFunctions.isAccountCashFlowRelevant(transferAccount);
                string cashFlowLineItemID = oDataTable.GetValue("CashFlowLineItemID", i);
                if (cashFlowRelevant && string.IsNullOrEmpty(cashFlowLineItemID))
                    errorText = errorText + BDOSResources.getTranslate("TheFollowingFieldIsMandatory") + " : \"" + BDOSResources.getTranslate("CashFlowLineItemID") + "\"! ";

                if (!string.IsNullOrEmpty(errorText))
                {
                    errorText = BDOSResources.getTranslate("OperationCompletedUnSuccessfully") + "! " + errorText + BDOSResources.getTranslate("TableRow") + " : " + (i + 1);
                    return null;
                }

                oPayments.ProjectCode = projectCod;
                oPayments.DocObjectCode = SAPbobsCOM.BoPaymentsObjectType.bopot_IncomingPayments;
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
                oPayments.AccountPayments.ProjectCode = projectCod;
                oPayments.AccountPayments.AccountCode = GLAccountCode;
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

                //outgoing - ის დოკუმენტის მოძებნა და მიბმა --->
                string outDoc;
                BDOSInternetBanking.getPairPaymentsDocument("OVPM", oDataTable.GetValue("PaymentID", i), oDataTable.GetValue("DocumentNumber", i), oDataTable.GetValue("ExternalPaymentID", i), "currencyExchange", out outDoc);
                if (!string.IsNullOrEmpty(outDoc))
                {
                    oPayments.UserFields.Fields.Item("U_outDoc").Value = outDoc;
                }
                //outgoing - ის დოკუმენტის მოძებნა და მიბმა <---

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

                        decimal DocRate = Convert.ToDecimal(oPaymentsNew.DocRate);
                        string DocCurrency = oPaymentsNew.DocCurrency;
                        DateTime DocDate = oPaymentsNew.DocDate;

                        oDataTable.SetValue("DocEntry", i, docEntry.ToString());
                        oDataTable.SetValue("DocNum", i, docNum.ToString());

                        bool bubbleEvent;
                        createJrnEntry(null, docEntry.ToString(), null, out bubbleEvent, out errorText);

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

        public static string createDocumentTransferFromBPCashType(SAPbouiCOM.DataTable oDataTable, SAPbouiCOM.Form oForm, int i, out int docEntry, out int docNum, out string errorText)
        {
            errorText = null;
            docEntry = 0;
            docNum = 0;
            SAPbobsCOM.Payments oPayments = null;
            oPayments = (SAPbobsCOM.Payments)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments);
            SAPbobsCOM.Payments oPaymentsNew = null;
            oPaymentsNew = (SAPbobsCOM.Payments)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments);
            SAPbobsCOM.SBObob oSBOBob = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);

            bool automaticPaymentInternetBanking = true;

            string dpTxt = "";

            try
            {
                SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                string localCurrency = CommonFunctions.getLocalCurrency();

                DateTime docDate = oDataTable.GetValue("DocumentDate", i);

                string projectCod = oDataTable.GetValue("Project", i);

                string cashAccount = oDataTable.GetValue("CashAccount", i);
                if (string.IsNullOrEmpty(cashAccount))
                    errorText = BDOSResources.getTranslate("TheFollowingFieldIsMandatory") + " : \"" + BDOSResources.getTranslate("CashAccount") + "\"! ";

                string currency = oDataTable.GetValue("Currency", i);
                string currencySapCode = CommonFunctions.getCurrencySapCode(currency);
                if (string.IsNullOrEmpty(currencySapCode))
                    errorText = errorText + BDOSResources.getTranslate("CouldNotFindCurrency") + " \"" + currency + "\"! ";

                string partnerCurrency = oDataTable.GetValue("PartnerCurrency", i);
                string partnerCurrencySapCode = CommonFunctions.getCurrencySapCode(partnerCurrency);
                if (string.IsNullOrEmpty(partnerCurrencySapCode))
                    errorText = errorText + BDOSResources.getTranslate("CouldNotFindCurrency") + " \"" + partnerCurrencySapCode + "\"! ";

                bool cashFlowRelevant = CommonFunctions.isAccountCashFlowRelevant(cashAccount);
                string cashFlowLineItemID = oDataTable.GetValue("CashFlowLineItemID", i);
                if (cashFlowRelevant && string.IsNullOrEmpty(cashFlowLineItemID))
                    errorText = errorText + BDOSResources.getTranslate("TheFollowingFieldIsMandatory") + " : \"" + BDOSResources.getTranslate("CashFlowLineItemID") + "\"! ";

                string blnkAgr = oDataTable.GetValue("BlnkAgr", i);

                string GLAccountCode = oDataTable.GetValue("GLAccountCode", i);
                decimal paymentOnAccount = Convert.ToDecimal(oDataTable.GetValue("PaymentOnAccount", i), NumberFormatInfo.InvariantInfo);
                if (string.IsNullOrEmpty(GLAccountCode) && paymentOnAccount > 0)
                    errorText = BDOSResources.getTranslate("TheFollowingFieldIsMandatory") + " : \"" + BDOSResources.getTranslate("GLAccountCode") + "\"! ";


                if (!string.IsNullOrEmpty(errorText))
                {
                    errorText = BDOSResources.getTranslate("OperationCompletedUnSuccessfully") + "! " + errorText;
                    return null;
                }

                string cardCode = oDataTable.GetValue("CardCode", i);
                oPayments.CardCode = cardCode;
                oPayments.CardName = oDataTable.GetValue("CardName", i);
                string BPCurrency = oDataTable.GetValue("BPCurrency", i);

                oPayments.ControlAccount = GLAccountCode;
                oPayments.ProjectCode = projectCod;
                oPayments.DocObjectCode = SAPbobsCOM.BoPaymentsObjectType.bopot_IncomingPayments;
                oPayments.DocTypte = SAPbobsCOM.BoRcptTypes.rCustomer;
                oPayments.DocDate = docDate;
                oPayments.TaxDate = docDate;

                if (!string.IsNullOrEmpty(blnkAgr))
                {
                    oPayments.BlanketAgreement = Convert.ToInt32(blnkAgr);
                }

                decimal docRate = Convert.ToDecimal(oDataTable.GetValue("DocRateIN", i), NumberFormatInfo.InvariantInfo);
                decimal cashSumLC = 0;
                decimal cashSumFC = 0;
                decimal amount = Convert.ToDecimal(oDataTable.GetValue("Amount", i), NumberFormatInfo.InvariantInfo);

                decimal invoicesamount = Convert.ToDecimal(oDataTable.GetValue("InvoicesAmount", i), NumberFormatInfo.InvariantInfo);

                decimal addDPAmt = 0;
                if (automaticPaymentInternetBanking)
                {
                    addDPAmt = Convert.ToDecimal(oDataTable.GetValue("AddDownPaymentAmount", i), NumberFormatInfo.InvariantInfo);
                }

                if (currencySapCode == partnerCurrencySapCode)
                {
                    if (BPCurrency != "##" && BPCurrency != localCurrency)
                    {
                        if (partnerCurrencySapCode == localCurrency)
                        {
                            oPayments.DocCurrency = BPCurrency;
                            oPayments.LocalCurrency = SAPbobsCOM.BoYesNoEnum.tYES;
                            if (docRate != 0)
                                oPayments.DocRate = Convert.ToDouble(docRate, NumberFormatInfo.InvariantInfo);
                            else
                                oPayments.DocRate = oSBOBob.GetCurrencyRate(BPCurrency, docDate).Fields.Item("CurrencyRate").Value;
                            docRate = Convert.ToDecimal(oPayments.DocRate);
                            cashSumLC = amount;
                            cashSumFC = amount / docRate;
                        }
                        else if (partnerCurrencySapCode == BPCurrency)
                        {
                            oPayments.DocCurrency = BPCurrency;
                            oPayments.LocalCurrency = SAPbobsCOM.BoYesNoEnum.tNO;
                            if (docRate != 0)
                                oPayments.DocRate = Convert.ToDouble(docRate, NumberFormatInfo.InvariantInfo);
                            else
                                oPayments.DocRate = oSBOBob.GetCurrencyRate(BPCurrency, docDate).Fields.Item("CurrencyRate").Value;
                            docRate = Convert.ToDecimal(oPayments.DocRate);
                            cashSumLC = amount * docRate;
                            cashSumFC = amount;
                        }
                        else
                        {
                            errorText = BDOSResources.getTranslate("OperationCompletedUnSuccessfully") + "! " + BDOSResources.getTranslate("CheckCurrencies") + "! "; // + errorText + BDOSResources.getTranslate("TableRow") + " : " + (i + 1)
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
                            cashSumLC = amount;
                            cashSumFC = 0;
                        }
                        else
                        {
                            errorText = BDOSResources.getTranslate("OperationCompletedUnSuccessfully") + "! " + BDOSResources.getTranslate("CheckCurrencies") + "! "; // + errorText + BDOSResources.getTranslate("TableRow") + " : " + (i + 1)
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
                            cashSumLC = amount;
                            cashSumFC = 0;
                        }
                        else if (partnerCurrencySapCode != localCurrency)
                        {
                            oPayments.DocCurrency = partnerCurrencySapCode;
                            oPayments.LocalCurrency = SAPbobsCOM.BoYesNoEnum.tNO;
                            if (docRate != 0)
                                oPayments.DocRate = Convert.ToDouble(docRate, NumberFormatInfo.InvariantInfo);
                            else
                                oPayments.DocRate = oSBOBob.GetCurrencyRate(partnerCurrencySapCode, docDate).Fields.Item("CurrencyRate").Value;
                            docRate = Convert.ToDecimal(oPayments.DocRate);
                            cashSumLC = amount * docRate;
                            cashSumFC = amount;
                        }
                        else
                        {
                            errorText = BDOSResources.getTranslate("OperationCompletedUnSuccessfully") + "! " + BDOSResources.getTranslate("CheckCurrencies") + "! "; // + errorText + BDOSResources.getTranslate("TableRow") + " : " + (i + 1)
                            return null;
                        }
                    }
                }
                else
                {
                    errorText = BDOSResources.getTranslate("OperationCompletedUnSuccessfully") + "! " + BDOSResources.getTranslate("CheckCurrencies") + "! "; // + errorText + BDOSResources.getTranslate("TableRow") + " : " + (i + 1)
                    return null;
                }

                cashSumLC = CommonFunctions.roundAmountByGeneralSettings(cashSumLC, "Sum");
                cashSumFC = CommonFunctions.roundAmountByGeneralSettings(cashSumFC, "Sum");
                amount = CommonFunctions.roundAmountByGeneralSettings(amount, "Sum");

                oPayments.CashAccount = cashAccount;
                oPayments.CashSum = Convert.ToDouble(amount, NumberFormatInfo.InvariantInfo);

                string expression = "LineNumExportMTR = '" + oDataTable.GetValue("LineNum", i) + "'";
                DataRow[] foundRows = BDOSInternetBanking.TableExportMTRForDetail.Select(expression);
                string docType;
                if (foundRows.Count() > 0)
                {
                    for (int j = 0; j < foundRows.Count(); j++)
                    {
                        docType = Convert.ToString(foundRows[j]["DocType"]);
                        oPayments.Invoices.DocEntry = Convert.ToInt32(foundRows[j]["DocEntry"]);
                        oPayments.Invoices.InstallmentId = Convert.ToInt32(foundRows[j]["InstallmentID"]);

                        if (docType == "13")
                            oPayments.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_Invoice;
                        else if (docType == "14")
                            oPayments.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_CredItnote;
                        else if (docType == "203")
                            oPayments.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_DownPayment;
                        else if (docType == "30")
                        {
                            oPayments.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_JournalEntry;
                            oPayments.Invoices.DocLine = Convert.ToInt32(foundRows[j]["LineID"]);
                        }

                        oPayments.Invoices.SumApplied = Convert.ToDouble(foundRows[j]["TotalPaymentLocal"], NumberFormatInfo.InvariantInfo);
                        if (foundRows[j]["Currency"].ToString() != localCurrency)
                            oPayments.Invoices.AppliedFC = Convert.ToDouble(foundRows[j]["TotalPayment"], NumberFormatInfo.InvariantInfo);
                        oPayments.Invoices.Add();
                    }
                }

                if (automaticPaymentInternetBanking)
                {
                    //Tu Tanxa darcha iqmneba avansi da emateba cxrilshi invoisebtan ertad
                    int dpdocEntry;
                    int dpdocNum;

                    if (addDPAmt > 0)
                    {
                        oDataTable.SetValue("AddDownPaymentAmount", i, Convert.ToDouble(addDPAmt));
                        dpTxt = ARDownPaymentRequest.createDocumentTransferFromBPType(oDataTable, oForm, i, cardCode, BPCurrency, out dpdocEntry, out dpdocNum, out errorText);

                        oPayments.Invoices.DocEntry = dpdocEntry;
                        oPayments.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_DownPayment;

                        oPayments.Invoices.SumApplied = Convert.ToDouble(addDPAmt);
                    }

                    if (!string.IsNullOrEmpty(errorText))
                    {
                        return null;
                    }
                }

                Marshal.FinalReleaseComObject(oRecordSet);
                oRecordSet = null;

                if (cashFlowRelevant)
                {
                    oPayments.PrimaryFormItems.CashFlowLineItemID = Convert.ToInt32(cashFlowLineItemID);
                    oPayments.PrimaryFormItems.AmountFC = Convert.ToDouble(cashSumFC, NumberFormatInfo.InvariantInfo);
                    if (oPayments.DocCurrency == localCurrency)
                        oPayments.PrimaryFormItems.AmountLC = Convert.ToDouble(cashSumLC, NumberFormatInfo.InvariantInfo);
                    oPayments.PrimaryFormItems.PaymentMeans = SAPbobsCOM.PaymentMeansTypeEnum.pmtCash;
                    oPayments.PrimaryFormItems.Add();
                }

                int returnCode = oPayments.Add();
                if (returnCode != 0)
                {
                    int errCode;
                    string errMsg;
                    Program.oCompany.GetLastError(out errCode, out errMsg);
                    errorText = BDOSResources.getTranslate("OperationCompletedUnSuccessfully") + "! " + errMsg + "! "; // + BDOSResources.getTranslate("TableRow") + " : " + (i + 1)
                    return null;
                }
                else
                {
                    bool newDoc = oPaymentsNew.GetByKey(Convert.ToInt32(Program.oCompany.GetNewObjectKey()));
                    if (newDoc)
                    {
                        docEntry = oPaymentsNew.DocEntry;
                        docNum = oPaymentsNew.DocNum;

                        return BDOSResources.getTranslate("OperationCompletedSuccessfully") + "! "; // + BDOSResources.getTranslate("TableRow") + " : " + (i + 1)
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

        public static bool isARDownPaymentVATAccrual(int docEntry)
        {
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                String query = "";
                query = query + "SELECT \"DocEntry\"  AS \"ARDownPaymentVATAccrual\", " + "\n";
                query = query + "       \"U_baseDoc\" AS \"ARDownPayment\" " + "\n";
                query = query + "FROM  (SELECT \"RCT2\".\"DocNum\"    AS \"IncomingPayment\", " + "\n";
                query = query + "              \"RCT2\".\"DocEntry\"  AS \"ARDownPayment\", " + "\n";
                query = query + "              \"ODPI\".\"DocStatus\" AS \"ARDownPaymentStatus\" " + "\n";
                query = query + "       FROM   \"RCT2\" " + "\n";
                query = query + "              INNER JOIN \"ODPI\" " + "\n";
                query = query + "                      ON \"RCT2\".\"DocEntry\" = \"ODPI\".\"DocEntry\" " + "\n";
                query = query + "       WHERE  \"RCT2\".\"DocNum\" = '" + docEntry + "' " + "\n";
                query = query + "              AND \"RCT2\".\"InvType\" = '203' " + "\n";
                query = query + "              AND \"ODPI\".\"DocStatus\" = 'C' " + "\n";
                query = query + "              AND \"ODPI\".\"CANCELED\" = 'N') AS \"T1\" " + "\n";
                query = query + "      INNER JOIN \"@BDOSARDV\" " + "\n";
                query = query + "              ON \"T1\".\"ARDownPayment\" = \"@BDOSARDV\".\"U_baseDoc\" " + "\n";
                query = query + "WHERE  \"Canceled\" = 'N' ";
                query = query + "       AND \"U_baseDocT\" = '203' " + "\n";

                oRecordSet.DoQuery(query.ToString());
                if (!oRecordSet.EoF)
                    return true;
                return false;
            }
            catch
            {
                throw;
            }
            finally
            {
                Marshal.ReleaseComObject(oRecordSet);
            }
        }

        public static void cancellation(SAPbouiCOM.Form oForm, int docEntry)
        {
            try
            {
                JournalEntry.cancellation(oForm, docEntry, "24", out var errorText);
                if (!string.IsNullOrEmpty(errorText))
                {
                    throw new Exception(errorText);
                }
            }
            catch
            {
                throw;
            }
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
            catch
            {
                throw;
            }
            finally
            {
                oForm.Freeze(false);
            }
        }
    }
}
