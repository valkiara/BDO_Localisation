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
    static partial class BDOSInternetBankingDocuments
    {
        public static decimal addDownPaymentAmount;
        public static decimal downPaymentAmount;
        public static decimal invoicesAmount;
        public static decimal paymentOnAccount;
        public static SAPbouiCOM.Form oFormBDOSInternetBanking;
        public static decimal docRateIN;
        public static bool automaticPaymentInternetBanking;
        public static bool openFromBlnkAgr = false;

        public static void createForm(  SAPbouiCOM.Form FormBDOSInternetBanking, DateTime docDate, string cardCode, string cardName, string BPCurrency, decimal amount, string currency, decimal downPaymentAmount, decimal invoicesAmount, decimal paymentOnAccount, decimal addDPAmount, decimal docRateIN, string transactionType,out SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            oForm = null;

            oFormBDOSInternetBanking = FormBDOSInternetBanking;

            int formHeight = Program.uiApp.Desktop.Width;
            int formWidth = Program.uiApp.Desktop.Width;

            //ფორმის აუცილებელი თვისებები
            string formTitle = (openFromBlnkAgr ? BDOSResources.getTranslate("Payment") : BDOSResources.getTranslate("InternetBanking") + " (" + BDOSResources.getTranslate("InDetail") + ")");

            Dictionary<string, object> formProperties = new Dictionary<string, object>();
            formProperties.Add("UniqueID", "BDOSINBDOC");
            formProperties.Add("BorderStyle", SAPbouiCOM.BoFormBorderStyle.fbs_Sizable);
            formProperties.Add("Title", formTitle);
            formProperties.Add("ClientWidth", formWidth);
            formProperties.Add("ClientHeight", formHeight);

            bool newForm;
            bool formExist = FormsB1.createForm( formProperties, out oForm, out newForm, out errorText);

            if (errorText != null)
            {
                return;
            }

            if (formExist == true)
            {
                if (newForm == true)
                {
                    errorText = null;
                    Dictionary<string, object> formItems;
                    string itemName = "";
                    NumberFormatInfo Nfi = new NumberFormatInfo() { NumberDecimalSeparator = "." };

                    int left = 6;
                    int top = 6;
                    int height_e = 15;
                    int height = oForm.ClientHeight - top - 8 * height_e - 1 - 30;
                    int width = oForm.ClientWidth;

                    int left_s = 6;
                    int left_e = 90;
                    int width_s = 80;
                    int width_e = 148;

                    formItems = new Dictionary<string, object>();
                    itemName = "dateS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height_e);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("Date"));
                    formItems.Add("LinkTo", "dateE");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    string docDateStr = docDate.ToString("yyyyMMdd");
                    formItems = new Dictionary<string, object>();
                    itemName = "dateE";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_DATE);
                    formItems.Add("Length", 1);
                    formItems.Add("Size", 20);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("Left", left_e);
                    formItems.Add("Width", width_e / 2);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height_e);
                    formItems.Add("UID", itemName);
                    formItems.Add("Enabled", (openFromBlnkAgr ? true : false));
                    formItems.Add("ValueEx", docDateStr);
                    formItems.Add("AffectsFormMode", false);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left_s = left_e + width_e;

                    formItems = new Dictionary<string, object>();
                    itemName = "cardCodeS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height_e);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("CardCode"));
                    formItems.Add("LinkTo", "cardCodeE");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left_e = left_s + width_s + 25;

                    formItems = new Dictionary<string, object>();
                    itemName = "cardCodeE"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("Length", 15);
                    formItems.Add("Size", 15);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("Left", left_e);
                    formItems.Add("Width", width_e / 2);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height_e);
                    formItems.Add("UID", itemName);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("Enabled", false);
                    formItems.Add("ValueEx", cardCode);
                    formItems.Add("AffectsFormMode", false);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "cardCodeNE"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_LONG_TEXT);
                    formItems.Add("Length", 15);
                    formItems.Add("Size", 100);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("Left", left_e + width_e / 2 + 2);
                    formItems.Add("Width", width_e);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height_e);
                    formItems.Add("UID", itemName);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("Enabled", false);
                    formItems.Add("ValueEx", cardName);
                    formItems.Add("AffectsFormMode", false);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "cardCodeLB"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    formItems.Add("Left", left_e - 20);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height_e);
                    formItems.Add("UID", itemName);
                    formItems.Add("LinkTo", "cardCodeE");
                    formItems.Add("LinkedObjectType", "2"); //SAPbouiCOM.BoLinkedObject.lf_BusinessPartner, Business Partner object

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    top = top + 2 * height_e + 1;

                    formItems = new Dictionary<string, object>();
                    itemName = "docRateINS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_e);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height_e);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("BPCurrencyRate"));
                    formItems.Add("LinkTo", "docRateINE");
                    if (BPCurrency != "##" && BPCurrency != Program.MainCurrency)
                        formItems.Add("Visible", true);
                    else
                        formItems.Add("Visible", false);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "docRateINE"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_RATE);
                    formItems.Add("Length", 15);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("Left", left_e + width_e / 2 + 2);
                    formItems.Add("Width", width_e);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height_e);
                    formItems.Add("UID", itemName);
                    formItems.Add("DisplayDesc", true);
                    //formItems.Add("Enabled", true);
                    //formItems.Add("ValueEx", rate);
                    formItems.Add("AffectsFormMode", false);
                    if (BPCurrency != "##" && BPCurrency != Program.MainCurrency)
                    {
                        formItems.Add("Visible", true);
                        if (docRateIN == 0)
                        {
                            SAPbobsCOM.SBObob oSBOBob = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);
                            double rate = oSBOBob.GetCurrencyRate(BPCurrency, docDate).Fields.Item("CurrencyRate").Value;
                            formItems.Add("ValueEx", rate.ToString(Nfi));
                        }
                        else
                            formItems.Add("ValueEx", docRateIN.ToString(Nfi));
                    }
                    else
                        formItems.Add("Visible", false);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left_s = 6;
                    left_e = 127;
                    width_s = 121;
                    width_e = 148;

                    formItems = new Dictionary<string, object>();
                    itemName = "amountFBS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height_e);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", (openFromBlnkAgr ? BDOSResources.getTranslate("PaymentAmount") : BDOSResources.getTranslate("AmountFromBank")));
                    formItems.Add("LinkTo", "amountFBE");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "amountFBE"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SUM);
                    formItems.Add("Length", 15);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("Left", left_e);
                    formItems.Add("Width", width_e / 2);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height_e);
                    formItems.Add("UID", itemName);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("Enabled", (openFromBlnkAgr ? true : false));
                    formItems.Add("ValueEx", amount.ToString(Nfi));
                    formItems.Add("AffectsFormMode", false);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    Dictionary<string, string> listValidValuesDict = CommonFunctions.getCurrencyListForValidValues();

                    formItems = new Dictionary<string, object>();
                    itemName = "currencyCB"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("Length", 30);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                    formItems.Add("Left", left_e + width_e / 2 + 5);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height_e);
                    formItems.Add("UID", itemName);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("ValidValues", listValidValuesDict);
                    formItems.Add("Description", BDOSResources.getTranslate("Currency"));
                    formItems.Add("Enabled", false);
                    formItems.Add("ValueEx", currency);
                    formItems.Add("AffectsFormMode", false);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    top = top + height_e + 1;

                    formItems = new Dictionary<string, object>();
                    itemName = "totalAmtDS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height_e);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("TotalAmountDue"));
                    formItems.Add("LinkTo", "totalAmtDE");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "totalAmtDE"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SUM);
                    formItems.Add("Length", 15);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("Left", left_e);
                    formItems.Add("Width", width_e / 2);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height_e);
                    formItems.Add("UID", itemName);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("Enabled", false);
                    formItems.Add("ValueEx", (invoicesAmount + downPaymentAmount).ToString(Nfi));
                    formItems.Add("AffectsFormMode", false);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    top = top + height_e + 1;

                    formItems = new Dictionary<string, object>();
                    itemName = "dPaymentS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s + 10);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height_e);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("DownPaymentAmount")); //ავანსის თანხა
                    formItems.Add("LinkTo", "dPaymentE");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "dPaymentE"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SUM);
                    formItems.Add("Length", 15);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("Left", left_e + 10);
                    formItems.Add("Width", width_e / 2);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height_e);
                    formItems.Add("UID", itemName);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("Enabled", false);
                    formItems.Add("ValueEx", downPaymentAmount.ToString(Nfi));
                    formItems.Add("AffectsFormMode", false);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    top = top + height_e + 1;

                    formItems = new Dictionary<string, object>();
                    itemName = "invAmtS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s + 10);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height_e);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("InvoicesAmount")); //ინვოისის თანხა
                    formItems.Add("LinkTo", "invAmtE");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "invAmtE"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SUM);
                    formItems.Add("Length", 15);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("Left", left_e + 10);
                    formItems.Add("Width", width_e / 2);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height_e);
                    formItems.Add("UID", itemName);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("Enabled", false);
                    formItems.Add("ValueEx", invoicesAmount.ToString(Nfi));
                    formItems.Add("AffectsFormMode", false);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    top = top + height_e + 1;

                    formItems = new Dictionary<string, object>();
                    itemName = "onAccountS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height_e);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("PaymentOnAccount"));
                    formItems.Add("LinkTo", "onAccountE");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "onAccountE"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SUM);
                    formItems.Add("Length", 15);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("Left", left_e);
                    formItems.Add("Width", width_e / 2);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height_e);
                    formItems.Add("UID", itemName);
                    formItems.Add("DisplayDesc", true);

                    if (automaticPaymentInternetBanking)
                        formItems.Add("Enabled", true);
                    else
                        formItems.Add("Enabled", false);

                    formItems.Add("ValueEx", paymentOnAccount.ToString(Nfi));
                    formItems.Add("AffectsFormMode", false);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    //Automatic Payment
                    if (automaticPaymentInternetBanking)
                    {
                        top = top + height_e + 1;

                        formItems = new Dictionary<string, object>();
                        itemName = "AddDPAmtS"; //10 characters
                        formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                        formItems.Add("Left", left_s);
                        formItems.Add("Width", width_s);
                        formItems.Add("Top", top);
                        formItems.Add("Height", height_e);
                        formItems.Add("UID", itemName);
                        formItems.Add("Caption", BDOSResources.getTranslate("AddDPAmount"));
                        formItems.Add("LinkTo", "DpAmountE");

                        FormsB1.createFormItem(oForm, formItems, out errorText);
                        if (errorText != null)
                        {
                            return;
                        }

                        formItems = new Dictionary<string, object>();
                        itemName = "AddDPAmtE"; //10 characters
                        formItems.Add("isDataSource", true);
                        formItems.Add("DataSource", "UserDataSources");
                        formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SUM);
                        formItems.Add("Length", 15);
                        formItems.Add("TableName", "");
                        formItems.Add("Alias", itemName);
                        formItems.Add("Bound", true);
                        formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                        formItems.Add("Left", left_e);
                        formItems.Add("Width", width_e / 2);
                        formItems.Add("Top", top);
                        formItems.Add("Height", height_e);
                        formItems.Add("UID", itemName);
                        formItems.Add("DisplayDesc", true);
                        formItems.Add("Enabled", false);
                        formItems.Add("ValueEx", addDPAmount.ToString(Nfi));
                        formItems.Add("AffectsFormMode", false);

                        FormsB1.createFormItem(oForm, formItems, out errorText);
                        if (errorText != null)
                        {
                            return;
                        }
                    }
                    //Automatic Payment

                    if (openFromBlnkAgr)
                    {
                        top = top + height_e + 1;

                        string objectType = "1"; //Account
                        bool multiSelection = false;
                        string uniqueID_lf_CashAct = "CashAct_CFL";
                        FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_lf_CashAct);

                        left_e = 343;
                        formItems = new Dictionary<string, object>();
                        itemName = "CashActS"; //10 characters
                        formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                        formItems.Add("Left", left_e);
                        formItems.Add("Width", width_s);
                        formItems.Add("Top", top);
                        formItems.Add("Height", height_e);
                        formItems.Add("UID", itemName);
                        formItems.Add("Caption", BDOSResources.getTranslate("CashAccount"));
                        formItems.Add("LinkTo", "CashActE");

                        FormsB1.createFormItem(oForm, formItems, out errorText);
                        if (errorText != null)
                        {
                            return;
                        }

                        formItems = new Dictionary<string, object>();
                        itemName = "CashActE"; //10 characters
                        formItems.Add("isDataSource", true);
                        formItems.Add("DataSource", "UserDataSources");
                        formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                        formItems.Add("Length", 15);
                        formItems.Add("Size", 15);
                        formItems.Add("TableName", "");
                        formItems.Add("Alias", itemName);
                        formItems.Add("Bound", true);
                        formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                        formItems.Add("Left", left_e + width_e / 2 + 2);
                        formItems.Add("Width", width_e);
                        formItems.Add("Top", top);
                        formItems.Add("Height", height_e);
                        formItems.Add("UID", itemName);
                        formItems.Add("DisplayDesc", true);
                        formItems.Add("Enabled", true);
                        formItems.Add("AffectsFormMode", false);
                        formItems.Add("ChooseFromListUID", uniqueID_lf_CashAct);
                        formItems.Add("ChooseFromListAlias", "AcctCode");

                        FormsB1.createFormItem(oForm, formItems, out errorText);
                        if (errorText != null)
                        {
                            return;
                        }

                        //golden arrow
                        formItems = new Dictionary<string, object>();
                        itemName = "CashActL"; //10 characters
                        formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                        formItems.Add("Left", left_e + width_e / 2 - 15);
                        formItems.Add("Top", top);
                        formItems.Add("Height", 14);
                        formItems.Add("UID", itemName);
                        formItems.Add("LinkTo", "CashActE");
                        formItems.Add("LinkedObjectType", objectType);

                        FormsB1.createFormItem(oForm, formItems, out errorText);
                        if (errorText != null)
                        {
                            return;
                        }

                        top = top + height_e + 1;

                        string uniqueID_lf_GLAct = "GLAct_CFL";
                        FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_lf_GLAct);

                        formItems = new Dictionary<string, object>();
                        itemName = "GLActS"; //10 characters
                        formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                        formItems.Add("Left", left_e);
                        formItems.Add("Width", width_s);
                        formItems.Add("Top", top);
                        formItems.Add("Height", height_e);
                        formItems.Add("UID", itemName);
                        formItems.Add("Caption", BDOSResources.getTranslate("GLAccount"));
                        formItems.Add("LinkTo", "GLActE");

                        FormsB1.createFormItem(oForm, formItems, out errorText);
                        if (errorText != null)
                        {
                            return;
                        }

                        formItems = new Dictionary<string, object>();
                        itemName = "GLActE"; //10 characters
                        formItems.Add("isDataSource", true);
                        formItems.Add("DataSource", "UserDataSources");
                        formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                        formItems.Add("Length", 15);
                        formItems.Add("Size", 15);
                        formItems.Add("TableName", "");
                        formItems.Add("Alias", itemName);
                        formItems.Add("Bound", true);
                        formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                        formItems.Add("Left", left_e + width_e / 2 + 2);
                        formItems.Add("Width", width_e);
                        formItems.Add("Top", top);
                        formItems.Add("Height", height_e);
                        formItems.Add("UID", itemName);
                        formItems.Add("DisplayDesc", true);
                        formItems.Add("Enabled", true);
                        formItems.Add("AffectsFormMode", false);
                        formItems.Add("ChooseFromListUID", uniqueID_lf_GLAct);
                        formItems.Add("ChooseFromListAlias", "AcctCode");

                        FormsB1.createFormItem(oForm, formItems, out errorText);
                        if (errorText != null)
                        {
                            return;
                        }

                        //golden arrow
                        formItems = new Dictionary<string, object>();
                        itemName = "GLActL"; //10 characters
                        formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                        formItems.Add("Left", left_e + width_e / 2 - 15);
                        formItems.Add("Top", top);
                        formItems.Add("Height", 14);
                        formItems.Add("UID", itemName);
                        formItems.Add("LinkTo", "GLActE");
                        formItems.Add("LinkedObjectType", objectType);

                        FormsB1.createFormItem(oForm, formItems, out errorText);
                        if (errorText != null)
                        {
                            return;
                        }

                        top = top + height_e + 1;

                        multiSelection = false;
                        objectType = "242"; //CashFlowLineItem
                        string uniqueID_lf_CFW = "CFW_CFL";
                        FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_lf_CFW);

                        formItems = new Dictionary<string, object>();
                        itemName = "CfwOnActS"; //10 characters
                        formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                        formItems.Add("Left", left_e);
                        formItems.Add("Width", width_s);
                        formItems.Add("Top", top);
                        formItems.Add("Height", height_e);
                        formItems.Add("UID", itemName);
                        formItems.Add("Description", BDOSResources.getTranslate("CashFlowLineItemID"));
                        formItems.Add("Caption", BDOSResources.getTranslate("CashFlow"));
                        formItems.Add("LinkTo", "CfwOnActE");

                        FormsB1.createFormItem(oForm, formItems, out errorText);
                        if (errorText != null)
                        {
                            return;
                        }

                        formItems = new Dictionary<string, object>();
                        itemName = "CfwOnActE"; //10 characters
                        formItems.Add("isDataSource", true);
                        formItems.Add("DataSource", "UserDataSources");
                        formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                        formItems.Add("Length", 15);
                        formItems.Add("Size", 15);
                        formItems.Add("TableName", "");
                        formItems.Add("Alias", itemName);
                        formItems.Add("Bound", true);
                        formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                        formItems.Add("Left", left_e + width_e / 2 + 2);
                        formItems.Add("Width", width_e);
                        formItems.Add("Top", top);
                        formItems.Add("Height", height_e);
                        formItems.Add("UID", itemName);
                        formItems.Add("DisplayDesc", true);
                        formItems.Add("Enabled", true);
                        formItems.Add("AffectsFormMode", false);
                        formItems.Add("ChooseFromListUID", uniqueID_lf_CFW);
                        formItems.Add("ChooseFromListAlias", "CFWId");

                        FormsB1.createFormItem(oForm, formItems, out errorText);
                        if (errorText != null)
                        {
                            return;
                        }

                        left_s = 6;
                        left_e = 127;
                    }

                    top = top + 2 * height_e + 1;

                    formItems = new Dictionary<string, object>();
                    itemName = "InvoiceMTR"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
                    formItems.Add("Left", left);
                    formItems.Add("Width", width);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("AffectsFormMode", false);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("InvoiceMTR").Specific));
                    oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single;
                    SAPbouiCOM.Columns oColumns = oMatrix.Columns;
                    SAPbouiCOM.Column oColumn;

                    SAPbouiCOM.DataTable oDataTable;

                    oDataTable = oForm.DataSources.DataTables.Add("InvoiceMTR");
                    oDataTable.Columns.Add("LineNum", SAPbouiCOM.BoFieldsType.ft_Integer, 50); //ინდექსი 
                    oDataTable.Columns.Add("CheckBox", SAPbouiCOM.BoFieldsType.ft_Text, 1); // 
                    oDataTable.Columns.Add("DocEntry", SAPbouiCOM.BoFieldsType.ft_Integer, 50); //ენთრი
                    oDataTable.Columns.Add("DocNum", SAPbouiCOM.BoFieldsType.ft_Integer, 50); //ნომერი
                    oDataTable.Columns.Add("InstallmentID", SAPbouiCOM.BoFieldsType.ft_Integer, 6); //გადარიცხვის ID
                    oDataTable.Columns.Add("LineID", SAPbouiCOM.BoFieldsType.ft_Integer, 11);
                    oDataTable.Columns.Add("DocType", SAPbouiCOM.BoFieldsType.ft_Text, 50); //დოკუმენტის ტიპი
                    oDataTable.Columns.Add("DocDate", SAPbouiCOM.BoFieldsType.ft_Date, 50); //თარიღი
                    oDataTable.Columns.Add("DueDate", SAPbouiCOM.BoFieldsType.ft_Date, 50); //თარიღი
                    oDataTable.Columns.Add("Arrears", SAPbouiCOM.BoFieldsType.ft_Text, 1); //* აჩვენებს, რომ Due Date ნაკლებია ან ტოლი გადახდის თარიღზე
                    oDataTable.Columns.Add("OverdueDays", SAPbouiCOM.BoFieldsType.ft_Integer, 50); //გადახდის თარიღსა და Due Date-ს შორის სხვაობა
                    oDataTable.Columns.Add("Comments", SAPbouiCOM.BoFieldsType.ft_Text, 254); //კომენტარი
                    oDataTable.Columns.Add("Total", SAPbouiCOM.BoFieldsType.ft_Sum); //თანხა
                    oDataTable.Columns.Add("BalanceDue", SAPbouiCOM.BoFieldsType.ft_Sum); //დოკუმენტის დაურეკონსილირებელი თანხა - ვალის ნაშთი
                    oDataTable.Columns.Add("TotalPayment", SAPbouiCOM.BoFieldsType.ft_Sum); //Default - Balance Due
                    oDataTable.Columns.Add("Currency", SAPbouiCOM.BoFieldsType.ft_Text, 50); //დოკუმენტის ვალუტა
                    oDataTable.Columns.Add("TotalPaymentLocal", SAPbouiCOM.BoFieldsType.ft_Sum); //Default - Balance Due

                    
                    oDataTable.Columns.Add("Project", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20);
                    oDataTable.Columns.Add("BlnkAgr", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20);
                    oDataTable.Columns.Add("UseBlaAgRt", SAPbouiCOM.BoFieldsType.ft_Text, 1); 
                    

                    oDataTable.Columns.Add("Test", SAPbouiCOM.BoFieldsType.ft_Text, 100);


                    bool multiSelectionAgr = false;
                    string objectTypeAgr = "1250000025"; //Blanket Agreement
                    string uniqueID_BlnkAgrCFL = "BlnkAgr_CFL";
                    FormsB1.addChooseFromList(oForm, multiSelectionAgr, objectTypeAgr, uniqueID_BlnkAgrCFL);

                    bool multiSelectionPr = false;
                    string objectTypePr = "63"; //Project
                    string uniqueID_lf_Project = "Project_CFLA";
                    FormsB1.addChooseFromList(oForm, multiSelectionPr, objectTypePr, uniqueID_lf_Project);



                    SAPbouiCOM.LinkedButton oLink;

                    string UID = "InvoiceMTR";

                    foreach (SAPbouiCOM.DataColumn column in oDataTable.Columns)
                    {
                        string columnName = column.Name;

                        if (columnName == "LineNum")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = "#";
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind(UID, columnName);
                            oColumn.AffectsFormMode = false;
                        }
                        else if (columnName == "CheckBox")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Selected");
                            oColumn.Editable = true;
                            oColumn.ValOff = "N";
                            oColumn.ValOn = "Y";
                            oColumn.DataBind.Bind(UID, columnName);
                            oColumn.AffectsFormMode = false;
                        }
                        else if (columnName == "DocEntry")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate(columnName);
                            oColumn.Editable = false;
                            oLink = oColumn.ExtendedObject;
                            if(transactionType == OperationTypeFromIntBank.ReturnFromSupplier.ToString())
                            {
                                oLink.LinkedObjectType = "163"; 
                            }
                            else
                            {
                            oLink.LinkedObjectType = "13"; // - A/R Invoice, "14" - A/R Credit Note, A/R Down Payment Request - "203", Journal Entry - "30"
                            }
                            oColumn.DataBind.Bind(UID, columnName);
                            oColumn.AffectsFormMode = false;
                        }
                        else if (columnName == "InstallmentID")
                        {
                            oColumn = oColumns.Add("InstlmntID", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate(columnName);
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind(UID, columnName);
                            oColumn.AffectsFormMode = false;
                        }
                        else if (columnName == "LineID")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate(columnName);
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind(UID, columnName);
                            oColumn.AffectsFormMode = false;
                            oColumn.Visible = false;
                        }

                        else if (columnName == "DocType")
                        {
                            oColumn = oColumns.Add("DocType", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate(columnName);
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind(UID, columnName);
                            oColumn.DisplayDesc = true;
                            oColumn.TitleObject.Sortable = true;
                            oColumn.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
                            oColumn.AffectsFormMode = false;

                            //"13" - A/R Invoice, "14" - A/R Credit Note, A/R Down Payment Request - "203", Journal Entry - "30"

                            oColumn.ValidValues.Add("13", "IN"); //BDOSResources.getTranslate("ARInvoice")
                            oColumn.ValidValues.Add("163", "AP"); //BDOSResources.getTranslate("ARInvoice")
                            oColumn.ValidValues.Add("14", "CN"); //BDOSResources.getTranslate("ARCreditNote")
                            oColumn.ValidValues.Add("203", "DT"); //BDOSResources.getTranslate("ARDownPaymentRequest")
                            oColumn.ValidValues.Add("30", "JE"); //BDOSResources.getTranslate("JournalEntry")
                        }
                        else if (columnName == "Arrears")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = "*";
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind(UID, columnName);
                            oColumn.AffectsFormMode = false;
                        }
                        else if (columnName == "TotalPayment")
                        {
                            oColumn = oColumns.Add("TotalPymnt", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate(columnName);
                            oColumn.Editable = (openFromBlnkAgr ? false : true);
                            oColumn.DataBind.Bind(UID, columnName);
                            oColumn.AffectsFormMode = false;
                        }
                        else if (columnName == "TotalPaymentLocal")
                        {
                            oColumn = oColumns.Add("TotalPmntL", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate(columnName);
                            oColumn.Editable = (openFromBlnkAgr ? true : false);
                            oColumn.DataBind.Bind(UID, columnName);
                            oColumn.AffectsFormMode = false;
                        }
                        else if (columnName == "OverdueDays")
                        {
                            oColumn = oColumns.Add("OverdueDay", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate(columnName);
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind(UID, columnName);
                            oColumn.AffectsFormMode = false;
                        }
                        else if (columnName == "Comments")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("DocumentRemarks");
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind(UID, columnName);
                            oColumn.AffectsFormMode = false;
                        }
                        else if (columnName == "BlnkAgr")
                        {
                            oColumn = oColumns.Add("BlnkAgr", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate(columnName);
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind(UID, columnName);
                            oColumn.ChooseFromListUID = uniqueID_BlnkAgrCFL;
                            oColumn.ChooseFromListAlias = "AbsID";
                            oLink = oColumn.ExtendedObject;
                            oLink.LinkedObjectType = "1250000025";
                        }
                        else if (columnName == "Project")
                        {
                            oColumn = oColumns.Add("Project", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate(columnName);
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind(UID, columnName);
                            oColumn.ChooseFromListUID = uniqueID_lf_Project;
                            oColumn.ChooseFromListAlias = "PrjCode";
                            oLink = oColumn.ExtendedObject;
                            oLink.LinkedObjectType = "63";
                        }
                        else
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate(columnName);
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind(UID, columnName);
                            oColumn.AffectsFormMode = false;
                        }
                    }
                    oMatrix.Clear();
                    oMatrix.LoadFromDataSource();
                    oMatrix.AutoResizeColumns();

                    //ღილაკები
                    top = oForm.ClientHeight - 25;
                    height_e = height_e + 4;
                    width_s = 65;

                    itemName = "1";
                    formItems = new Dictionary<string, object>();
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height_e);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", (openFromBlnkAgr ? BDOSResources.getTranslate("Add") : "OK"));

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left_s = left_s + width_s + 2;

                    itemName = "2";
                    formItems = new Dictionary<string, object>();
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height_e);
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }
                }

                oForm.Visible = true;
                oForm.Select();
            }
            GC.Collect();
        }

        public static void reArrangeFormItems(SAPbouiCOM.Form oForm)
        {
            int left = 6;
            int top = 6;
            int height_e = 15;
            int height = oForm.ClientHeight - top - 8 * height_e - 1 - 30;
            int width = oForm.ClientWidth;

            SAPbouiCOM.Item oItem = oForm.Items.Item("dateS");

            oItem = oForm.Items.Item("dateE");
            oItem.Top = top;
            oItem = oForm.Items.Item("cardCodeS");
            oItem.Top = top;
            oItem = oForm.Items.Item("cardCodeE");
            oItem.Top = top;
            oItem = oForm.Items.Item("cardCodeNE");
            oItem.Top = top;
            oItem = oForm.Items.Item("cardCodeLB");
            oItem.Top = top;

            top = top + 2 * height_e + 1;

            oItem = oForm.Items.Item("amountFBS");
            oItem.Top = top;
            oItem = oForm.Items.Item("amountFBE");
            oItem.Top = top;
            oItem = oForm.Items.Item("currencyCB");
            oItem.Top = top;
            oItem = oForm.Items.Item("docRateINS");
            oItem.Top = top;
            oItem = oForm.Items.Item("docRateINE");
            oItem.Top = top;

            top = top + height_e + 1;

            oItem = oForm.Items.Item("totalAmtDS");
            oItem.Top = top;
            oItem = oForm.Items.Item("totalAmtDE");
            oItem.Top = top;
            if (openFromBlnkAgr)
            {
                oItem = oForm.Items.Item("CashActS");
                oItem.Top = top;
                oItem = oForm.Items.Item("CashActE");
                oItem.Top = top;
                oItem = oForm.Items.Item("CashActL");
                oItem.Top = top;
            }

            top = top + height_e + 1;

            oItem = oForm.Items.Item("dPaymentS");
            oItem.Top = top;
            oItem = oForm.Items.Item("dPaymentE");
            oItem.Top = top;
            if (openFromBlnkAgr)
            {
                oItem = oForm.Items.Item("GLActS");
                oItem.Top = top;
                oItem = oForm.Items.Item("GLActE");
                oItem.Top = top;
                oItem = oForm.Items.Item("GLActL");
                oItem.Top = top;
            }

            top = top + height_e + 1;

            oItem = oForm.Items.Item("invAmtS");
            oItem.Top = top;
            oItem = oForm.Items.Item("invAmtE");
            oItem.Top = top;
            if (openFromBlnkAgr)
            {
                oItem = oForm.Items.Item("CfwOnActS");
                oItem.Top = top;
                oItem = oForm.Items.Item("CfwOnActE");
                oItem.Top = top;
            }

            top = top + height_e + 1;

            oItem = oForm.Items.Item("onAccountS");
            oItem.Top = top;
            oItem = oForm.Items.Item("onAccountE");
            oItem.Top = top;

            top = top + 2 * height_e + 1;

            oItem = oForm.Items.Item("InvoiceMTR");
            oItem.Top = top;
            oItem.Height = height;
            oItem.Width = width;
            oItem.Left = left;

            top = oForm.ClientHeight - 25;

            oItem = oForm.Items.Item("1");
            oItem.Top = top;
            oItem = oForm.Items.Item("2");
            oItem.Top = top;
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

        public static void matrixColumnSetLinkedObjectTypeInvoicesMTR(SAPbouiCOM.Form oForm, SAPbouiCOM.ItemEvent pVal, out string errorText)
        {
            errorText = null;

            try
            {
                if (pVal.ColUID == "DocEntry")
                {
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED & pVal.BeforeAction == true)
                    {
                        SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("InvoiceMTR").Specific));

                        SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("InvoiceMTR");
                        string docType = oDataTable.GetValue("DocType", pVal.Row - 1);

                        SAPbouiCOM.Column oColumn;

                        if (docType == "13")
                        {
                            oColumn = oMatrix.Columns.Item(pVal.ColUID);
                            SAPbouiCOM.LinkedButton oLink = oColumn.ExtendedObject;
                            oLink.LinkedObjectType = docType; //ARInvoice
                        }
                        else if (docType == "14")
                        {
                            oColumn = oMatrix.Columns.Item(pVal.ColUID);
                            SAPbouiCOM.LinkedButton oLink = oColumn.ExtendedObject;
                            oLink.LinkedObjectType = docType; //ARCreditNote
                        }
                        else if (docType == "163")
                        {
                            oColumn = oMatrix.Columns.Item(pVal.ColUID);
                            SAPbouiCOM.LinkedButton oLink = oColumn.ExtendedObject;
                            oLink.LinkedObjectType = docType; //ARCreditNote
                        }
                        else if (docType == "203")
                        {
                            oColumn = oMatrix.Columns.Item(pVal.ColUID);
                            SAPbouiCOM.LinkedButton oLink = oColumn.ExtendedObject;
                            oLink.LinkedObjectType = docType; //ARDownPaymentRequest
                        }
                        else if (docType == "30")
                        {
                            oColumn = oMatrix.Columns.Item(pVal.ColUID);
                            SAPbouiCOM.LinkedButton oLink = oColumn.ExtendedObject;
                            oLink.LinkedObjectType = docType; //Journal Entry
                        }
                    }
                }
                else
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

        //public static void checkUncheckInvoicesMTR(SAPbouiCOM.Form oForm, string checkOperation, out string errorText)
        //{
        //    errorText = null;
        //    try
        //    {
        //        oForm.Freeze(true);
        //        SAPbouiCOM.CheckBox oCheckBox;
        //        SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("InvoiceMTR").Specific));

        //        for (int j = 1; j <= oMatrix.RowCount; j++)
        //        {
        //            oCheckBox = oMatrix.Columns.Item("CheckBox").Cells.Item(j).Specific;
        //            oCheckBox.Checked = (checkOperation == "checkB");
        //        }
        //        oForm.Freeze(false);
        //    }
        //    catch (Exception ex)
        //    {
        //        errorText = ex.Message;
        //    }
        //    finally
        //    {
        //        oForm.Freeze(false);
        //        GC.Collect();
        //    }
        //}

        public static void checkUncheck(SAPbouiCOM.Form oForm, SAPbouiCOM.ItemEvent pVal, out bool bubbleEvent, out string errorText)
        {
            errorText = null;
            bubbleEvent = true;
            //NumberFormatInfo Nfi = new NumberFormatInfo() { NumberDecimalSeparator = "." };
            NumberFormatInfo Nfi = new NumberFormatInfo() { NumberDecimalSeparator = CultureInfo.InvariantCulture.NumberFormat.NumberDecimalSeparator, NumberGroupSeparator = CultureInfo.InvariantCulture.NumberFormat.NumberGroupSeparator };

            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("InvoiceMTR").Specific));
            oForm.Freeze(true);
            oMatrix.FlushToDataSource();
            oForm.Update();
            oForm.Freeze(false);

            try
            {
                if (pVal.BeforeAction == true)
                {

                }
                else
                {
                    SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("InvoiceMTR");

                    invoicesAmount = 0;
                    downPaymentAmount = 0;

                    decimal amountFB = Convert.ToDecimal(oForm.DataSources.UserDataSources.Item("amountFBE").ValueEx, Nfi);
                    string currency = oForm.DataSources.UserDataSources.Item("CurrencyCB").ValueEx;
                    decimal totalAmtD = (invoicesAmount + downPaymentAmount);
                    decimal dPayment = downPaymentAmount;
                    decimal invAmt = invoicesAmount;
                    decimal onAccount = paymentOnAccount;
                    bool unselectedAll = true;

                    for (int i = 0; i < oDataTable.Rows.Count; i++)
                    {
                        bool selected = oDataTable.GetValue("CheckBox", i) == "Y" ? true : false;
                        string docType = oDataTable.GetValue("DocType", i);
                        decimal totalPayment = Convert.ToDecimal(oDataTable.GetValue("TotalPayment", i));
                        decimal totalPaymentLocal = Convert.ToDecimal(oDataTable.GetValue("TotalPaymentLocal", i));
                        totalPaymentLocal = CommonFunctions.roundAmountByGeneralSettings(totalPaymentLocal, "Sum");
                        string docCur = Convert.ToString(oDataTable.GetValue("Currency", i));

                        if (selected == true)
                        {
                            unselectedAll = false;

                            if (currency == Program.MainCurrency)
                            {
                                totalPayment = totalPaymentLocal;
                            }

                            if (docType == "13" || docType == "163" || docType == "14" || docType == "30")
                            {
                                if (selected == true)
                                    invAmt = invAmt + totalPayment;
                                else
                                    invAmt = invAmt - totalPayment;
                            }
                            else if (docType == "203")
                            {
                                if (selected == true)
                                    dPayment = dPayment + totalPayment;
                                else
                                    dPayment = dPayment - totalPayment;
                            }
                        }
                    }
                    totalAmtD = invAmt + dPayment;
                    onAccount = amountFB - totalAmtD;

                    if (amountFB < totalAmtD)
                    {
                        if (openFromBlnkAgr)
                        {
                            errorText = BDOSResources.getTranslate("ReconciliationDifferenceMustBeZeroBeforeReconciling");
                        }
                        else
                        {
                        errorText = BDOSResources.getTranslate("TotalAmountDue") + " " + BDOSResources.getTranslate("Above") + " " + BDOSResources.getTranslate("AmountFromBank");
                        }
                        SAPbouiCOM.CheckBox oCheckBox = ((SAPbouiCOM.CheckBox)oMatrix.Columns.Item("CheckBox").Cells.Item(pVal.Row).Specific);
                        oCheckBox.Checked = false;
                        bubbleEvent = false;
                        return;
                    }

                    if (unselectedAll == false)
                    {
                        oForm.DataSources.UserDataSources.Item("amountFBE").ValueEx = FormsB1.ConvertDecimalToString(Convert.ToDecimal(amountFB));
                        oForm.DataSources.UserDataSources.Item("totalAmtDE").ValueEx = FormsB1.ConvertDecimalToString(Convert.ToDecimal(totalAmtD));
                        oForm.DataSources.UserDataSources.Item("dPaymentE").ValueEx = FormsB1.ConvertDecimalToString(Convert.ToDecimal(dPayment));
                        oForm.DataSources.UserDataSources.Item("invAmtE").ValueEx = FormsB1.ConvertDecimalToString(Convert.ToDecimal(invAmt));

                        if (automaticPaymentInternetBanking)
                        {
                            decimal amountOnAccountE = Convert.ToDecimal(oForm.DataSources.UserDataSources.Item("onAccountE").ValueEx, Nfi);

                            oForm.DataSources.UserDataSources.Item("AddDPAmtE").ValueEx = FormsB1.ConvertDecimalToString(Convert.ToDecimal(onAccount - amountOnAccountE));
                        }
                        else
                        {
                            oForm.DataSources.UserDataSources.Item("onAccountE").ValueEx = FormsB1.ConvertDecimalToString(Convert.ToDecimal(onAccount));
                        }
                    }
                    else
                    {
                        invoicesAmount = 0;
                        downPaymentAmount = 0;

                        /*Converting needs to be corrected !!!! */
                        oForm.DataSources.UserDataSources.Item("amountFBE").ValueEx = amountFB.ToString(Nfi);
                        oForm.DataSources.UserDataSources.Item("totalAmtDE").ValueEx = (invoicesAmount + downPaymentAmount).ToString(Nfi);
                        oForm.DataSources.UserDataSources.Item("dPaymentE").ValueEx = downPaymentAmount.ToString(Nfi);
                        oForm.DataSources.UserDataSources.Item("invAmtE").ValueEx = invoicesAmount.ToString(Nfi);

                        if (automaticPaymentInternetBanking)
                        {
                            decimal amountOnAccountE = Convert.ToDecimal(oForm.DataSources.UserDataSources.Item("onAccountE").ValueEx, Nfi);

                            oForm.DataSources.UserDataSources.Item("AddDPAmtE").ValueEx = (onAccount - amountOnAccountE).ToString(Nfi);
                        }
                        else
                        {
                            //oForm.DataSources.UserDataSources.Item("onAccountE").ValueEx = amountFB.ToString(Nfi);
                            oForm.DataSources.UserDataSources.Item("onAccountE").ValueEx = amountFB.ToString(Nfi);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
        }

        public static void fillInvoicesMTR(  SAPbouiCOM.Form oForm, string blnkAgr, string transactionType, out string errorText)
        {
            errorText = null;

            SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("InvoiceMTR");
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string dateE = oForm.DataSources.UserDataSources.Item("dateE").ValueEx;
            DateTime date = FormsB1.DateFormats(dateE, "yyyyMMdd");
            string cardCodeE = oForm.DataSources.UserDataSources.Item("cardCodeE").ValueEx;

            if (string.IsNullOrEmpty(dateE) || string.IsNullOrEmpty(cardCodeE))
            {
                errorText = BDOSResources.getTranslate("TheFollowingFieldsAreMandatory") + " : \"" + oForm.Items.Item("dateS").Specific.caption + "\", \"" + oForm.Items.Item("cardCodeS").Specific.caption + "\"";
                Program.uiApp.SetStatusBarMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                return;
            }
            string query = "";
            if (transactionType == OperationTypeFromIntBank.ReturnFromSupplier.ToString())
            {
                query = GetInvoicesMTRQuerySupplier(dateE, cardCodeE, blnkAgr);
            }
            else
            {
                query = GetInvoicesMTRQuery(dateE, cardCodeE, blnkAgr);
            }
            oRecordSet.DoQuery(query);

            oDataTable.Rows.Clear();

            try
            {
                int rowIndex = 0;
                int DocEntry;
                int AgrNo;
                string PrjCode;
                int DocNum;
                int InstallmentID;
                string DocType;
                DateTime DueDate;
                string expression;
                DataRow[] foundRows;
                int OverdueDays;
                string UseBlaAgRt;

                while (!oRecordSet.EoF)
                {
                    DocEntry = Convert.ToInt32(oRecordSet.Fields.Item("DocEntry").Value);
                    AgrNo = Convert.ToInt32(oRecordSet.Fields.Item("AgrNo").Value);
                    PrjCode = Convert.ToString(oRecordSet.Fields.Item("Project").Value);
                    DocNum = Convert.ToInt32(oRecordSet.Fields.Item("DocNum").Value);
                    InstallmentID = Convert.ToInt32(oRecordSet.Fields.Item("InstallmentID").Value);
                    DocType = Convert.ToString(oRecordSet.Fields.Item("ObjType").Value);
                    DueDate = oRecordSet.Fields.Item("DueDate").Value;
                    UseBlaAgRt =  Convert.ToString(oRecordSet.Fields.Item("UseBlaAgRt").Value);

                    expression = "DocEntry = '" + DocEntry + "' and DocNum = '" + DocNum + "' and DocType = '" + DocType + "'" + " and LineNumExportMTR = '" + BDOSInternetBanking.CurrentRowExportMTRForDetail + "' and DueDate = '" + DueDate + "' and InstallmentID = '" + InstallmentID + "'";
                    foundRows = BDOSInternetBanking.TableExportMTRForDetail.Select(expression);

                    oDataTable.Rows.Add();

                    if (foundRows.Count() > 0)
                    {
                        OverdueDays = Convert.ToInt32(foundRows[0]["OverdueDays"]);
                        oDataTable.SetValue("LineNum", rowIndex, rowIndex + 1);
                        oDataTable.SetValue("CheckBox", rowIndex, foundRows[0]["CheckBox"]);
                        oDataTable.SetValue("DocEntry", rowIndex, DocEntry);
                        oDataTable.SetValue("DocNum", rowIndex, DocNum);
                        oDataTable.SetValue("InstallmentID", rowIndex, foundRows[0]["InstallmentID"]);
                        oDataTable.SetValue("LineID", rowIndex, foundRows[0]["LineID"]);
                        oDataTable.SetValue("DocType", rowIndex, DocType);
                        oDataTable.SetValue("DocDate", rowIndex, foundRows[0]["DocDate"]);
                        oDataTable.SetValue("DueDate", rowIndex, foundRows[0]["DueDate"]);
                        oDataTable.SetValue("Arrears", rowIndex, OverdueDays >= 0 ? "*" : "");
                        oDataTable.SetValue("OverdueDays", rowIndex, OverdueDays);
                        oDataTable.SetValue("Comments", rowIndex, foundRows[0]["Comments"]);
                        oDataTable.SetValue("Total", rowIndex, Convert.ToDouble(foundRows[0]["Total"]));
                        oDataTable.SetValue("BalanceDue", rowIndex, Convert.ToDouble(foundRows[0]["BalanceDue"]));
                        oDataTable.SetValue("TotalPayment", rowIndex, Convert.ToDouble(foundRows[0]["TotalPayment"]));
                        oDataTable.SetValue("Currency", rowIndex, foundRows[0]["Currency"]);
                        oDataTable.SetValue("TotalPaymentLocal", rowIndex, Convert.ToDouble(foundRows[0]["TotalPaymentLocal"]));
                        oDataTable.SetValue("BlnkAgr", rowIndex, AgrNo);
                        oDataTable.SetValue("Project", rowIndex, PrjCode);
                        oDataTable.SetValue("UseBlaAgRt", rowIndex, UseBlaAgRt);
                        

                    }
                    else
                    {
                        expression = "DocEntry = '" + DocEntry + "' and DocNum = '" + DocNum + "' and DocType = '" + DocType + "' and DueDate = '" + DueDate + "' and InstallmentID = '" + InstallmentID + "'";

                        foundRows = BDOSInternetBanking.TableExportMTRForDetail.Select(expression);

                        decimal OpenAmount = 0;
                        decimal TotalPayment = 0;
                        decimal TotalPaymentLocal = 0;
                        decimal InsTotal = 0;
                        decimal rate = 0;
                        string DocCur = Convert.ToString(oRecordSet.Fields.Item("DocCur").Value);

                        if (string.IsNullOrEmpty(DocCur))
                            DocCur = Program.MainCurrency;

                        TotalPaymentLocal = Convert.ToDecimal(oRecordSet.Fields.Item("OpenAmount").Value);
                        if (Program.MainCurrency == DocCur)
                        {
                            OpenAmount = Convert.ToDecimal(oRecordSet.Fields.Item("OpenAmount").Value);
                            TotalPayment = Convert.ToDecimal(oRecordSet.Fields.Item("OpenAmount").Value);
                            InsTotal = Convert.ToDecimal(oRecordSet.Fields.Item("InsTotal").Value);
                        }
                        else
                        {
                            NumberFormatInfo Nfi = new NumberFormatInfo() { NumberDecimalSeparator = CultureInfo.InvariantCulture.NumberFormat.NumberDecimalSeparator, NumberGroupSeparator = CultureInfo.InvariantCulture.NumberFormat.NumberGroupSeparator };

                            rate = Convert.ToDecimal(oForm.DataSources.UserDataSources.Item("docRateINE").ValueEx, Nfi);

                            OpenAmount = Convert.ToDecimal(oRecordSet.Fields.Item("OpenAmountFC").Value);
                            TotalPayment = Convert.ToDecimal(oRecordSet.Fields.Item("OpenAmountFC").Value);
                            InsTotal = Convert.ToDecimal(oRecordSet.Fields.Item("InsTotalFC").Value);
                            if (rate != 0)
                                TotalPaymentLocal = TotalPayment * rate;
                            else
                            {
                                SAPbobsCOM.SBObob oSBOBob = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);
                                rate = Convert.ToDecimal(oSBOBob.GetCurrencyRate(DocCur, date).Fields.Item("CurrencyRate").Value);
                                TotalPaymentLocal = TotalPayment * rate;
                            }
                        }


                        if (foundRows.Count() > 0)
                        {
                            TotalPayment = TotalPayment - foundRows.Sum(row => row.Field<decimal>("TotalPayment"));
                            TotalPaymentLocal = TotalPaymentLocal - foundRows.Sum(row => row.Field<decimal>("TotalPaymentLocal"));
                        }

                        if (AgrNo != 0 && UseBlaAgRt == "Y" && DocCur != Program.MainCurrency)
                        {
                            rate = BlanketAgreement.GetBlAgremeentCurrencyRate(AgrNo, out DocCur, date);
                            TotalPaymentLocal = TotalPayment * rate;
                        }


                        OverdueDays = Convert.ToInt32(oRecordSet.Fields.Item("OverdueDays").Value);
                        oDataTable.SetValue("LineNum", rowIndex, rowIndex + 1);
                        oDataTable.SetValue("CheckBox", rowIndex, "N");
                        oDataTable.SetValue("DocEntry", rowIndex, DocEntry);
                        oDataTable.SetValue("DocNum", rowIndex, DocNum);
                        oDataTable.SetValue("InstallmentID", rowIndex, oRecordSet.Fields.Item("InstallmentID").Value);
                        oDataTable.SetValue("LineID", rowIndex, oRecordSet.Fields.Item("LineID").Value);
                        oDataTable.SetValue("DocType", rowIndex, DocType);
                        oDataTable.SetValue("DocDate", rowIndex, oRecordSet.Fields.Item("DocDate").Value.ToString("yyyyMMdd") == "18991230" ? "" : oRecordSet.Fields.Item("DocDate").Value.ToString("yyyyMMdd"));
                        oDataTable.SetValue("DueDate", rowIndex, oRecordSet.Fields.Item("DueDate").Value.ToString("yyyyMMdd") == "18991230" ? "" : oRecordSet.Fields.Item("DueDate").Value.ToString("yyyyMMdd"));
                        oDataTable.SetValue("Arrears", rowIndex, OverdueDays >= 0 ? "*" : "");
                        oDataTable.SetValue("OverdueDays", rowIndex, OverdueDays);
                        oDataTable.SetValue("Comments", rowIndex, oRecordSet.Fields.Item("Comments").Value);
                        oDataTable.SetValue("Total", rowIndex, Convert.ToDouble(InsTotal));
                        oDataTable.SetValue("BalanceDue", rowIndex, Convert.ToDouble(OpenAmount));
                        oDataTable.SetValue("TotalPayment", rowIndex, Convert.ToDouble(TotalPayment));
                        oDataTable.SetValue("Currency", rowIndex, DocCur);
                        oDataTable.SetValue("TotalPaymentLocal", rowIndex, Convert.ToDouble(TotalPaymentLocal));
                        oDataTable.SetValue("BlnkAgr", rowIndex, AgrNo);
                        oDataTable.SetValue("Project", rowIndex, PrjCode);
                        oDataTable.SetValue("UseBlaAgRt", rowIndex, UseBlaAgRt);
                        
                    }
                    oRecordSet.MoveNext();
                    rowIndex++;
                }

                SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("InvoiceMTR").Specific));
                oForm.Freeze(true);
                oMatrix.Clear();
                oMatrix.LoadFromDataSource();
                oMatrix.AutoResizeColumns();
                oForm.Update();
                oForm.Freeze(false);
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
            finally
            {
                oForm.Freeze(false);
                Marshal.ReleaseComObject(oRecordSet);
                oRecordSet = null;
            }
        }

        public static string GetInvoicesMTRQuery( string dateE, string cardCodeE, string blnkAgr)
        {
            DateTime date = FormsB1.DateFormats(dateE, "yyyyMMdd");
            string betweenDays = "";

            if (Program.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
            {
                betweenDays = @"DAYS_BETWEEN(T0.""DueDate"",'" + dateE + @"') AS ""OverdueDays"""; //date.ToString("yyyy-MM-dd")
            }
            else
            {

                betweenDays = @"DATEDIFF(DAY, T0.""DueDate"", '" + dateE + @"')  AS ""OverdueDays"" "; //date.ToString("yyyy-MM-dd")
            }

            //WHERE TT0.""isIns"" = 'N' - moixsna filtri Reserve invoice
            string str = @"SELECT
            	 T0.""DocEntry"" AS ""DocEntry"",
                 T0.""AgrNo""  AS ""AgrNo"",
                 T0.""U_UseBlaAgRt""  AS ""UseBlaAgRt"", 
                 T0.""Project""  AS ""Project"",
	             T0.""DocNum"" AS ""DocNum"",
                 T0.""DocCur"" AS ""DocCur"",
            	 T0.""CardCode"" AS ""CardCode"",
            	 T0.""CardName"" AS ""CardName"",
            	 T0.""DocDate"" AS ""DocDate"",
            	 T0.""DueDate"" AS ""DueDate"",
                 T0.""LoanOperationType"",
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
                 TT0.""U_UseBlaAgRt"",
                 TT0.""AgrNo"",
                 TT0.""Project"",
            	 TT0.""DocNum"" AS ""DocNum"",
                 TT0.""DocCur"" AS ""DocCur"",
            	 T3.""CardCode"" AS ""CardCode"",
            	 T3.""CardName"" AS ""CardName"",
            	 TT0.""DocDate"" AS ""DocDate"",
	             TT1.""DueDate"" AS ""DueDate""," +                    
                (CommonFunctions.IsDevelopment() ? @"TT0.""U_BDOSLnOpTp""" : "''") + @" AS ""LoanOperationType"",            
            	 TT0.""ObjType"" AS ""ObjType"",
            	 TT0.""Comments"" AS ""Comments"",
                 TT1.""InstlmntID"" AS ""InstlmntID"", 
                 '0' AS ""LineID"",           	 
            	 SUM(TT1.""InsTotal"" - TT1.""PaidToDate"") AS ""OpenAmount"",
            	 SUM(TT1.""InsTotal"") AS ""InsTotal"",
                 SUM(TT1.""InsTotalFC"" - TT1.""PaidFC"") AS ""OpenAmountFC"",
            	 SUM(TT1.""InsTotalFC"") AS ""InsTotalFC"" 
            	FROM OINV TT0 
            	INNER JOIN INV6 TT1 ON TT0.""DocEntry"" = TT1.""DocEntry"" 
            	INNER JOIN OCRD T3 ON TT0.""CardCode"" = T3.""CardCode""" +
                
                (string.IsNullOrEmpty(blnkAgr) ? "" : @"AND TT0.""AgrNo"" = '" + blnkAgr.Trim() + "'") +
                @"AND TT0.""DocDate"" <= '" + dateE + @"' 
	            AND TT0.""CardCode"" = N'" + cardCodeE + @"' 
            	AND (TT0.""DocStatus"" = 'O' 
            		OR (TT1.""Status"" = 'O' 
            			AND TT0.""CANCELED"" = 'N')) 
            	GROUP BY TT0.""DocEntry"",
                 TT0.""U_UseBlaAgRt"",
                 TT0.""AgrNo"",
                 TT0.""Project"",
            	 TT0.""DocNum"",
                 TT0.""DocCur"",
            	 T3.""CardCode"",
            	 T3.""CardName"",
            	 TT0.""DocDate"",
            	 TT1.""DueDate""," +
                (CommonFunctions.IsDevelopment() ? @" TT0.""U_BDOSLnOpTp"", " : "") + @"
	             TT0.""ObjType"",
            	 TT0.""Comments"",
                 TT1.""InstlmntID"" 
            	UNION ALL SELECT
            	 TT0.""DocEntry"",
                 TT0.""U_UseBlaAgRt"",
                 TT0.""AgrNo"",
                 TT0.""Project"",
            	 TT0.""DocNum"" AS ""DocNum"",
                 TT0.""DocCur"" AS ""DocCur"",
            	 T3.""CardCode"" AS ""CardCode"",
            	 T3.""CardName"" AS ""CardName"",
            	 TT0.""DocDate"" AS ""DocDate"",
            	 TT1.""DueDate"" AS ""DueDate"",
                 '' AS ""LoanOperationType"",
            	 TT0.""ObjType"" AS ""ObjType"",
            	 TT0.""Comments"" AS ""Comments"",
                 TT1.""InstlmntID"" AS ""InstlmntID"",
                 '0' AS ""LineID"",
            	 -SUM(TT1.""InsTotal"" - TT1.""PaidToDate"") AS ""OpenAmount"",
            	 -SUM(TT1.""InsTotal"") AS ""InsTotal"",
                 -SUM(TT1.""InsTotalFC"" - TT1.""PaidFC"") AS ""OpenAmountFC"",
            	 -SUM(TT1.""InsTotalFC"") AS ""InsTotalFC""
            	FROM ORIN TT0 
            	INNER JOIN RIN6 TT1 ON TT0.""DocEntry"" = TT1.""DocEntry"" 
            	INNER JOIN OCRD T3 ON TT0.""CardCode"" = T3.""CardCode"" 
            	WHERE TT0.""isIns"" = 'N'
                AND TT0.""DocDate"" <= '" + dateE + @"' 
            	AND TT0.""CardCode"" = N'" + cardCodeE + @"'
            	AND (TT0.""DocStatus"" = 'O' 
            		OR (TT1.""Status"" = 'O' 
            			AND TT0.""CANCELED"" = 'N')) 
            	GROUP BY TT0.""DocEntry"",
                 TT0.""U_UseBlaAgRt"",
                TT0.""AgrNo"",
                 TT0.""Project"",
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
                 TT0.""U_UseBlaAgRt"",
                 TT0.""AgrNo"",
                 TT0.""Project"",
            	 TT0.""DocNum"" AS ""DocNum"",
                 TT0.""DocCur"" AS ""DocCur"",
            	 T3.""CardCode"" AS ""CardCode"",
            	 T3.""CardName"" AS ""CardName"",
            	 TT0.""DocDate"" AS ""DocDate"",
            	 TT1.""DueDate"" AS ""DueDate"",
                 '' AS ""LoanOperationType"",
            	 TT0.""ObjType"" AS ""ObjType"",
            	 TT0.""Comments"" AS ""Comments"",
                 TT1.""InstlmntID"" AS ""InstlmntID"",
                 '0' AS ""LineID"",
            	 SUM(TT1.""InsTotal"" - TT1.""PaidToDate"") AS ""OpenAmount"",
            	 SUM(TT1.""InsTotal"") AS ""InsTotal"",
                 SUM(TT1.""InsTotalFC"" - TT1.""PaidFC"") AS ""OpenAmountFC"",
            	 SUM(TT1.""InsTotalFC"") AS ""InsTotalFC"" 
            	FROM ODPI TT0 
	            INNER JOIN DPI6 TT1 ON TT0.""DocEntry"" = TT1.""DocEntry"" 
            	INNER JOIN OCRD T3 ON TT0.""CardCode"" = T3.""CardCode"" 
            	WHERE TT0.""isIns"" = 'N' 
            	AND TT0.""DocDate"" <= '" + dateE + @"'            	
            	AND TT0.""CardCode"" = N'" + cardCodeE + @"'
            	AND (TT0.""DocStatus"" = 'O' 
            		OR (TT1.""Status"" = 'O' 
            			AND TT0.""CANCELED"" = 'N')) 
            	GROUP BY TT0.""DocEntry"",
                 TT0.""U_UseBlaAgRt"",
                TT0.""AgrNo"",
                 TT0.""Project"",
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
                 TT0.""U_UseBlaAgRt"",
                 TT0.""AgrNo"",
                 TT0.""Project"",
            	 TT0.""Number"" AS ""DocNum"",
                 TT0.""TransCurr"" AS ""DocCur"",
            	 T3.""CardCode"" AS ""CardCode"",
            	 T3.""CardName"" AS ""CardName"",
            	 TT0.""RefDate"" AS ""DocDate"",
            	 TT1.""DueDate"" AS ""DueDate"",
                 '' AS ""LoanOperationType"",
            	 TT0.""ObjType"" AS ""ObjType"",
            	 TT1.""LineMemo"" AS ""Comments"",
                 '1' AS ""InstlmntID"",
                 TT1.""Line_ID"" AS ""LineID"",
            	 SUM(TT1.""BalDueDeb"" - TT1.""BalDueCred"") AS ""OpenAmount"",
            	 SUM(TT1.""Debit"" - TT1.""Credit"" ) AS ""InsTotal"",
                 SUM(TT1.""BalFcDeb"" - TT1.""BalFcCred"") AS ""OpenAmountFC"",
            	 SUM(TT1.""FCDebit"" - TT1.""FCCredit"" ) AS ""InsTotalFC"" 
            	 FROM OJDT TT0 
            	 INNER JOIN JDT1 TT1 ON TT0.""TransId"" = TT1.""TransId"" 
            	 INNER JOIN OCRD T3 ON TT1.""ShortName"" = T3.""CardCode"" 
            	 WHERE TT0.""RefDate"" <= '" + dateE + @"'
            	 AND TT1.""ShortName"" = N'" + cardCodeE + @"'
            	 AND TT0.""TransType"" IN ('30', '24', '46')  
            	 AND TT0.""BtfStatus"" = 'O' AND (""TT1"".""BalDueDeb"" > '0' OR ""TT1"".""BalDueCred"" > '0') AND ""TT1"".""DprId"" IS NULL
            	 GROUP BY TT0.""TransId"",
                 TT0.""U_UseBlaAgRt"",
                     TT0.""AgrNo"",
                 TT0.""Project"",
	             	 TT0.""Number"",
                     TT0.""TransCurr"",
	             	 T3.""CardCode"",
	             	 T3.""CardName"",
	             	 TT0.""RefDate"",
	             	 TT1.""DueDate"",
	             	 TT0.""ObjType"",
	             	 TT1.""LineMemo"",
                     TT1.""Line_ID"") T0 
            
            ORDER BY T0.""LoanOperationType"" DESC,
                     T0.""DueDate"",
            	 T0.""DocNum""";

            return str;
        }

        public static string GetInvoicesMTRQuerySupplier(string dateE, string cardCodeE, string blnkAgr)
        {
            DateTime date = FormsB1.DateFormats(dateE, "yyyyMMdd");
            string betweenDays = "";

            if (Program.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
            {
                betweenDays = @"DAYS_BETWEEN(T0.""DueDate"",'" + dateE + @"') AS ""OverdueDays"""; //date.ToString("yyyy-MM-dd")
            }
            else
            {

                betweenDays = @"DATEDIFF(DAY, T0.""DueDate"", '" + dateE + @"')  AS ""OverdueDays"" "; //date.ToString("yyyy-MM-dd")
            }

            //WHERE TT0.""isIns"" = 'N' - moixsna filtri Reserve invoice
            string str = @"SELECT
            	 T0.""DocEntry"" AS ""DocEntry"",
                 T0.""U_UseBlaAgRt"" as ""UseBlaAgRt"",                 
                 T0.""AgrNo""  AS ""AgrNo"",
                 
                 T0.""Project""  AS ""Project"",
	             T0.""DocNum"" AS ""DocNum"",
                 T0.""DocCur"" AS ""DocCur"",
            	 T0.""CardCode"" AS ""CardCode"",
            	 T0.""CardName"" AS ""CardName"",
            	 T0.""DocDate"" AS ""DocDate"",
            	 T0.""DueDate"" AS ""DueDate"",
                 T0.""LoanOperationType"",
            	 T0.""OpenAmount"" AS ""OpenAmount"",
            	 T0.""InsTotal""  AS ""InsTotal"",
            	 T0.""OpenAmountFC""  AS ""OpenAmountFC"",
            	 T0.""InsTotalFC""  AS ""InsTotalFC"",
            	 T0.""ObjType"" AS ""ObjType"",
            	 T0.""Comments"" AS ""Comments"",
                 T0.""InstlmntID"" AS ""InstallmentID"",
                 T0.""LineID"" AS ""LineID""," + betweenDays + @"             
            FROM ( SELECT
            	 TT0.""DocEntry"",
                 TT0.""U_UseBlaAgRt"",
                 TT0.""AgrNo"",
                 TT0.""Project"",
            	 TT0.""DocNum"" AS ""DocNum"",
                 TT0.""DocCur"" AS ""DocCur"",
            	 T3.""CardCode"" AS ""CardCode"",
            	 T3.""CardName"" AS ""CardName"",
            	 TT0.""DocDate"" AS ""DocDate"",
	             TT1.""DueDate"" AS ""DueDate""," +
                (CommonFunctions.IsDevelopment() ? @"TT0.""U_BDOSLnOpTp""" : "''") + @" AS ""LoanOperationType"",            
            	 TT0.""ObjType"" AS ""ObjType"",
            	 TT0.""Comments"" AS ""Comments"",
                 TT1.""InstlmntID"" AS ""InstlmntID"", 
                 '0' AS ""LineID"",           	 
            	 SUM((TT1.""InsTotal"" - TT1.""PaidToDate"")*-1 ) AS ""OpenAmount"",
            	 SUM(TT1.""InsTotal""*-1) AS ""InsTotal"",
                 SUM((TT1.""InsTotalFC"" - TT1.""PaidFC"")*-1) AS ""OpenAmountFC"",
            	 SUM(TT1.""InsTotalFC""*-1) AS ""InsTotalFC"" 
            	FROM OCPI TT0 
            	INNER JOIN CPI6 TT1 ON TT0.""DocEntry"" = TT1.""DocEntry"" 
            	INNER JOIN OCRD T3 ON TT0.""CardCode"" = T3.""CardCode""" +

                (string.IsNullOrEmpty(blnkAgr) ? "" : @"AND TT0.""AgrNo"" = '" + blnkAgr.Trim() + "'") +
                @"AND TT0.""DocDate"" <= '" + dateE + @"' 
	            AND TT0.""CardCode"" = N'" + cardCodeE + @"' 
            	AND (TT0.""DocStatus"" = 'O' 
            		OR (TT1.""Status"" = 'O' 
            			AND TT0.""CANCELED"" = 'N')) 
            	GROUP BY TT0.""DocEntry"",
                 TT0.""U_UseBlaAgRt"",
                 TT0.""AgrNo"",
                 TT0.""Project"",
            	 TT0.""DocNum"",
                 TT0.""DocCur"",
            	 T3.""CardCode"",
            	 T3.""CardName"",
            	 TT0.""DocDate"",
            	 TT1.""DueDate""," +
                (CommonFunctions.IsDevelopment() ? @" TT0.""U_BDOSLnOpTp"", " : "") + @"
	             TT0.""ObjType"",
            	 TT0.""Comments"",
                 TT1.""InstlmntID"" 
            	 
                UNION ALL SELECT
            	 TT0.""TransId"" AS ""DocEntry"",
                 TT0.""U_UseBlaAgRt"",
                 TT0.""AgrNo"",
                 TT0.""Project"",
            	 TT0.""Number"" AS ""DocNum"",
                 TT0.""TransCurr"" AS ""DocCur"",
            	 T3.""CardCode"" AS ""CardCode"",
            	 T3.""CardName"" AS ""CardName"",
            	 TT0.""RefDate"" AS ""DocDate"",
            	 TT1.""DueDate"" AS ""DueDate"",
                 '' AS ""LoanOperationType"",
            	 TT0.""ObjType"" AS ""ObjType"",
            	 TT1.""LineMemo"" AS ""Comments"",
                 '1' AS ""InstlmntID"",
                 TT1.""Line_ID"" AS ""LineID"",
            	 SUM(TT1.""BalDueDeb"" - TT1.""BalDueCred"") AS ""OpenAmount"",
            	 SUM(TT1.""Debit"" - TT1.""Credit"" ) AS ""InsTotal"",
                 SUM(TT1.""BalFcDeb"" - TT1.""BalFcCred"") AS ""OpenAmountFC"",
            	 SUM(TT1.""FCDebit"" - TT1.""FCCredit"" ) AS ""InsTotalFC"" 
            	 FROM OJDT TT0 
            	 INNER JOIN JDT1 TT1 ON TT0.""TransId"" = TT1.""TransId"" 
            	 INNER JOIN OCRD T3 ON TT1.""ShortName"" = T3.""CardCode"" 
            	 WHERE TT0.""RefDate"" <= '" + dateE + @"'
            	 AND TT1.""ShortName"" = N'" + cardCodeE + @"'
            	 AND TT0.""TransType"" IN ('30', '24', '46')  
            	 AND TT0.""BtfStatus"" = 'O' AND (""TT1"".""BalDueDeb"" > '0' OR ""TT1"".""BalDueCred"" > '0') AND ""TT1"".""DprId"" IS NULL
            	 GROUP BY TT0.""TransId"",
                 TT0.""U_UseBlaAgRt"",
                     TT0.""AgrNo"",
                 TT0.""Project"",
	             	 TT0.""Number"",
                     TT0.""TransCurr"",
	             	 T3.""CardCode"",
	             	 T3.""CardName"",
	             	 TT0.""RefDate"",
	             	 TT1.""DueDate"",
	             	 TT0.""ObjType"",
	             	 TT1.""LineMemo"",
                     TT1.""Line_ID"") T0 
            
            WHERE (ROUND (T0.""OpenAmount"", 2) <> '0' OR ROUND (T0.""OpenAmountFC"", 2) <> '0')
            
            ORDER BY T0.""LoanOperationType"" DESC,
                     T0.""DueDate"",
            	 T0.""DocNum""";

            return str;
        }

        static void countPaymentForRow(SAPbouiCOM.Form oForm, SAPbouiCOM.DataTable oDataTable, int row, string colUID)
        {
            NumberFormatInfo Nfi = new NumberFormatInfo() { NumberDecimalSeparator = CultureInfo.InvariantCulture.NumberFormat.NumberDecimalSeparator, NumberGroupSeparator = CultureInfo.InvariantCulture.NumberFormat.NumberGroupSeparator };

            decimal totalPayment = Convert.ToDecimal(oDataTable.GetValue("TotalPayment", row));
            decimal totalPaymentLocal = Convert.ToDecimal(oDataTable.GetValue("TotalPaymentLocal", row));
            string blnkAgr = oDataTable.GetValue("BlnkAgr", row);
            string docCur = Convert.ToString(oDataTable.GetValue("Currency", row));
            decimal rate = Convert.ToDecimal(oForm.DataSources.UserDataSources.Item("docRateINE").ValueEx, Nfi);
            string UseBlaAgRt = oDataTable.GetValue("UseBlaAgRt", row);


            if (docCur != Program.MainCurrency)
            {
                string dateE = oForm.DataSources.UserDataSources.Item("dateE").ValueEx;
                DateTime docDate = FormsB1.DateFormats(dateE, "yyyyMMdd");
                decimal docRate;
                SAPbobsCOM.SBObob oSBOBob = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);

                if (rate == 0)
                    docRate = Convert.ToDecimal(oSBOBob.GetCurrencyRate(docCur, docDate).Fields.Item("CurrencyRate").Value);
                else
                    docRate = rate;

                if ( blnkAgr!="" && blnkAgr!="0" && UseBlaAgRt=="Y")
                    docRate = BlanketAgreement.GetBlAgremeentCurrencyRate(Convert.ToInt32(blnkAgr), out docCur, docDate);


                if (colUID == "TotalPymnt")
                {
                oDataTable.SetValue("TotalPaymentLocal", row, Convert.ToDouble(totalPayment * docRate));
            }
            else
            {
                    oDataTable.SetValue("TotalPayment", row, Convert.ToDouble(totalPaymentLocal / docRate));
                }
            }
            else
            {
                if (colUID == "TotalPymnt")
                {
                oDataTable.SetValue("TotalPaymentLocal", row, Convert.ToDouble(totalPayment));
            }
                else
                {
                    oDataTable.SetValue("TotalPayment", row, Convert.ToDouble(totalPaymentLocal));
                }
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
                SAPbouiCOM.UserDataSources oUserDataSources = oForm.DataSources.UserDataSources;

                if (beforeAction == false)
                {
                    SAPbouiCOM.DataTable oDataTableSelectedObjects = null;
                    oDataTableSelectedObjects = oCFLEvento.SelectedObjects;

                    if (oDataTableSelectedObjects != null)
                    {
                        if (sCFL_ID == "CashAct_CFL")
                        {
                            string AcctCode = oDataTableSelectedObjects.GetValue("AcctCode", 0);
                            oUserDataSources.Item("CashActE").ValueEx = AcctCode;
                        }
                        else if (sCFL_ID == "GLAct_CFL")
                        {
                            string AcctCode = oDataTableSelectedObjects.GetValue("AcctCode", 0);
                            oUserDataSources.Item("GLActE").ValueEx = AcctCode;
                        }
                        else if (sCFL_ID == "CFW_CFL")
                        {
                            string CFWId = oDataTableSelectedObjects.GetValue("CFWId", 0).ToString();
                            oUserDataSources.Item("CfwOnActE").ValueEx = CFWId;
                        }
                    }
                }
                else
                {
                    if (sCFL_ID == "CashAct_CFL")
                    {
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
                    else if (sCFL_ID == "GLAct_CFL")
                    {
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
                        oCon.CondVal = "Y";

                        oCFL.SetConditions(oCons);
                    }
                    else if (sCFL_ID == "CFW_CFL")
                    {
                        SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();

                        SAPbouiCOM.Condition oCon = oCons.Add();
                        oCon.Alias = "Postable"; 
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = "Y";

                        oCFL.SetConditions(oCons);
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

        public static void uiApp_ItemEvent(  string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_CLOSE && pVal.BeforeAction == true && pVal.InnerEvent == false)
            {
                int answer = Program.uiApp.MessageBox(BDOSResources.getTranslate("DoYouWantToClose") + "?", 1, BDOSResources.getTranslate("Yes"), BDOSResources.getTranslate("No"), "");

                if (answer != 1)
                {
                    BubbleEvent = false;
                }
            }

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_RESIZE & pVal.BeforeAction == false)
                {
                    oForm.Freeze(true);
                    resizeForm( oForm, out errorText);
                    oForm.Freeze(false);
                }

                if (pVal.ItemUID == "onAccountE" && pVal.BeforeAction == false && pVal.ItemChanged == true)
                {
                    NumberFormatInfo Nfi = new NumberFormatInfo() { NumberDecimalSeparator = CultureInfo.InvariantCulture.NumberFormat.NumberDecimalSeparator, NumberGroupSeparator = CultureInfo.InvariantCulture.NumberFormat.NumberGroupSeparator };


                    decimal amountFBE = Convert.ToDecimal(oForm.DataSources.UserDataSources.Item("amountFBE").ValueEx, Nfi);
                    decimal totalAmtDe = Convert.ToDecimal(oForm.DataSources.UserDataSources.Item("totalAmtDe").ValueEx, Nfi);
                    decimal onAccount = Convert.ToDecimal(oForm.DataSources.UserDataSources.Item("onAccountE").ValueEx, Nfi);

                    oForm.DataSources.UserDataSources.Item("AddDPAmtE").ValueEx = FormsB1.ConvertDecimalToString(Convert.ToDecimal(amountFBE - totalAmtDe - onAccount));
                }

                if (pVal.ItemUID == "docRateINE" && pVal.BeforeAction == false && pVal.ItemChanged == true)
                {
                    NumberFormatInfo Nfi = new NumberFormatInfo() { NumberDecimalSeparator = CultureInfo.InvariantCulture.NumberFormat.NumberDecimalSeparator, NumberGroupSeparator = CultureInfo.InvariantCulture.NumberFormat.NumberGroupSeparator };

                    SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("InvoiceMTR").Specific));
                    oForm.Freeze(true);
                    oMatrix.FlushToDataSource();
                    oForm.Update();
                    oForm.Freeze(false);

                    try
                    {
                        SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("InvoiceMTR");

                        for (int i = 0; i < oDataTable.Rows.Count; i++)
                        {
                            countPaymentForRow(oForm, oDataTable, i, "TotalPymnt");
                            oDataTable.SetValue("CheckBox", i, "N");
                        }
                        oForm.Freeze(true);
                        oMatrix.LoadFromDataSource();
                        oForm.Update();
                        oForm.Freeze(false);
                    }
                    catch
                    {

                    }
                    checkUncheck(oForm, pVal, out BubbleEvent, out errorText);
                }

                if (pVal.ItemUID == "amountFBE" && pVal.BeforeAction == false && pVal.ItemChanged == true)
                {
                    decimal amountFB = Convert.ToDecimal(oForm.DataSources.UserDataSources.Item("amountFBE").ValueEx, CultureInfo.InvariantCulture);
                    decimal amountINV = Convert.ToDecimal(oForm.DataSources.UserDataSources.Item("invAmtE").ValueEx, CultureInfo.InvariantCulture);
                    if (amountFB < amountINV)
                    {
                        Program.uiApp.SetStatusBarMessage(BDOSResources.getTranslate("UncheckInvoices"), SAPbouiCOM.BoMessageTime.bmt_Short, true);
                        decimal amountDP = Convert.ToDecimal(oForm.DataSources.UserDataSources.Item("AddDPAmtE").ValueEx, CultureInfo.InvariantCulture);
                        oForm.DataSources.UserDataSources.Item("amountFBE").ValueEx = (amountINV + amountDP).ToString(CultureInfo.InvariantCulture);
                    }
                    checkUncheck(oForm, pVal, out BubbleEvent, out errorText);
                }

                if (pVal.ItemUID == "InvoiceMTR")
                {
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED)
                        matrixColumnSetLinkedObjectTypeInvoicesMTR(oForm, pVal, out errorText);
                    else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.ColUID == "CheckBox")
                        checkUncheck(oForm, pVal, out BubbleEvent, out errorText);
                    else if ((pVal.ColUID == "TotalPymnt") && pVal.BeforeAction == false && pVal.ItemChanged == true)
                    {
                        NumberFormatInfo Nfi = new NumberFormatInfo() { NumberDecimalSeparator = CultureInfo.InvariantCulture.NumberFormat.NumberDecimalSeparator, NumberGroupSeparator = CultureInfo.InvariantCulture.NumberFormat.NumberGroupSeparator };

                        SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("InvoiceMTR").Specific));
                        oForm.Freeze(true);
                        oMatrix.FlushToDataSource();
                        oForm.Update();
                        oForm.Freeze(false);

                        try
                        {
                            SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("InvoiceMTR");
                            decimal totalPayment = Convert.ToDecimal(oDataTable.GetValue("TotalPayment", pVal.Row - 1));
                            decimal balanceDue = Convert.ToDecimal(oDataTable.GetValue("BalanceDue", pVal.Row - 1));

                            if (totalPayment > balanceDue)
                            {
                                errorText = BDOSResources.getTranslate("TotalPayment") + " " + BDOSResources.getTranslate("Above") + " " + BDOSResources.getTranslate("BalanceDue");
                                Program.uiApp.SetStatusBarMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                oDataTable.SetValue("TotalPayment", pVal.Row - 1, 0);
                                oDataTable.SetValue("TotalPaymentLocal", pVal.Row - 1, 0);
                            }
                            else
                            {
                                countPaymentForRow(oForm, oDataTable, pVal.Row - 1, "TotalPymnt");
                            }

                            oForm.Freeze(true);
                            oMatrix.LoadFromDataSource();
                            oForm.Update();
                            oForm.Freeze(false);
                        }
                        catch
                        {

                        }

                        checkUncheck(oForm, pVal, out BubbleEvent, out errorText);
                    }
                    else if ((pVal.ColUID == "TotalPmntL") && pVal.BeforeAction == false && pVal.ItemChanged == true)
                    {
                        NumberFormatInfo Nfi = new NumberFormatInfo() { NumberDecimalSeparator = CultureInfo.InvariantCulture.NumberFormat.NumberDecimalSeparator, NumberGroupSeparator = CultureInfo.InvariantCulture.NumberFormat.NumberGroupSeparator };

                        SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("InvoiceMTR").Specific));
                        oForm.Freeze(true);
                        oMatrix.FlushToDataSource();
                        oForm.Update();
                        oForm.Freeze(false);

                        try
                        {
                            SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("InvoiceMTR");

                            countPaymentForRow(oForm, oDataTable, pVal.Row - 1, "TotalPmntL");

                            decimal totalPayment = Convert.ToDecimal(oDataTable.GetValue("TotalPayment", pVal.Row - 1));
                            decimal balanceDue = Convert.ToDecimal(oDataTable.GetValue("BalanceDue", pVal.Row - 1));

                            if (totalPayment > balanceDue)
                            {
                                errorText = BDOSResources.getTranslate("TotalPayment") + " " + BDOSResources.getTranslate("Above") + " " + BDOSResources.getTranslate("BalanceDue");
                                Program.uiApp.SetStatusBarMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                oDataTable.SetValue("TotalPayment", pVal.Row - 1, 0);
                                oDataTable.SetValue("TotalPaymentLocal", pVal.Row - 1, 0);
                            }

                            oForm.Freeze(true);
                            oMatrix.LoadFromDataSource();
                            oForm.Update();
                            oForm.Freeze(false);
                        }
                        catch
                        {

                        }

                        checkUncheck(oForm, pVal, out BubbleEvent, out errorText);
                    }
                    if (string.IsNullOrEmpty(errorText) == false)
                    {
                        Program.uiApp.SetStatusBarMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    }
                }
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK)
                {
                    if (pVal.BeforeAction == true)
                    {
                        if (pVal.ItemUID == "1") //OK
                        {
                            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("InvoiceMTR").Specific));
                            oMatrix.FlushToDataSource();
                            string checkBox;
                            SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("InvoiceMTR");

                            string expression = "LineNumExportMTR = '" + BDOSInternetBanking.CurrentRowExportMTRForDetail + "'";
                            DataRow[] foundRows;
                            foundRows = BDOSInternetBanking.TableExportMTRForDetail.Select(expression);
                            if (foundRows.Count() > 0)
                            {
                                for (int i = 0; i < foundRows.Count(); i++)
                                {
                                    foundRows[i].Delete();
                                }
                                BDOSInternetBanking.TableExportMTRForDetail.AcceptChanges();
                            }

                            for (int i = 0; i < oDataTable.Rows.Count; i++)
                            {
                                checkBox = oDataTable.GetValue("CheckBox", i);
                                if (checkBox == "Y")
                                {
                                    DataRow dataRow = BDOSInternetBanking.TableExportMTRForDetail.Rows.Add();
                                    dataRow["LineNumExportMTR"] = BDOSInternetBanking.CurrentRowExportMTRForDetail;
                                    dataRow["LineNum"] = Convert.ToInt32(oDataTable.GetValue("LineNum", i));
                                    dataRow["CheckBox"] = oDataTable.GetValue("CheckBox", i).ToString();
                                    dataRow["DocEntry"] = Convert.ToInt32(oDataTable.GetValue("DocEntry", i));
                                    dataRow["DocNum"] = Convert.ToInt32(oDataTable.GetValue("DocNum", i));
                                    dataRow["InstallmentID"] = Convert.ToInt32(oDataTable.GetValue("InstallmentID", i));
                                    dataRow["LineID"] = Convert.ToInt32(oDataTable.GetValue("LineID", i));
                                    dataRow["DocType"] = oDataTable.GetValue("DocType", i).ToString();
                                    dataRow["DocDate"] = Convert.ToDateTime(oDataTable.GetValue("DocDate", i));
                                    dataRow["DueDate"] = Convert.ToDateTime(oDataTable.GetValue("DueDate", i));
                                    dataRow["Arrears"] = oDataTable.GetValue("Arrears", i).ToString();
                                    dataRow["OverdueDays"] = Convert.ToInt32(oDataTable.GetValue("OverdueDays", i));
                                    dataRow["Comments"] = oDataTable.GetValue("Comments", i).ToString();
                                    dataRow["Total"] = Convert.ToDecimal(oDataTable.GetValue("Total", i));
                                    dataRow["BalanceDue"] = Convert.ToDecimal(oDataTable.GetValue("BalanceDue", i));
                                    dataRow["TotalPayment"] = Convert.ToDecimal(oDataTable.GetValue("TotalPayment", i));
                                    dataRow["Currency"] = oDataTable.GetValue("Currency", i).ToString();
                                    dataRow["TotalPaymentLocal"] = Convert.ToDecimal(oDataTable.GetValue("TotalPaymentLocal", i));
                                    dataRow["Project"] = oDataTable.GetValue("Project", i).ToString();
                                    dataRow["BlnkAgr"] = oDataTable.GetValue("BlnkAgr", i).ToString();
                                    dataRow["UseBlaAgRt"] = oDataTable.GetValue("UseBlaAgRt", i);
                                    

                                }
                            }

                            NumberFormatInfo Nfi = new NumberFormatInfo() { NumberDecimalSeparator = CultureInfo.InvariantCulture.NumberFormat.NumberDecimalSeparator, NumberGroupSeparator = CultureInfo.InvariantCulture.NumberFormat.NumberGroupSeparator };
                            //decimal amountFB = Convert.ToDecimal(oForm.DataSources.UserDataSources.Item("amountFBE").ValueEx, Nfi);
                            //decimal totalAmtD = Convert.ToDecimal(oForm.DataSources.UserDataSources.Item("totalAmtDE").ValueEx, Nfi);
                            downPaymentAmount = Convert.ToDecimal(oForm.DataSources.UserDataSources.Item("dPaymentE").ValueEx, Nfi);
                            invoicesAmount = Convert.ToDecimal(oForm.DataSources.UserDataSources.Item("invAmtE").ValueEx, Nfi);
                            paymentOnAccount = Convert.ToDecimal(oForm.DataSources.UserDataSources.Item("onAccountE").ValueEx, Nfi);

                            if (automaticPaymentInternetBanking)
                            {
                                addDownPaymentAmount = Convert.ToDecimal(oForm.DataSources.UserDataSources.Item("AddDPAmtE").ValueEx, Nfi);
                            }

                            docRateIN = Convert.ToDecimal(oForm.DataSources.UserDataSources.Item("docRateINE").ValueEx, Nfi);
                            //ბლენქით ეგრიმენთიდან
                            if (openFromBlnkAgr)
                            {
                                string dateE = oForm.DataSources.UserDataSources.Item("dateE").ValueEx;
                                DateTime date = FormsB1.DateFormats(dateE, "yyyyMMdd");
                                string cashAct = oForm.DataSources.UserDataSources.Item("CashActE").ValueEx;

                                BlanketAgreement.TableForPaymentDetail.SetValue("DocumentDate", 0, date);
                                BlanketAgreement.TableForPaymentDetail.SetValue("PaymentOnAccount", 0, Convert.ToDouble(paymentOnAccount));
                                BlanketAgreement.TableForPaymentDetail.SetValue("Amount", 0, Convert.ToDouble(downPaymentAmount + invoicesAmount + paymentOnAccount + addDownPaymentAmount));
                                BlanketAgreement.TableForPaymentDetail.SetValue("InvoicesAmount", 0, Convert.ToDouble(invoicesAmount));
                                BlanketAgreement.TableForPaymentDetail.SetValue("AddDownPaymentAmount", 0, Convert.ToDouble(addDownPaymentAmount));
                                BlanketAgreement.TableForPaymentDetail.SetValue("GLAccountCode", 0, oForm.DataSources.UserDataSources.Item("GLActE").ValueEx);
                                BlanketAgreement.TableForPaymentDetail.SetValue("CashAccount", 0, cashAct);
                                BlanketAgreement.TableForPaymentDetail.SetValue("CashFlowLineItemID", 0, oForm.DataSources.UserDataSources.Item("CfwOnActE").ValueEx);

                                int docEntry;
                                int docNum;
                                string info = IncomingPayment.createDocumentTransferFromBPCashType(BlanketAgreement.TableForPaymentDetail, oForm, 0, out docEntry, out docNum, out errorText);
                                if (String.IsNullOrEmpty(errorText) == false)
                                {
                                    Program.uiApp.SetStatusBarMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                    BubbleEvent = false;
                                }
                                else if (String.IsNullOrEmpty(info) == false)
                                {
                                    Program.uiApp.SetStatusBarMessage(info, SAPbouiCOM.BoMessageTime.bmt_Short, false);
                                }
                            }
                            else
                            {
                            SAPbouiCOM.Matrix oMatrixBDOSInternetBanking = ((SAPbouiCOM.Matrix)(oFormBDOSInternetBanking.Items.Item("exportMTR").Specific));

                            int selectedRow = oMatrixBDOSInternetBanking.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder);

                            BDOSInternetBanking.updateExportMTRRow(oFormBDOSInternetBanking);

                            oMatrixBDOSInternetBanking.SelectRow(selectedRow,true,false);

                            oForm.Close();
                        }
                    }
                    }
                    else
                    {

                    }
                }

                if ((pVal.ItemUID == "CashActE" || pVal.ItemUID == "GLActE" || pVal.ItemUID == "CfwOnActE") & pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                {
                    SAPbouiCOM.IChooseFromListEvent oCFLEvento = null;
                    oCFLEvento = ((SAPbouiCOM.IChooseFromListEvent)(pVal));

                    chooseFromList(oForm, oCFLEvento, pVal, out errorText);
                }

            }
        }
    }
}
