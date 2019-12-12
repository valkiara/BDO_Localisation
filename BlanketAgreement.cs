using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Data;

namespace BDO_Localisation_AddOn
{
    static partial class BlanketAgreement
    {
        public static SAPbouiCOM.DataTable TableForPaymentDetail;

        public static void createUserFields(out string errorText)
        {
            errorText = null;

            Dictionary<string, object> fieldskeysMap;

            //Gross Price
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSMinRat");
            fieldskeysMap.Add("TableName", "OOAT");
            fieldskeysMap.Add("Description", "Min rate");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Rate);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSMaxRat");
            fieldskeysMap.Add("TableName", "OOAT");
            fieldskeysMap.Add("Description", "Max rate");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Rate);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            //Gross Price
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSGrsPrc");
            fieldskeysMap.Add("TableName", "OAT1");
            fieldskeysMap.Add("Description", "Gross price");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Price);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSGrsAmt");
            fieldskeysMap.Add("TableName", "OAT1");
            fieldskeysMap.Add("Description", "Gross amount");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            GC.Collect();
        }

        public static void createFormItems(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            //გადახდის ვიზარდისთვის გვჭირდება ეს ცხრილი
            TableForPaymentDetail = oForm.DataSources.DataTables.Add("TableForPaymentDetail");
            TableForPaymentDetail.Columns.Add("CheckBox", SAPbouiCOM.BoFieldsType.ft_Text, 1);
            TableForPaymentDetail.Columns.Add("LineNum", SAPbouiCOM.BoFieldsType.ft_Integer, 100);
            TableForPaymentDetail.Columns.Add("DocumentDate", SAPbouiCOM.BoFieldsType.ft_Date);
            TableForPaymentDetail.Columns.Add("GLAccountCode", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 15);
            TableForPaymentDetail.Columns.Add("CashAccount", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 15);
            TableForPaymentDetail.Columns.Add("Project", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20);
            TableForPaymentDetail.Columns.Add("CashFlowLineItemID", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 11);
            TableForPaymentDetail.Columns.Add("Currency", SAPbouiCOM.BoFieldsType.ft_Text, 100);
            TableForPaymentDetail.Columns.Add("PartnerCurrency", SAPbouiCOM.BoFieldsType.ft_Text, 100);
            TableForPaymentDetail.Columns.Add("BlnkAgr", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20);
            TableForPaymentDetail.Columns.Add("CardCode", SAPbouiCOM.BoFieldsType.ft_Text, 100);
            TableForPaymentDetail.Columns.Add("CardName", SAPbouiCOM.BoFieldsType.ft_Text, 100);
            TableForPaymentDetail.Columns.Add("BPCurrency", SAPbouiCOM.BoFieldsType.ft_Text, 100);
            TableForPaymentDetail.Columns.Add("DocRateIN", SAPbouiCOM.BoFieldsType.ft_Rate, 100);
            TableForPaymentDetail.Columns.Add("Amount", SAPbouiCOM.BoFieldsType.ft_Sum, 100);
            TableForPaymentDetail.Columns.Add("InvoicesAmount", SAPbouiCOM.BoFieldsType.ft_Sum, 100);
            TableForPaymentDetail.Columns.Add("AddDownPaymentAmount", SAPbouiCOM.BoFieldsType.ft_Sum, 100);
            TableForPaymentDetail.Columns.Add("PaymentOnAccount", SAPbouiCOM.BoFieldsType.ft_Sum, 100);
            //გადახდის ვიზარდისთვის გვჭირდება ეს ცხრილი

            SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("1250000045").Specific;
            SAPbouiCOM.Columns oColumns = oMatrix.Columns;
            SAPbouiCOM.Column oColumn = oMatrix.Columns.Item("U_BDOSGrsAmt");
            oColumn.Editable = false;

            Dictionary<string, object> formItems;
            string itemName = "";
            SAPbouiCOM.Item oItem = oForm.Items.Item("234000001");

            formItems = new Dictionary<string, object>();
            itemName = "BDOSPay"; //10 characters
            formItems.Add("Caption", BDOSResources.getTranslate("Payment"));
            formItems.Add("Size", 8);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            formItems.Add("Left", oItem.Left - oItem.Width - 5);
            formItems.Add("Width", oItem.Width);
            formItems.Add("Top", oItem.Top);
            formItems.Add("Height", oItem.Height);
            formItems.Add("UID", itemName);
            formItems.Add("Enabled", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            //------------------

            oItem = oForm.Items.Item("1250000035");
            int left = oItem.Left;
            oItem = oForm.Items.Item("1720000061");

            formItems = new Dictionary<string, object>();
            itemName = "RateRange";
            formItems.Add("Size", 20);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left);
            formItems.Add("Width", 60);
            formItems.Add("Top", oItem.Top);
            formItems.Add("Caption", BDOSResources.getTranslate("RateRange"));
            formItems.Add("UID", itemName);
            formItems.Add("FromPane", 0);
            formItems.Add("ToPane", 0);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            oItem = oForm.Items.Item("RateRange");

            formItems = new Dictionary<string, object>();
            itemName = "MinRate"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OOAT");
            formItems.Add("Alias", "U_BDOSMinRat");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", oItem.Left + oItem.Width + 10);
            formItems.Add("Width", 50);
            formItems.Add("Top", oItem.Top);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);


            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }
            oItem = oForm.Items.Item("MinRate");

            formItems = new Dictionary<string, object>();
            itemName = "Hyphen";
            formItems.Add("Size", 20);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", oItem.Left + oItem.Width + 5);
            formItems.Add("Width", 10);
            formItems.Add("Top", oItem.Top);
            formItems.Add("Caption", "-");
            formItems.Add("UID", itemName);
            formItems.Add("FromPane", 0);
            formItems.Add("ToPane", 0);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            oItem = oForm.Items.Item("Hyphen");

            formItems = new Dictionary<string, object>();
            itemName = "MaxRate"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OOAT");
            formItems.Add("Alias", "U_BDOSMaxRat");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", oItem.Left + oItem.Width + 5);
            formItems.Add("Width", 50);
            formItems.Add("Top", oItem.Top);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);


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

            setVisibleFormItems(oForm, out errorText);
        }

        public static void setVisibleFormItems(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            try
            {
                oForm.Items.Item("BDOSPay").Enabled = (oForm.Mode != SAPbouiCOM.BoFormMode.fm_FIND_MODE && oForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE &&
                                                        oForm.DataSources.DBDataSources.Item("OOAT").GetValue("Status", 0).Trim() == "A");
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

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD & pVal.BeforeAction == true)
                {
                    createFormItems(oForm, out errorText);
                    Program.FORM_LOAD_FOR_VISIBLE = true;
                    Program.FORM_LOAD_FOR_ACTIVATE = true;

                    formDataLoad(oForm, out errorText);
                    setVisibleFormItems(oForm, out errorText);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE & pVal.BeforeAction == false)
                {
                    if (Program.FORM_LOAD_FOR_ACTIVATE == true)
                    {
                        setVisibleFormItems(oForm, out errorText);
                        Program.FORM_LOAD_FOR_ACTIVATE = false;
                    }
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_COMBO_SELECT)
                {
                    if ((pVal.ItemUID == "1250000036") && pVal.BeforeAction == false)
                    {
                        setVisibleFormItems(oForm, out errorText);
                    }
                }

                if (pVal.ItemUID == "1250000045" && (pVal.ColUID == "U_BDOSGrsPrc" || pVal.ColUID == "1250000009" || pVal.ColUID == "1250000007") & pVal.ItemChanged && pVal.EventType == SAPbouiCOM.BoEventTypes.et_VALIDATE & pVal.BeforeAction == false)
                {
                    oForm.Freeze(true);

                    SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("1250000045").Specific;
                    string ItemCode = oMatrix.Columns.Item("1250000001").Cells.Item(pVal.Row).Specific.Value;
                    decimal UnitPrice = (FormsB1.cleanStringOfNonDigits(oMatrix.Columns.Item("1250000009").Cells.Item(pVal.Row).Specific.Value));
                    decimal Qty = (FormsB1.cleanStringOfNonDigits(oMatrix.Columns.Item("1250000007").Cells.Item(pVal.Row).Specific.Value));
                    decimal GrossPrice = (FormsB1.cleanStringOfNonDigits(oMatrix.Columns.Item("U_BDOSGrsPrc").Cells.Item(pVal.Row).Specific.Value));
                    int row = pVal.Row;
                    decimal VatRate = CommonFunctions.GetVatGroupRate("", ItemCode);

                    if (pVal.ColUID == "1250000009")
                    {
                        GrossPrice = CommonFunctions.roundAmountByGeneralSettings(UnitPrice * (100 + VatRate) / 100, "Price");
                        oMatrix.Columns.Item("U_BDOSGrsPrc").Cells.Item(row).Specific.String = FormsB1.ConvertDecimalToStringForEditboxStrings(GrossPrice);
                        decimal amount = GrossPrice * Qty;
                        oMatrix.Columns.Item("U_BDOSGrsAmt").Cells.Item(row).Specific.String = FormsB1.ConvertDecimalToStringForEditboxStrings(amount);
                    }
                    else if (pVal.ColUID == "U_BDOSGrsPrc")
                    {
                        UnitPrice = CommonFunctions.roundAmountByGeneralSettings(GrossPrice * 100 / (100 + VatRate), "Price");
                        oMatrix.Columns.Item("1250000009").Cells.Item(row).Specific.String = FormsB1.ConvertDecimalToStringForEditboxStrings(UnitPrice);
                        decimal amount = GrossPrice * Qty;
                        oMatrix.Columns.Item("U_BDOSGrsAmt").Cells.Item(row).Specific.String = FormsB1.ConvertDecimalToStringForEditboxStrings(amount);
                    }
                    else
                    {
                        decimal amount = GrossPrice * Qty;
                        oMatrix.Columns.Item("U_BDOSGrsAmt").Cells.Item(row).Specific.String = FormsB1.ConvertDecimalToStringForEditboxStrings(amount);
                    }
                    oForm.Freeze(false);
                }

                if (pVal.ItemUID == "BDOSPay" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.BeforeAction == false)
                {
                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE &&
                                                        oForm.DataSources.DBDataSources.Item("OOAT").GetValue("Status", 0).Trim() == "A")
                    {
                        SAPbouiCOM.DBDataSource dBDataSource = oForm.DataSources.DBDataSources.Item("OOAT");

                        DateTime docDate = DateTime.Today;
                        string cardCode = dBDataSource.GetValue("BpCode", 0);
                        string cardName = dBDataSource.GetValue("BpName", 0);
                        string bpCurrency = dBDataSource.GetValue("BPCurr", 0);
                        string blnkAgr = dBDataSource.GetValue("AbsID", 0);
                        string locCurr = CommonFunctions.getLocalCurrency();

                        TableForPaymentDetail.Rows.Clear();
                        TableForPaymentDetail.Rows.Add();
                        TableForPaymentDetail.SetValue("LineNum", 0, 0);
                        TableForPaymentDetail.SetValue("CardCode", 0, cardCode);
                        TableForPaymentDetail.SetValue("CardName", 0, cardName);
                        TableForPaymentDetail.SetValue("BPCurrency", 0, bpCurrency);
                        TableForPaymentDetail.SetValue("PartnerCurrency", 0, locCurr);
                        TableForPaymentDetail.SetValue("currency", 0, locCurr);
                        TableForPaymentDetail.SetValue("BlnkAgr", 0, blnkAgr);
                        TableForPaymentDetail.SetValue("Project", 0, dBDataSource.GetValue("Project", 0));

                        SAPbouiCOM.Form oFormInternetBankingDocuments;

                        BDOSInternetBankingDocuments.automaticPaymentInternetBanking = true;
                        BDOSInternetBankingDocuments.openFromBlnkAgr = true;
                        BDOSInternetBankingDocuments.createForm(oForm, docDate, cardCode, cardName, bpCurrency, 0, locCurr, 0, 0, 0, 0, 0, out oFormInternetBankingDocuments, out errorText);
                        BDOSInternetBanking.TableExportMTRForDetail = BDOSInternetBanking.create_TableExportMTRForDetail();
                        BDOSInternetBankingDocuments.fillInvoicesMTR(oFormInternetBankingDocuments, blnkAgr, out errorText);
                    }
                }
            }
        }

        public static void uiApp_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(BusinessObjectInfo.FormTypeEx, Program.currentFormCount);

            if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD & BusinessObjectInfo.BeforeAction == false)
            {
                formDataLoad(oForm, out errorText);
            }
        }

        public static decimal GetBlAgremeentCurrencyRate(int docEntry, DateTime? docDate = null, decimal docRate = 0)
        {
            decimal minRate = 0;
            decimal maxRate = 0;
            string docCurr = null;
            decimal rateByBlnktAgr = 0;

            try
            {
                SAPbobsCOM.Recordset oRecordSetC = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string query = @"SELECT 
                    ""OOAT"".""AbsID"" AS ""docEntry"", 
                    ""OOAT"".""BPCurr"" AS ""DocCurr"",
                    ""OOAT"".""U_BDOSMinRat"" AS ""MinRate"",
                    ""OOAT"".""U_BDOSMaxRat"" AS ""MaxRate""
                    FROM ""OOAT"" AS ""OOAT"" 
                    WHERE ""OOAT"".""AbsID"" = '" + docEntry + @"'";

                oRecordSetC.DoQuery(query);

                while (!oRecordSetC.EoF)
                {
                    minRate = Convert.ToDecimal(oRecordSetC.Fields.Item("MinRate").Value, CultureInfo.InvariantCulture);
                    maxRate = Convert.ToDecimal(oRecordSetC.Fields.Item("MaxRate").Value, CultureInfo.InvariantCulture);
                    docCurr = oRecordSetC.Fields.Item("DocCurr").Value;
                    oRecordSetC.MoveNext();
                    break;
                }

                if (docDate.HasValue)
                {
                    SAPbobsCOM.SBObob oSBOBob = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);
                    SAPbobsCOM.Recordset RateRecordset = oSBOBob.GetCurrencyRate(docCurr, docDate.Value);

                    while (!RateRecordset.EoF)
                    {
                        decimal TempRate = Convert.ToDecimal(RateRecordset.Fields.Item("CurrencyRate").Value);

                        if (TempRate > maxRate)
                            rateByBlnktAgr = maxRate;
                        else if (TempRate < minRate)
                            rateByBlnktAgr = minRate;
                        else rateByBlnktAgr = TempRate;

                        RateRecordset.MoveNext();
                    }
                }
                else
                {
                    decimal TempRate = docRate;

                    if (TempRate > maxRate)
                        rateByBlnktAgr = maxRate;
                    else if (TempRate < minRate)
                        rateByBlnktAgr = minRate;
                    else rateByBlnktAgr = TempRate;
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            return rateByBlnktAgr;
        }

        public static bool UsesCurrencyExchangeRates(int absID)
        {
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                string query = @"SELECT ""OOAT"".""AbsID"" 
                FROM ""OOAT""
                WHERE (""OOAT"".""U_BDOSMinRat"" > 0 OR ""OOAT"".""U_BDOSMaxRat"" > 0) 
                     AND ""OOAT"".""AbsID"" = '" + absID + @"'";

                oRecordSet.DoQuery(query);

                if (!oRecordSet.EoF)
                {
                    return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                Marshal.ReleaseComObject(oRecordSet);
            }
        }
    }
}