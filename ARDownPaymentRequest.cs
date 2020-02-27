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
    static partial class ARDownPaymentRequest
    {
        public static void createUserFields(out string errorText)
        {
            errorText = null;

            Dictionary<string, object> fieldskeysMap;

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSDPMAmt");
            fieldskeysMap.Add("TableName", "DPI1");
            fieldskeysMap.Add("Description", "VAT Base Amount (Gross)");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSDPMVat");
            fieldskeysMap.Add("TableName", "DPI1");
            fieldskeysMap.Add("Description", "VAT Amount");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            GC.Collect();
        }

        public static void createFormItems(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            Dictionary<string, object> formItems = null;

            string itemName = "";

            // -------------------- Use blanket agreement rates-----------------
            int pane = 7;
            int left = oForm.Items.Item("1720002167").Left;
            int height = oForm.Items.Item("1720002167").Height;
            int top = oForm.Items.Item("1720002167").Top + height + 5;

            formItems = new Dictionary<string, object>();
            itemName = "UsBlaAgRtS"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "ODPI");
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
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);


            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            GC.Collect();
        }

        public static void getAmount(int docEntry, out double gTotal, out double lineVat)
        {
            gTotal = 0;
            lineVat = 0;

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string query = @"SELECT 
            ""DPI1"".""DocEntry"" AS ""docEntry"", 
            SUM(""DPI1"".""U_BDOSDPMAmt"") AS ""GTotal"", 
            SUM(""DPI1"".""U_BDOSDPMVat"") AS ""LineVat"" 
            FROM ""DPI1"" AS ""DPI1"" 
            WHERE ""DPI1"".""DocEntry"" = '" + docEntry + @"' 
            GROUP BY ""DPI1"".""DocEntry""";

            try
            {
                oRecordSet.DoQuery(query);

                while (!oRecordSet.EoF)
                {
                    gTotal = oRecordSet.Fields.Item("GTotal").Value;
                    lineVat = oRecordSet.Fields.Item("LineVat").Value;

                    oRecordSet.MoveNext();
                    break;
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
            }
        }

        public static void uiApp_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
        }

        public static void uiApp_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD && pVal.BeforeAction)
                {
                    ARDownPayment.createFormItems(oForm, out errorText);
                    createFormItems(oForm, out errorText);
                }

                else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
                {
                    if (pVal.ItemUID == "1" && pVal.BeforeAction)
                    {
                        CommonFunctions.fillDocRate(oForm, "ODPI");
                    }

                    else if(pVal.ItemUID == "UsBlaAgRtS" && !pVal.BeforeAction)
                    {
                        SAPbouiCOM.EditText oBlankAgr = (SAPbouiCOM.EditText)oForm.Items.Item("1980002192").Specific;

                        if (string.IsNullOrEmpty(oBlankAgr.Value))
                        {
                            Program.uiApp.SetStatusBarMessage(errorText = BDOSResources.getTranslate("EmptyBlaAgrError"), SAPbouiCOM.BoMessageTime.bmt_Short);
                            SAPbouiCOM.CheckBox oUsBlaAgRtCB = (SAPbouiCOM.CheckBox)oForm.Items.Item("UsBlaAgRtS").Specific;
                            oUsBlaAgRtCB.Checked = false;
                            oForm.Items.Item("1980002192").Click();
                        }
                    }
                }
            }
        }

        public static string createDocumentTransferFromBPType(SAPbouiCOM.DataTable oDataTable, SAPbouiCOM.Form oForm, int i, string cardCode, string bpCurrency, out int docEntry, out int docNum, out string errorText)
        {
            errorText = "";

            docNum = 0;
            docEntry = 0;

            SAPbobsCOM.Documents oDPM = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDownPayments);

            decimal addDPAmt = Convert.ToDecimal(oDataTable.GetValue("AddDownPaymentAmount", i), NumberFormatInfo.InvariantInfo);

            string currency = oDataTable.GetValue("Currency", i);
            string currencySapCode = CommonFunctions.getCurrencySapCode(currency);

            string GLAccountCode = CommonFunctions.getOADM("U_BDOSAtPayA").ToString();
            string linesText = CommonFunctions.getOADM("U_BDOSRcDPPr").ToString();

            if (string.IsNullOrEmpty(cardCode))
            {
                string partnerAccountNumber = oDataTable.GetValue("PartnerAccountNumber", i);
                string partnerCurrency = oDataTable.GetValue("PartnerCurrency", i);

                string localCurrency = CommonFunctions.getLocalCurrency();

                string partnerCurrencySapCode = CommonFunctions.getCurrencySapCode(partnerCurrency);
                if (string.IsNullOrEmpty(partnerCurrencySapCode))
                    partnerCurrencySapCode = localCurrency;

                string partnerTaxCode = oDataTable.GetValue("PartnerTaxCode", i);

                SAPbobsCOM.Recordset oRecordSet = CommonFunctions.getBPBankInfo(partnerAccountNumber + partnerCurrency, partnerTaxCode, "C");
                if (oRecordSet == null)
                {
                    errorText = BDOSResources.getTranslate("CouldNotFindBusinessPartner") + "! " + BDOSResources.getTranslate("Account") + " \"" + partnerAccountNumber + partnerCurrency + "\"";
                    if (string.IsNullOrEmpty(partnerTaxCode) == false)
                    {
                        errorText = errorText + ", " + BDOSResources.getTranslate("Tin") + " \"" + partnerTaxCode + "\"! ";
                    }
                    else errorText = errorText + "! ";
                }

                cardCode = oRecordSet.Fields.Item("CardCode").Value;
                bpCurrency = oRecordSet.Fields.Item("Currency").Value;
            }

            if (string.IsNullOrEmpty(errorText) == false)
            {
                errorText = BDOSResources.getTranslate("OperationCompletedUnSuccessfully") + "! " + errorText + BDOSResources.getTranslate("TableRow") + " : " + (i + 1);
                return null;
            }

            //string GLAccountCode = oDataTable.GetValue("GLAccountCode", i);
            oDPM.CardCode = cardCode;

            oDPM.DownPaymentType = SAPbobsCOM.DownPaymentTypeEnum.dptRequest;
            oDPM.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service;

            var DocDate = oDataTable.GetValue("DocumentDate", i);
            oDPM.DocDueDate = DocDate;
            oDPM.DocDate = DocDate;
            oDPM.TaxDate = DocDate;

            string BPCurrency = bpCurrency;

            oDPM.Lines.AccountCode = GLAccountCode;
            oDPM.Lines.ItemDescription = linesText;
            //oDPM.Lines.Quantity = 1;

            //if (partnerCurrencySapCode != currencySapCode)
            //{
            oDPM.Lines.Currency = currencySapCode;
            //}

            oDPM.Lines.PriceAfterVAT = Convert.ToDouble(addDPAmt);

            int returnCode = oDPM.Add();
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
                bool newDoc = oDPM.GetByKey(Convert.ToInt32(Program.oCompany.GetNewObjectKey()));
                if (newDoc == true)
                {
                    docEntry = oDPM.DocEntry;
                    docNum = oDPM.DocNum;
                }

                return BDOSResources.getTranslate("OperationCompletedSuccessfully") + "! " + BDOSResources.getTranslate("TableRow") + " : " + (i + 1);
            }
        }

        public static bool checkDocumentForTaxInvoice(int docEntry, DateTime docDate, DateTime docDateForMonth, out bool primary, out DataTable confirmedInvoices, out string errorText)
        {
            errorText = null;
            primary = false;
            confirmedInvoices = null;
            DataTable nonConfirmedInvoices = null;
            DateTime firstDay = new DateTime(docDateForMonth.Year, docDateForMonth.Month, 1);
            DateTime lastDay = firstDay.AddMonths(1).AddDays(-1);

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string query = @"SELECT
	             ""ODPI"".""DocDate"",
	             ""ODPI"".""DocEntry"",
	             ""ODPI"".""DocNum"",
	             ""ODPI"".""Posted"",
	             ""ODPI"".""CEECFlag"",
	             ""ODPI"".""DpmStatus"",
	             ""DPI1"".""BaseEntry"",
                 ""BDO_TAXS"".""DocEntry"" AS ""invDocEntry"",
	             ""BDO_TAXS"".""DocNum"" AS ""invDocNum"",
	             ""BDO_TAXS"".""U_status"",
	             ""BDO_TAXS"".""U_invID"",
	             ""BDO_TAXS"".""U_number"",
	             ""BDO_TAXS"".""U_series"" 
            FROM ""ODPI"" AS ""ODPI"" 
            INNER JOIN ""DPI1"" AS ""DPI1"" ON ""ODPI"".""DocEntry"" = ""DPI1"".""DocEntry"" 
            INNER JOIN (SELECT
            	 ""BDO_TXS1"".""U_baseDoc"" AS ""U_baseDoc"",
            	 ""BDO_TAXS"".""DocEntry"",
	             ""BDO_TAXS"".""DocNum"",            	 
            	 ""BDO_TAXS"".""U_status"" AS ""U_status"",
            	 ""BDO_TAXS"".""U_invID"" AS ""U_invID"",
            	 ""BDO_TAXS"".""U_number"" AS ""U_number"",
            	 ""BDO_TAXS"".""U_series"" AS ""U_series"" 
            	FROM ""@BDO_TXS1"" AS ""BDO_TXS1"" 
            	INNER JOIN ""@BDO_TAXS"" AS ""BDO_TAXS"" ON ""BDO_TXS1"".""DocEntry"" = ""BDO_TAXS"".""DocEntry"" 
            	WHERE ""BDO_TAXS"".""U_downPaymnt"" = 'Y' 
            	AND (""BDO_TAXS"".""Canceled"" = 'N' AND ""BDO_TAXS"".""U_status"" NOT IN ('removed',
            	 'canceled'))
            	AND ""BDO_TXS1"".""U_baseDocT"" = 'ARDownPaymentRequest' 
            	GROUP BY ""BDO_TXS1"".""U_baseDoc"",
          	     ""BDO_TAXS"".""DocEntry"",
	             ""BDO_TAXS"".""DocNum"",
            	 ""BDO_TAXS"".""U_status"",
            	 ""BDO_TAXS"".""U_invID"",
            	 ""BDO_TAXS"".""U_number"",
            	 ""BDO_TAXS"".""U_series"" ) AS ""BDO_TAXS"" ON ""ODPI"".""DocEntry"" = ""BDO_TAXS"".""U_baseDoc"" 
            WHERE 
                --""Posted"" = 'Y' AND 
                ""DPI1"".""BaseEntry"" IN (SELECT
            	 ""DPI1"".""BaseEntry"" AS ""BaseEntry"" 
            	FROM ""DPI1"" 
            	WHERE ""DPI1"".""DocEntry"" = " + docEntry + " " +
                @"AND ""DPI1"".""BaseType"" = 203 
            	GROUP BY ""DPI1"".""BaseEntry"")
            AND ""ODPI"".""DocDate"" <= '" + docDate.ToString("yyyyMMdd") + "' " +
            @"AND ""ODPI"".""DocDate"" >= '" + firstDay.ToString("yyyyMMdd") + @"' AND ""ODPI"".""DocDate"" <= '" + lastDay.ToString("yyyyMMdd") + "' " +
            @"AND ""ODPI"".""DocEntry"" < '" + docEntry + "' " +
            @"GROUP BY ""ODPI"".""DocDate"",
            	 ""ODPI"".""DocEntry"",
            	 ""ODPI"".""DocNum"",
            	 ""ODPI"".""Posted"",
            	 ""ODPI"".""CEECFlag"",
            	 ""ODPI"".""DpmStatus"",
            	 ""DPI1"".""BaseEntry"",
                 ""BDO_TAXS"".""DocEntry"",
	             ""BDO_TAXS"".""DocNum"",
            	 ""BDO_TAXS"".""U_status"",
            	 ""BDO_TAXS"".""U_invID"",
            	 ""BDO_TAXS"".""U_number"",
            	 ""BDO_TAXS"".""U_series""
            ORDER BY ""ODPI"".""DocDate"" DESC,
             ""ODPI"".""DocEntry"" DESC";

            try
            {
                oRecordSet.DoQuery(query);
                int recordCount = oRecordSet.RecordCount;

                if (recordCount == 0)
                {
                    primary = true;
                    return true;
                }
                else
                {
                    string invStatus;

                    confirmedInvoices = new DataTable();
                    confirmedInvoices.Columns.Add("DocEntry", typeof(int));
                    confirmedInvoices.Columns.Add("DocNum", typeof(int));
                    confirmedInvoices.Columns.Add("BaseEntry", typeof(int));
                    confirmedInvoices.Columns.Add("U_invID", typeof(string));
                    confirmedInvoices.Columns.Add("U_number", typeof(string));
                    confirmedInvoices.Columns.Add("U_series", typeof(string));
                    confirmedInvoices.Columns.Add("InvDocEntry", typeof(int));
                    confirmedInvoices.Columns.Add("InvDocNum", typeof(int));

                    nonConfirmedInvoices = new DataTable();
                    nonConfirmedInvoices.Columns.Add("DocEntry", typeof(int));
                    nonConfirmedInvoices.Columns.Add("DocNum", typeof(int));
                    nonConfirmedInvoices.Columns.Add("BaseEntry", typeof(int));
                    nonConfirmedInvoices.Columns.Add("U_invID", typeof(string));
                    nonConfirmedInvoices.Columns.Add("U_number", typeof(string));
                    nonConfirmedInvoices.Columns.Add("U_series", typeof(string));
                    nonConfirmedInvoices.Columns.Add("InvDocEntry", typeof(int));
                    nonConfirmedInvoices.Columns.Add("InvDocNum", typeof(int));

                    while (!oRecordSet.EoF)
                    {
                        invStatus = oRecordSet.Fields.Item("U_status").Value.ToString();
                        DataRow taxDataRow;
                        if (invStatus == "confirmed" || invStatus == "correctionConfirmed" || invStatus == "primary" || invStatus == "corrected")
                            taxDataRow = confirmedInvoices.Rows.Add();
                        else
                            taxDataRow = nonConfirmedInvoices.Rows.Add();

                        taxDataRow["DocEntry"] = Convert.ToInt32(oRecordSet.Fields.Item("DocEntry").Value);
                        taxDataRow["DocNum"] = Convert.ToInt32(oRecordSet.Fields.Item("DocNum").Value);
                        taxDataRow["BaseEntry"] = Convert.ToInt32(oRecordSet.Fields.Item("BaseEntry").Value);
                        taxDataRow["U_invID"] = oRecordSet.Fields.Item("U_invID").Value.ToString();
                        taxDataRow["U_number"] = oRecordSet.Fields.Item("U_number").Value.ToString();
                        taxDataRow["U_series"] = oRecordSet.Fields.Item("U_series").Value.ToString();
                        taxDataRow["InvDocEntry"] = Convert.ToInt32(oRecordSet.Fields.Item("InvDocEntry").Value);
                        taxDataRow["InvDocNum"] = Convert.ToInt32(oRecordSet.Fields.Item("InvDocNum").Value);

                        oRecordSet.MoveNext();
                    }

                    if (confirmedInvoices.Rows.Count > 0)
                    {
                        primary = false;
                    }
                    if (nonConfirmedInvoices.Rows.Count > 0)
                    {
                        List<int> oList = nonConfirmedInvoices.AsEnumerable().Select(r => r.Field<int>("InvDocNum")).ToList();
                        errorText = BDOSResources.getTranslate("OnARDownPaymentRequestThereIsAnotherARDownPaymentInvoiceWithTaxInvoiceSentTheStatusOfWhichShouldBeFromThisList") + " : " + "\"" + BDOSResources.getTranslate("deleted") + "\", \"" + BDOSResources.getTranslate("Canceled") + "\", \"" + BDOSResources.getTranslate("Denied") + "\", \"" + BDOSResources.getTranslate("Confirmed") + "\", \"" + BDOSResources.getTranslate("CorrectionConfirmed") + "\"! ";
                        if (oList.Count > 1)
                            errorText = errorText + '\n' + "\"" + BDOSResources.getTranslate("TaxInvoiceSent") + "\" " + BDOSResources.getTranslate("DocumentsSNumbersAre") + " : " + string.Join(",", oList);
                        else
                            errorText = errorText + '\n' + "\"" + BDOSResources.getTranslate("TaxInvoiceSent") + "\" " + BDOSResources.getTranslate("DocumentSNumberIs") + " : " + string.Join(",", oList);

                        return false;
                    }
                    return true;
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
                return false;
            }
            finally
            {
                Marshal.ReleaseComObject(oRecordSet);
                oRecordSet = null;
            }
        }

        public static List<int> getAllConnectedDoc(List<int> docEntry)
        {
            List<int> connectedDocList = new List<int>();

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string query = @"SELECT
            	 ""DPI1"".""DocEntry"" 
            FROM ""DPI1"" 
            WHERE ""DPI1"".""BaseEntry"" IN (SELECT
            	 ""DPI1"".""BaseEntry"" 
            	FROM ""DPI1"" 
            	WHERE ""DPI1"".""DocEntry"" IN (" + string.Join(",", docEntry) + @") 
            	AND ""DPI1"".""BaseType"" = '203') 
            AND ""DPI1"".""BaseType"" = '203' 
            AND ""DPI1"".""DocEntry"" NOT IN (" + string.Join(",", docEntry) + @")
            GROUP BY ""DPI1"".""DocEntry""";

            try
            {
                oRecordSet.DoQuery(query);
                if (oRecordSet.RecordCount > 0)
                {
                    while (!oRecordSet.EoF)
                    {
                        connectedDocList.Add(Convert.ToInt32(oRecordSet.Fields.Item("DocEntry").Value));
                        oRecordSet.MoveNext();
                    }
                }
                return connectedDocList;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                oRecordSet = null;
                GC.Collect();
            }
        }
    }
}
