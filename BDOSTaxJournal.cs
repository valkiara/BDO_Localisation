using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Data;
using System.Runtime.InteropServices;

namespace BDO_Localisation_AddOn
{
    static partial class BDOSTaxJournal
    {
        //----------------------------->Tax Invoice Sent<-----------------------------       
        public static void rsOperationTaxInvoiceSent( SAPbouiCOM.Form oForm,  int oOperation, out string errorText)
        {
            errorText = null;
            bool opSuccess = true;
            string errorTextGoods = null;
            string errorTextWb = null;

            Dictionary<string, string> rsSettings = CompanyDetails.getRSSettings( out errorText);
            if (errorText != null)
            {
                Program.uiApp.StatusBar.SetSystemMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return;
            }

            string su = rsSettings["SU"];
            string sp = rsSettings["SP"];

            TaxInvoice oTaxInvoice = new TaxInvoice(rsSettings["ProtocolType"]);

            bool chek_service_user = oTaxInvoice.check_usr( su, sp, out errorText);
            if (chek_service_user == false)
            {
                errorText = BDOSResources.getTranslate("ServiceUserPasswordNotCorrect");
                Program.uiApp.StatusBar.SetSystemMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return;
            }

            try
            {
                oForm.Freeze(true);
                int answer = 0;

                ///////////////********************************************//////////////////
                SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("TxTable2").Specific));
                int rowCount = oMatrix.RowCount;

                if (oOperation == 5)
                {
                    DataTable oDataTable = new DataTable();
                    for (int oColumns = 0; oColumns <= oMatrix.Columns.Count - 1; oColumns++)
                    {
                        if (oMatrix.Columns.Item(oColumns).UniqueID != "TxChkBx")
                        {
                            oDataTable.Columns.Add(oMatrix.Columns.Item(oColumns).UniqueID);
                        }
                    }

                    List<int> docEntryARCreditNoteList = new List<int>();

                    for (int row = 1; row <= rowCount; row++)
                    {
                        SAPbouiCOM.CheckBox Edtfield = oMatrix.Columns.Item("TxChkBx").Cells.Item(row).Specific;
                        bool checkedLine = (Edtfield.Checked);

                        string docEntry = oMatrix.Columns.Item("Document").Cells.Item(row).Specific.Value;
                        string docType = oMatrix.Columns.Item("DocType").Cells.Item(row).Specific.Value;

                        if (checkedLine && docEntry == "" && (docType == "ARInvoice" || docType == "ARCreditNote"))
                        {
                            DataRow newRow = oDataTable.NewRow();
                            for (int col = 0; col <= oDataTable.Columns.Count - 1; col++)
                            {
                                string UniqueID = oDataTable.Columns[col].Caption;
                                newRow[UniqueID] = oMatrix.Columns.Item(UniqueID).Cells.Item(row).Specific.Value;
                            }
                            oDataTable.Rows.Add(newRow);

                            if (docType == "ARCreditNote")
                                docEntryARCreditNoteList.Add(Convert.ToInt32(newRow["InvEntry"]));
                        }
                        else if (checkedLine && docType == "ARDownPaymentVAT")
                        {
                            errorText = BDOSResources.getTranslate("UnitedTaxInvoiceCanNotContainDocumentWithType") + " " + BDOSResources.getTranslate("ARDownPaymentVAT") + "! " + BDOSResources.getTranslate("TableRow") + " : " + row;
                            Program.uiApp.StatusBar.SetSystemMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            errorText = null;
                            continue;
                        }
                        else if(checkedLine)
                        {
                            errorText = BDOSResources.getTranslate("TaxInvoiceAlreadyCreated") + "! " + BDOSResources.getTranslate("TableRow") + " : " + row;
                            Program.uiApp.StatusBar.SetSystemMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            errorText = null;
                            continue;
                        }
                    }

                    rowCount = oDataTable.Rows.Count;
                    if (rowCount > 0)
                    {
                        oDataTable.DefaultView.Sort = "CardCode ASC,OpDate ASC,DocType DESC";
                        oDataTable = oDataTable.DefaultView.ToTable();

                        SAPbobsCOM.CompanyService oCompanyService = null;
                        SAPbobsCOM.GeneralService oGeneralService = null;
                        SAPbobsCOM.GeneralData oGeneralData = null;

                        oCompanyService = Program.oCompany.GetCompanyService();
                        oGeneralService = oCompanyService.GetGeneralService("UDO_F_BDO_TAXS_D");

                        string docEntry = null;
                        string docType = null;
                        string invEntry = null;
                        string opDateStr = oDataTable.Rows[0]["OpDate"].ToString();
                        string cardCode = oDataTable.Rows[0]["CardCode"].ToString();
                        DateTime docDate = DateTime.ParseExact(oDataTable.Rows[0]["OpDate"].ToString(), "yyyyMMdd", null);
                        docDate = new DateTime(docDate.Year, docDate.Month, 1);
                        int newDocEntry = 0;
                        List<int> docEntryList = new List<int>();

                        oGeneralData = BDO_TaxInvoiceSent.createDocumentForUnion( cardCode, docDate, out errorText);
                        //int localRow = 0;

                        for (int row = 0; row < rowCount; row++)
                        {
                            newDocEntry = 0;
                            docEntry = oDataTable.Rows[row]["Document"].ToString();
                            docType = oDataTable.Rows[row]["DocType"].ToString();
                            invEntry = oDataTable.Rows[row]["InvEntry"].ToString();

                            if (cardCode != oDataTable.Rows[row]["CardCode"].ToString() || opDateStr != oDataTable.Rows[row]["OpDate"].ToString())
                            {
                                try
                                {
                                    if (string.IsNullOrEmpty(errorText) && oGeneralData != null)
                                    {
                                        decimal amount = 0;
                                        decimal amountTX = 0;
                                        decimal amtOutTX = amount - amountTX;
                                        SAPbobsCOM.GeneralDataCollection oChildren = oGeneralData.Child("BDO_TXS1");

                                        if (oChildren.Count > 0)
                                        {
                                            foreach (SAPbobsCOM.GeneralData oChild in oChildren)
                                            {
                                                if (oChild.GetProperty("U_baseDocT") == "ARInvoice")
                                                {
                                                    amount = amount + Convert.ToDecimal(oChild.GetProperty("U_amtBsDc"));
                                                    amountTX = amountTX + Convert.ToDecimal(oChild.GetProperty("U_tAmtBsDc"));
                                                }
                                                else if (oChild.GetProperty("U_baseDocT") == "ARCreditNote")
                                                {
                                                    amount = amount - Convert.ToDecimal(oChild.GetProperty("U_amtBsDc"));
                                                    amountTX = amountTX - Convert.ToDecimal(oChild.GetProperty("U_tAmtBsDc"));
                                                }
                                            }

                                            oGeneralData.SetProperty("U_amount", Convert.ToDouble(amount)); //თანხა დღგ-ის ჩათვლით
                                            oGeneralData.SetProperty("U_amountTX", Convert.ToDouble(amountTX)); //დღგ-ის თანხა
                                            oGeneralData.SetProperty("U_amtOutTX", Convert.ToDouble(amount - amountTX)); //თანხა დღგ-ის გარეშე

                                            var response = oGeneralService.Add(oGeneralData);
                                            newDocEntry = Convert.ToInt32(response.GetProperty("DocEntry"));
                                            docEntryList.Add(newDocEntry);

                                            string text = BDOSResources.getTranslate("UnitedTaxInvoiceCreatedSuccesfully") + " : " + newDocEntry;
                                            Program.uiApp.StatusBar.SetSystemMessage(text, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                                        }
                                        oGeneralData = null;
                                        //for (int i = localRow; i < row; i++)
                                        //{
                                        //    oDataTable.Rows[i]["Document"] = newDocEntry.ToString();
                                        //}
                                        //localRow = row;
                                    }
                                    else
                                    {
                                        errorText = BDOSResources.getTranslate("ErrorDocumentAdd") + " : " + errorText + " " + BDOSResources.getTranslate("BPCardCode") + " : " + cardCode + ", " + opDateStr;
                                        Program.uiApp.StatusBar.SetSystemMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    errorText = BDOSResources.getTranslate("ErrorDocumentAdd") + " : " + ex.Message + "! " + BDOSResources.getTranslate("BPCardCode") + " : " + cardCode + ", " + opDateStr;
                                    Program.uiApp.StatusBar.SetSystemMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                }

                                opDateStr = oDataTable.Rows[row]["OpDate"].ToString();
                                cardCode = oDataTable.Rows[row]["CardCode"].ToString();
                                docDate = DateTime.ParseExact(oDataTable.Rows[row]["OpDate"].ToString(), "yyyyMMdd", null);
                                docDate = new DateTime(docDate.Year, docDate.Month, 1);

                                oGeneralData = BDO_TaxInvoiceSent.createDocumentForUnion( cardCode, docDate, out errorText);
                            }

                            if (string.IsNullOrEmpty(errorText) && oGeneralData != null)
                            {
                                errorText = null;
                                if (docType == "ARInvoice")
                                {
                                    BDO_TaxInvoiceSent.createDocument( "13", Convert.ToInt32(invEntry), null, false, answer, oGeneralData, true, docEntryARCreditNoteList, out newDocEntry, out errorText);
                                }
                                else if (docType == "ARCreditNote")
                                {
                                    BDO_TaxInvoiceSent.createDocument( "14", Convert.ToInt32(invEntry), null, false, 1, oGeneralData, true, null, out newDocEntry, out errorText);
                                }
                                if (string.IsNullOrEmpty(errorText) == false)
                                {
                                    Program.uiApp.StatusBar.SetSystemMessage(errorText + " " + BDOSResources.getTranslate("TableRow") + " : " + oDataTable.Rows[row]["LineNum"].ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                    errorText = null;
                                    continue;
                                }
                            }
                        }

                        try
                        {
                            if (string.IsNullOrEmpty(errorText) && oGeneralData != null)
                            {
                                decimal amount = 0;
                                decimal amountTX = 0;
                                decimal amtOutTX = amount - amountTX;
                                SAPbobsCOM.GeneralDataCollection oChildren = oGeneralData.Child("BDO_TXS1");

                                if (oChildren.Count > 0)
                                {
                                    foreach (SAPbobsCOM.GeneralData oChild in oChildren)
                                    {
                                        if (oChild.GetProperty("U_baseDocT") == "ARInvoice")
                                        {
                                            amount = amount + Convert.ToDecimal(oChild.GetProperty("U_amtBsDc"));
                                            amountTX = amountTX + Convert.ToDecimal(oChild.GetProperty("U_tAmtBsDc"));
                                        }
                                        else if (oChild.GetProperty("U_baseDocT") == "ARCreditNote")
                                        {
                                            amount = amount - Convert.ToDecimal(oChild.GetProperty("U_amtBsDc"));
                                            amountTX = amountTX - Convert.ToDecimal(oChild.GetProperty("U_tAmtBsDc"));
                                        }
                                    }

                                    oGeneralData.SetProperty("U_amount", Convert.ToDouble(amount)); //თანხა დღგ-ის ჩათვლით
                                    oGeneralData.SetProperty("U_amountTX", Convert.ToDouble(amountTX)); //დღგ-ის თანხა
                                    oGeneralData.SetProperty("U_amtOutTX", Convert.ToDouble(amount - amountTX)); //თანხა დღგ-ის გარეშე

                                    var response = oGeneralService.Add(oGeneralData);
                                    newDocEntry = Convert.ToInt32(response.GetProperty("DocEntry"));
                                    docEntryList.Add(newDocEntry);

                                    string text = BDOSResources.getTranslate("UnitedTaxInvoiceCreatedSuccesfully") + " : " + newDocEntry;
                                    Program.uiApp.StatusBar.SetSystemMessage(text, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                                }
                                oGeneralData = null;
                            }
                            else
                            {
                                errorText = BDOSResources.getTranslate("ErrorDocumentAdd") + " : " + errorText + " " + BDOSResources.getTranslate("BPCardCode") + " : " + cardCode + ", " + opDateStr;
                                Program.uiApp.StatusBar.SetSystemMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            }
                        }
                        catch (Exception ex)
                        {
                            errorText = BDOSResources.getTranslate("ErrorDocumentAdd") + " : " + ex.Message + "! " + BDOSResources.getTranslate("BPCardCode") + " : " + cardCode + ", " + opDateStr;
                            Program.uiApp.StatusBar.SetSystemMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        }

                        //rs.ge - ზე შექმნა --->
                        for (int i = 0; i < docEntryList.Count; i++)
                        {
                            BDO_TaxInvoiceSent.operationRS( oTaxInvoice, "save", docEntryList[i], -1, new DateTime(), out errorText, out errorTextWb, out errorTextGoods);

                            if (errorText != null || errorTextWb != null || errorTextGoods != null)
                            {
                                string errorTextFull = "";

                                if (errorText != null)
                                {
                                    errorTextFull = errorTextFull + errorText;
                                }

                                if (errorTextWb != null)
                                {
                                    errorTextFull = errorTextFull + errorTextWb;
                                }

                                if (errorTextGoods != null)
                                {
                                    errorTextFull = errorTextFull + errorTextGoods;
                                }

                                Program.uiApp.StatusBar.SetSystemMessage(errorTextFull + " " + BDOSResources.getTranslate("TaxInvoice") + " : " + docEntryList[i], SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            }
                            else
                            {
                                Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("Operation") + " " + BDOSResources.getTranslate("RSSave") + " " + BDOSResources.getTranslate("DoneSuccessfully") + " " + BDOSResources.getTranslate("TaxInvoice") + " : " + docEntryList[i], SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                            }
                        }
                        //rs.ge - ზე შექმნა <---

                        //სტრიქონის განახლება --->
                        for (int i = 0; i < docEntryList.Count; i++)
                        {
                            updateMatrixRowTaxInvoiceSent( oForm, oMatrix, docEntryList[i]);
                        }
                        //სტრიქონის განახლება <---
                    }
                    return;
                }

                answer = 0;

                for (int row = 1; row <= rowCount; row++)
                {
                    SAPbouiCOM.CheckBox Edtfield = oMatrix.Columns.Item("TxChkBx").Cells.Item(row).Specific;
                    bool checkedLine = (Edtfield.Checked);

                    if (checkedLine)
                    {
                        string docEntry = oMatrix.Columns.Item("Document").Cells.Item(row).Specific.Value;
                        string invEntry = oMatrix.Columns.Item("InvEntry").Cells.Item(row).Specific.Value;
                        string docType = oMatrix.Columns.Item("DocType").Cells.Item(row).Specific.Value;
                        string corrType = oMatrix.Columns.Item("CorrType").Cells.Item(row).Specific.Value;
                        string corrDoc = oMatrix.Columns.Item("CorrDoc").Cells.Item(row).Specific.Value;
                        string statusDoc = oMatrix.Columns.Item("StatusDoc").Cells.Item(row).Specific.Value;

                        if ((oOperation == 0 || oOperation == 1) && (docType == "ARInvoice" || docType == "ARCreditNote") && answer == 0)
                        {
                            answer = Program.uiApp.MessageBox(BDOSResources.getTranslate("ARCDocumentsOnARIDoYouWantToCreateUnitedTaxInvoiceIncludingTheseDocuments") + "?", 1, BDOSResources.getTranslate("Yes"), BDOSResources.getTranslate("No"), ""); //არსებობს რეალიზაციის დოკუმენტზე რეალიზაციის კორექტირების დოკუმენტები, გსურთ შეიქმნას ერთიანი ფაქტურა ამ დოკუმენტების გათვალისწინებით                   
                        }

                        int newDocEntry = 0;
                        string opText = "";

                        if (oOperation == 0) //შენახვა
                        {
                            opText = BDOSResources.getTranslate("RSSave");

                            if (docEntry == "")
                            {
                                if (string.IsNullOrEmpty(corrDoc) == false && string.IsNullOrEmpty(corrType))
                                {
                                    errorText = BDOSResources.getTranslate("CorrectionReasonNotIndicated") + " " + BDOSResources.getTranslate(docType) + " : " + invEntry;
                                }
                                else if (docType == "ARInvoice")
                                {
                                    BDO_TaxInvoiceSent.createDocument( "13", Convert.ToInt32(invEntry), corrType, false, answer, null, false, null, out newDocEntry, out errorText);
                                }
                                else if (docType == "ARCreditNote")
                                {
                                    BDO_TaxInvoiceSent.createDocument( "14", Convert.ToInt32(invEntry), corrType, false, answer, null, false, null, out newDocEntry, out errorText);
                                }
                                else if (docType == "ARDownPaymentVAT")
                                {
                                    BDO_TaxInvoiceSent.createDocument( "UDO_F_BDO_ARDPV_D", Convert.ToInt32(invEntry), corrType, false, 0, null, false, null, out newDocEntry, out errorText);
                                }
                                docEntry = newDocEntry.ToString();
                            }
                            if (errorText == null)
                            {
                                BDO_TaxInvoiceSent.operationRS( oTaxInvoice, "save", Convert.ToInt32(docEntry), -1, new DateTime(), out errorText, out errorTextWb, out errorTextGoods);
                            }
                        }
                        else if (oOperation == 1) //გადაგზავნა
                        {
                            opText = BDOSResources.getTranslate("RSSend");

                            if (docEntry == "")
                            {
                                if (string.IsNullOrEmpty(corrDoc) == false && string.IsNullOrEmpty(corrType))
                                {
                                    errorText = BDOSResources.getTranslate("CorrectionReasonNotIndicated") + " " + BDOSResources.getTranslate(docType) + " : " + invEntry;
                                }
                                else if (docType == "ARInvoice")
                                {
                                    BDO_TaxInvoiceSent.createDocument( "13", Convert.ToInt32(invEntry), corrType, false, answer, null, false, null, out newDocEntry, out errorText);
                                }
                                else if (docType == "ARCreditNote")
                                {
                                    BDO_TaxInvoiceSent.createDocument( "14", Convert.ToInt32(invEntry), corrType, false, answer, null, false, null, out newDocEntry, out errorText);
                                }
                                else if (docType == "ARDownPaymentVAT")
                                {
                                    BDO_TaxInvoiceSent.createDocument( "UDO_F_BDO_ARDPV_D", Convert.ToInt32(invEntry), corrType, false, 0, null, false, null, out newDocEntry, out errorText);
                                }
                                docEntry = newDocEntry.ToString();
                            }

                            if (errorText == null)
                            {
                                BDO_TaxInvoiceSent.operationRS( oTaxInvoice, "send", Convert.ToInt32(docEntry), -1, new DateTime(), out errorText, out errorTextWb, out errorTextGoods);
                            }
                        }
                        else if (oOperation == 2) //წაშლა
                        {
                            opText = BDOSResources.getTranslate("RSDelete");
                            BDO_TaxInvoiceSent.operationRS( oTaxInvoice, "remove", Convert.ToInt32(docEntry), -1, new DateTime(), out errorText, out errorTextWb, out errorTextGoods);
                        }
                        else if (oOperation == 3) //გაუქმება
                        {
                            opText = BDOSResources.getTranslate("RSCancel");
                            BDO_TaxInvoiceSent.operationRS( oTaxInvoice, "cancel", Convert.ToInt32(docEntry), -1, new DateTime(), out errorText, out errorTextWb, out errorTextGoods);
                        }
                        else if (oOperation == 4) //სტატუსების განხლება
                        {
                            opText = BDOSResources.getTranslate("RSUpdateStatus");
                            if (docEntry != "")
                            {
                                BDO_TaxInvoiceSent.operationRS( oTaxInvoice, "updateStatus", Convert.ToInt32(docEntry), -1, new DateTime(), out errorText, out errorTextWb, out errorTextGoods);
                            }
                        }

                        if (errorText != null || errorTextWb != null || errorTextGoods != null)
                        {
                            string errorTextFull = "";

                            if (errorText != null)
                            {
                                errorTextFull = errorTextFull + errorText;
                            }

                            if (errorTextWb != null)
                            {
                                errorTextFull = errorTextFull + errorTextWb;
                            }

                            if (errorTextGoods != null)
                            {
                                errorTextFull = errorTextFull + errorTextGoods;
                            }

                            Program.uiApp.StatusBar.SetSystemMessage(errorTextFull + " " + BDOSResources.getTranslate("TableRow") + " : " + row.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            opSuccess = false;
                        }
                        else
                        {
                            Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("Operation") + " " + opText + " " + BDOSResources.getTranslate("DoneSuccessfully") + " " + BDOSResources.getTranslate("TableRow") + " : " + row.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                        }
                        if (docEntry != "" & docEntry != "0")
                        {
                            updateMatrixRowTaxInvoiceSent( oForm, oMatrix, Convert.ToInt32(docEntry));
                            //fillFromBaseTaxInvoiceSent(  oForm, false, out errorText);
                        }
                    }
                }
                if (opSuccess == false)
                {
                    Program.uiApp.MessageBox(BDOSResources.getTranslate("OperationForSomeDocumentFinishedWithErrorSeeLog"));
                }
                //if (answer == 2)
                //{
                //    updateMatrixRowTaxInvoiceSent( oForm, oMatrix, Convert.ToInt32(docEntry));
                //    //fillFromBaseTaxInvoiceSent(  oForm, false, out errorText);
                //}
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
            finally
            {
                oForm.Update();
                oForm.Freeze(false);
            }
        }

        private static void updateMatrixRowTaxInvoiceSent( SAPbouiCOM.Form oForm, SAPbouiCOM.Matrix oMatrix, int docEntry)
        {
            string d;

            try
            {
                NumberFormatInfo Nfi = new NumberFormatInfo() { NumberDecimalSeparator = "." };

                SAPbobsCOM.CompanyService oCompanyService = null;
                SAPbobsCOM.GeneralService oGeneralService = null;
                SAPbobsCOM.GeneralData oGeneralData = null;
                SAPbobsCOM.GeneralDataParams oGeneralParams = null;
                oCompanyService = Program.oCompany.GetCompanyService();
                oGeneralService = oCompanyService.GetGeneralService("UDO_F_BDO_TAXS_D");
                oGeneralData = ((SAPbobsCOM.GeneralData)(oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)));

                oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                oGeneralParams.SetProperty("DocEntry", docEntry);
                oGeneralData = oGeneralService.GetByParams(oGeneralParams);

                SAPbobsCOM.GeneralDataCollection oChildren = null;
                oChildren = oGeneralData.Child("BDO_TXS1");

                string declNumber = oGeneralData.GetProperty("U_declNumber");
                string corrInv = oGeneralData.GetProperty("U_corrInv");
                double amount = oGeneralData.GetProperty("U_amount"); //თანხა დღგ-ის ჩათვლით
                double amountTX = oGeneralData.GetProperty("U_amountTX"); //დღგ-ის თანხა
                string corrDocEntry = oGeneralData.GetProperty("U_corrDTxt");

                //if (corrInv == "Y")
                //{
                //    amount = oGeneralData.GetProperty("U_amtACor"); //თანხა დღგ-ის ჩათვლით
                //    amountTX = oGeneralData.GetProperty("U_amtTXACr"); //დღგ-ის თანხა               
                //}

                string opDate = oGeneralData.GetProperty("U_opDate").ToString("yyyyMMdd") == "18991230" ? "" : oGeneralData.GetProperty("U_opDate").ToString("yyyyMMdd");
                string sentDate = oGeneralData.GetProperty("U_sentDate").ToString("yyyyMMdd") == "18991230" ? "" : oGeneralData.GetProperty("U_sentDate").ToString("yyyyMMdd");
                string confDate = oGeneralData.GetProperty("U_confDate").ToString("yyyyMMdd") == "18991230" ? "" : oGeneralData.GetProperty("U_confDate").ToString("yyyyMMdd");

                SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("TxTable2");

                int count = oChildren.Count;
                string baseDocType;
                int baseDoc;
                int row;               

                if (count > 0)
                {
                    foreach (SAPbobsCOM.GeneralData oChild in oChildren)
                    {
                        baseDocType = oChild.GetProperty("U_baseDocT");
                        baseDoc = oChild.GetProperty("U_baseDoc");

                        for (int i = 0; i < oDataTable.Rows.Count; i++)
                        {
                            if (baseDocType == oDataTable.GetValue("DocType", i) && baseDoc == Convert.ToInt32(oDataTable.GetValue("InvoiceEntry", i)))
                            {
                                row = i + 1;
                                oMatrix.Columns.Item("LineNum").Cells.Item(row).Specific.Value = row;
                                oMatrix.Columns.Item("StatusDoc").Cells.Item(row).Specific.Select(oGeneralData.GetProperty("U_status"), SAPbouiCOM.BoSearchKey.psk_ByValue);
                                oMatrix.Columns.Item("TxID").Cells.Item(row).Specific.Value = oGeneralData.GetProperty("U_invID");
                                oMatrix.Columns.Item("TxSerie").Cells.Item(row).Specific.Value = oGeneralData.GetProperty("U_series");
                                oMatrix.Columns.Item("TxNum").Cells.Item(row).Specific.Value = oGeneralData.GetProperty("U_number");
                                oMatrix.Columns.Item("Document").Cells.Item(row).Specific.Value = docEntry;
                                oMatrix.Columns.Item("CorrDoc").Cells.Item(row).Specific.Value = corrDocEntry;
                                oMatrix.Columns.Item("DeclNum").Cells.Item(row).Specific.Value = declNumber;
                                oMatrix.Columns.Item("DeclStatus").Cells.Item(row).Specific.Value = (String.IsNullOrEmpty(declNumber) == true) ? BDOSResources.getTranslate("WithoutDeclaration") : BDOSResources.getTranslate("WithDeclaration");
                                oMatrix.Columns.Item("CardCode").Cells.Item(row).Specific.Value = oGeneralData.GetProperty("U_cardCode");
                                oMatrix.Columns.Item("RegDate").Cells.Item(row).Specific.Value = sentDate;
                                oMatrix.Columns.Item("ConfDate").Cells.Item(row).Specific.Value = confDate;
                                oMatrix.Columns.Item("Comment").Cells.Item(row).Specific.Value = oGeneralData.GetProperty("U_comment");
                                oMatrix.CommonSetting.SetCellBackColor(row, 2, FormsB1.getLongIntRGB(231, 231, 231));
                                oMatrix.CommonSetting.SetCellBackColor(row, 3, FormsB1.getLongIntRGB(231, 231, 231));
                                break;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                d = ex.Message;
            }
        }

        public static void addDeclTaxInvoiceSent(  SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            string errorTextWb = null;
            string errorTextGoods = null;
            oForm.Freeze(true);
            bool opSuccess = true;

            oForm.Items.Item("DeclNum2").Specific.Value = "";
            int seqNum = 0;

            Dictionary<string, string> rsSettings = CompanyDetails.getRSSettings( out errorText);
            if (errorText != null)
            {
                oForm.Update();
                oForm.Freeze(false);
                Program.uiApp.StatusBar.SetSystemMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return;
            }

            string su = rsSettings["SU"];
            string sp = rsSettings["SP"];

            TaxInvoice oTaxInvoice = new TaxInvoice(rsSettings["ProtocolType"]);

            bool chek_service_user = oTaxInvoice.check_usr( su, sp, out errorText);
            if (chek_service_user == false)
            {
                oForm.Update();
                oForm.Freeze(false);
                errorText = BDOSResources.getTranslate("ServiceUserPasswordNotCorrect");
                Program.uiApp.StatusBar.SetSystemMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return;
            }

            int DeclYear = Convert.ToInt32(oForm.Items.Item("DeclYear2").Specific.Value) * -1 + DateTime.Now.Year;
            int DeclMonth = Convert.ToInt32(oForm.Items.Item("DeclMonth2").Specific.Value) + 1;
            string DeclDateString = new DateTime(DeclYear, DeclMonth, 1).ToString("yyyyMM");
            DateTime DeclDate = new DateTime(DeclYear, DeclMonth, 1);

            //დეკლარაციის ნომრების მიღება
            DataTable TaxDeclTable = oTaxInvoice.get_seq_nums(DeclDateString, out errorText);

            for (int i = 0; i < TaxDeclTable.Rows.Count; i++)
            {
                DataRow TaxDeclRow = TaxDeclTable.Rows[i];
                seqNum = Convert.ToInt32(TaxDeclRow.ItemArray[0]);
                oForm.Items.Item("DeclNum2").Specific.Value = seqNum.ToString();
            }

            if (seqNum == 0)
            {
                oForm.Update();
                oForm.Freeze(false);
                errorText = BDOSResources.getTranslate("CantReceiveDeclarationDate") + " " + DeclDate;
                Program.uiApp.StatusBar.SetSystemMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return;
            }
            //oForm.Freeze(true);
            //დეკლარაციაში დამატება
            //string statusRS = null;
            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("TxTable2").Specific));
            int rowCount = oMatrix.RowCount;
            for (int row = 1; row <= rowCount; row++)
            {
                SAPbouiCOM.CheckBox Edtfield = oMatrix.Columns.Item("TxChkBx").Cells.Item(row).Specific;
                bool checkedLine = (Edtfield.Checked);

                if (checkedLine)
                {
                    int docEntry = Convert.ToInt32(oMatrix.Columns.Item("Document").Cells.Item(row).Specific.Value);
                    string statusDoc = oMatrix.Columns.Item("StatusDoc").Cells.Item(row).Specific.Value;
                    string declNum = oMatrix.Columns.Item("DeclNum").Cells.Item(row).Specific.Value;
                    //oMatrix.Columns.Item("TxChkBx").Cells.Item(row).Specific.Checked = false;

                    BDO_TaxInvoiceSent.operationRS( oTaxInvoice, "addToTheDeclaration", Convert.ToInt32(docEntry), seqNum, DeclDate, out errorText, out errorTextWb, out errorTextGoods);

                    if (errorText != null || errorTextWb != null || errorTextGoods != null)
                    {
                        string errorTextFull = "";
                        if (errorText != null)
                        {
                            errorTextFull = errorTextFull + errorText;
                        }

                        if (errorTextWb != null)
                        {
                            errorTextFull = errorTextFull + errorTextWb;
                        }

                        if (errorTextGoods != null)
                        {
                            errorTextFull = errorTextFull + errorTextGoods;
                        }

                        Program.uiApp.StatusBar.SetSystemMessage(errorTextFull + " " + BDOSResources.getTranslate("TableRow") + " : " + row.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        opSuccess = false;
                        continue;
                    }
                    else
                    {
                        Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("Operation") + " " + BDOSResources.getTranslate("RSAddDeclaration") + " " + BDOSResources.getTranslate("DoneSuccessfully") + " " + BDOSResources.getTranslate("TableRow") + " : " + row.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                        if (docEntry != 0)
                        {
                            updateMatrixRowTaxInvoiceSent( oForm, oMatrix, docEntry);
                            //fillFromBaseTaxInvoiceSent(  oForm, false, out errorText);
                        }
                    }
                }
            }
            if (opSuccess == false)
            {
                Program.uiApp.MessageBox(BDOSResources.getTranslate("OperationForSomeDocumentFinishedWithErrorSeeLog"));
            }
            oForm.Update();
            oForm.Freeze(false);
        }

        public static void fillFromBaseTaxInvoiceSent(  SAPbouiCOM.Form oForm, bool download, out string errorText)
        {
            errorText = null;

            SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("TxTable2");
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            DateTime OperationPeriodStart = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            DateTime OperationPeriodEnd = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month));
            OperationPeriodEnd = OperationPeriodEnd.AddSeconds(24 * 3600 - 1);

            string startDateStr = oForm.DataSources.UserDataSources.Item("StartDate2").ValueEx;
            string endDateStr = oForm.DataSources.UserDataSources.Item("EndDate2").ValueEx;

            string startDateOpStr = oForm.DataSources.UserDataSources.Item("StrtDatOp2").ValueEx;
            string endDateOpStr = oForm.DataSources.UserDataSources.Item("EndDateOp2").ValueEx;

            string CardCode = oForm.DataSources.UserDataSources.Item("CCode2").ValueEx;
            string Attach = oForm.DataSources.UserDataSources.Item("Attach2").ValueEx;

            string needTax = oForm.DataSources.UserDataSources.Item("needTax").ValueEx;

            string query;

            string queryARInvoice = @"SELECT
                 CAST(""OINV"".""DocEntry"" as nvarchar(30)) as ""InvoiceEntry"",
                 '' as ""CreditMemoEntry"",
                 ""OINV"".""DocDate"",
                 ""OINV"".""CardCode"",
	             ""OINV"".""DocNum"",
	             CAST(""@BDO_TAXS"".""DocEntry"" AS nvarchar(30)) AS ""DocEntry"",
	             ""@BDO_TAXS"".""U_status"",
	             ""@BDO_TAXS"".""U_declNumber"",
	             ""@BDO_TAXS"".""U_series"",
	             ""@BDO_TAXS"".""U_number"",
	             ""@BDO_TAXS"".""U_invID"",
	             ""@BDO_TAXS"".""U_opDate"",
	             ""@BDO_TAXS"".""U_sentDate"",
	             ""@BDO_TAXS"".""U_confDate"",
	             ""@BDO_TAXS"".""U_comment"",
	             ""@BDO_TAXS"".""U_corrType"",
	             CAST(""@BDO_TAXS"".""U_corrDoc"" as nvarchar(30)) as ""U_corrDoc"",
	             SUM(""INV1"".""GTotal"") AS ""U_amount"",
	             SUM(""INV1"".""LineVat"") AS ""U_amountTX"",
	             ""OCRD"".""LicTradNum"",
                 ""OCRD"".""U_BDO_NotInv"" AS ""BDO_NotInv"",
                ""OCRD"".""CardName"",
	             'ARInvoice' as ""DocType"",
	             ""OINV"".""DocEntry"" as ""BaseDoc"" 
            FROM ""OINV"" 
            LEFT JOIN ""INV1"" ON ""OINV"".""DocEntry"" = ""INV1"".""DocEntry"" 
            LEFT JOIN ""OCRD"" ON ""OINV"".""CardCode"" = ""OCRD"".""CardCode"" 
            LEFT JOIN (SELECT
	             ""@BDO_TXS1"".""DocEntry"" AS ""DocEntry"",
	             ""@BDO_TXS1"".""U_baseDoc"" AS ""U_baseDoc"",
	             ""@BDO_TAXS"".""U_status"",
	             ""@BDO_TAXS"".""U_declNumber"",
	             ""@BDO_TAXS"".""U_series"",
	             ""@BDO_TAXS"".""U_number"",
	             ""@BDO_TAXS"".""U_invID"",
            	 ""@BDO_TAXS"".""U_opDate"",
            	 ""@BDO_TAXS"".""U_sentDate"",
            	 ""@BDO_TAXS"".""U_confDate"",
            	 ""@BDO_TAXS"".""U_comment"",
            	 ""@BDO_TAXS"".""U_corrType"",
            	 CASE WHEN ""@BDO_TAXS"".""U_corrInv"" = 'Y' 
            	THEN ""@BDO_TAXS"".""U_corrDoc"" 
            	ELSE '0' 
            	END AS ""U_corrDoc"" 
            	FROM ""@BDO_TXS1"" 
            	INNER JOIN ""@BDO_TAXS"" ON ""@BDO_TXS1"".""DocEntry"" = ""@BDO_TAXS"".""DocEntry"" 
            	WHERE ""@BDO_TXS1"".""U_baseDocT"" = 'ARInvoice' 
            	AND ""@BDO_TAXS"".""U_elctrnic"" = 'Y' 
            	AND ""@BDO_TAXS"".""U_downPaymnt"" = 'N' 
            	AND ""@BDO_TAXS"".""Canceled"" = 'N' 
            	AND ""@BDO_TAXS"".""U_status"" NOT IN ('canceled',
            	 'removed',
            	 'paper')) AS ""@BDO_TAXS"" ON ""OINV"".""DocEntry"" = ""@BDO_TAXS"".""U_baseDoc"" 
            WHERE ""CANCELED"" = 'N'";

            if (startDateOpStr != "")
            {
                queryARInvoice = queryARInvoice + @" AND ""OINV"".""DocDate"" >= '" + startDateOpStr + "' ";
            }

            if (endDateOpStr != "")
            {
                queryARInvoice = queryARInvoice + @" AND ""OINV"".""DocDate"" <= '" + endDateOpStr + "' ";
            }

            if (startDateStr != "")
            {
                queryARInvoice = queryARInvoice + @" AND ""@BDO_TAXS"".""U_sentDate"" >= '" + startDateStr + "' ";
            }

            if (endDateStr != "")
            {
                queryARInvoice = queryARInvoice + @" AND ""@BDO_TAXS"".""U_sentDate"" <= '" + endDateStr + "' ";
            }

            if (CardCode != "")
            {
                queryARInvoice = queryARInvoice + @" AND ""OINV"".""CardCode"" = '" + CardCode + "' ";
            }

            if (Attach != "0")
            {
                if (Attach == "1")
                {
                    queryARInvoice = queryARInvoice + @" AND ""@BDO_TAXS"".""DocEntry"" > 0 ";
                }
                else if (Attach == "2")
                {
                    queryARInvoice = queryARInvoice + @" AND ""@BDO_TAXS"".""DocEntry"" IS NULL ";
                }
            }

            if (needTax=="1")
            {
                queryARInvoice = queryARInvoice + @" AND ""OCRD"".""U_BDO_NotInv"" ='N' ";
            }

            queryARInvoice = queryARInvoice + @"GROUP BY ""OINV"".""DocEntry"",
            	 ""OINV"".""CardCode"",
            	 ""OINV"".""DocNum"",
            	 ""@BDO_TAXS"".""DocEntry"",
            	 ""@BDO_TAXS"".""U_status"",
            	 ""@BDO_TAXS"".""U_declNumber"",
            	 ""@BDO_TAXS"".""U_series"",
            	 ""@BDO_TAXS"".""U_number"",
            	 ""@BDO_TAXS"".""U_invID"",
            	 ""@BDO_TAXS"".""U_opDate"",
            	 ""@BDO_TAXS"".""U_sentDate"",
            	 ""@BDO_TAXS"".""U_comment"",
            	 ""@BDO_TAXS"".""U_confDate"",
            	 ""@BDO_TAXS"".""U_corrType"",
                 ""@BDO_TAXS"".""U_corrDoc"",
            	 ""OCRD"".""LicTradNum"",
                  ""OCRD"".""U_BDO_NotInv"",
                ""OCRD"".""CardName"",
            	 ""OINV"".""DocDate"",
            	 ""DocType""";

            //დაბრუნებები

            query = queryARInvoice + " UNION ALL ";

            string queryARCreditNote = @"SELECT
                 CAST(""ORIN"".""DocEntry"" AS nvarchar(30)) AS ""InvoiceEntry"",
                 '' as ""CreditMemoEntry"",
                 ""ORIN"".""DocDate"",
                 ""ORIN"".""CardCode"",
                 ""ORIN"".""DocNum"",
                 CAST(""@BDO_TAXS"".""DocEntry"" AS nvarchar(30)) AS ""DocEntry"",
                 ""@BDO_TAXS"".""U_status"",
                 ""@BDO_TAXS"".""U_declNumber"",
            	 ""@BDO_TAXS"".""U_series"",
            	 ""@BDO_TAXS"".""U_number"",
            	 ""@BDO_TAXS"".""U_invID"",
            	 ""@BDO_TAXS"".""U_opDate"",
            	 ""@BDO_TAXS"".""U_sentDate"",
            	 ""@BDO_TAXS"".""U_confDate"",
            	 ""@BDO_TAXS"".""U_comment"",
            	 ""@BDO_TAXS"".""U_corrType"",
            	 CAST(""@BDO_TAXS"".""U_corrDoc"" as nvarchar(30)) as ""U_corrDoc"",
            	 SUM(""RIN1"".""GTotal"") AS ""U_amount"",
            	 SUM(""RIN1"".""LineVat"") AS ""U_amountTX"",
            	 ""OCRD"".""LicTradNum"",
                 ""OCRD"".""U_BDO_NotInv"",
                    ""OCRD"".""CardName"",
            	 'ARCreditNote' AS ""DocType"",
            	 ""RIN1"".""BaseEntry"" 
            FROM ""ORIN"" 
            INNER JOIN ""RIN1"" ON ""ORIN"".""DocEntry"" = ""RIN1"".""DocEntry"" 
            LEFT JOIN ""OCRD"" ON ""ORIN"".""CardCode"" = ""OCRD"".""CardCode"" 
            LEFT JOIN (SELECT
            	 ""@BDO_TXS1"".""DocEntry"" AS ""DocEntry"",
            	 ""@BDO_TXS1"".""U_baseDoc"" AS ""U_baseDoc"",
            	 ""@BDO_TAXS"".""U_status"",
            	 ""@BDO_TAXS"".""U_declNumber"",
            	 ""@BDO_TAXS"".""U_series"",
            	 ""@BDO_TAXS"".""U_number"",
            	 ""@BDO_TAXS"".""U_invID"",
            	 ""@BDO_TAXS"".""U_opDate"",
            	 ""@BDO_TAXS"".""U_sentDate"",
            	 ""@BDO_TAXS"".""U_confDate"",
            	 ""@BDO_TAXS"".""U_comment"",
            	 ""@BDO_TAXS"".""U_corrType"",
            	 CASE WHEN ""@BDO_TAXS"".""U_corrInv"" = 'Y' 
            	THEN ""@BDO_TAXS"".""U_corrDoc"" 
            	ELSE '0' 
            	END AS ""U_corrDoc"" 
            	FROM ""@BDO_TXS1"" 
            	INNER JOIN ""@BDO_TAXS"" ON ""@BDO_TXS1"".""DocEntry"" = ""@BDO_TAXS"".""DocEntry"" 
            	WHERE ""@BDO_TXS1"".""U_baseDocT"" = 'ARCreditNote' 
            	AND ""@BDO_TAXS"".""U_elctrnic"" = 'Y' 
            	AND ""@BDO_TAXS"".""U_downPaymnt"" = 'N' 
            	AND ""@BDO_TAXS"".""Canceled"" = 'N' 
            	AND ""@BDO_TAXS"".""U_status"" NOT IN ('canceled',
            	 'removed',
            	 'paper')) AS ""@BDO_TAXS"" ON ""ORIN"".""DocEntry"" = ""@BDO_TAXS"".""U_baseDoc"" 
            WHERE ""CANCELED"" = 'N' AND ""RIN1"".""BaseType"" <> '203'";
            //AND ""Posted"" = 'N'";

            if (startDateOpStr != "")
            {
                queryARCreditNote = queryARCreditNote + @" AND ""ORIN"".""DocDate"" >= '" + startDateOpStr + "' ";
            }

            if (endDateOpStr != "")
            {
                queryARCreditNote = queryARCreditNote + @" AND ""ORIN"".""DocDate"" <= '" + endDateOpStr + "' ";
            }

            if (startDateStr != "")
            {
                queryARCreditNote = queryARCreditNote + @" AND ""@BDO_TAXS"".""U_sentDate"" >= '" + startDateStr + "' ";
            }

            if (endDateStr != "")
            {
                queryARCreditNote = queryARCreditNote + @" AND ""@BDO_TAXS"".""U_sentDate"" <= '" + endDateStr + "' ";
            }

            if (CardCode != "")
            {
                queryARCreditNote = queryARCreditNote + @" AND ""ORIN"".""CardCode"" = '" + CardCode + "' ";
            }

            if (Attach != "0")
            {
                if (Attach == "1")
                {
                    queryARCreditNote = queryARCreditNote + @" AND ""@BDO_TAXS"".""DocEntry"" > 0 ";
                }
                else if (Attach == "2")
                {
                    queryARCreditNote = queryARCreditNote + @" AND ""@BDO_TAXS"".""DocEntry"" IS NULL ";
                }
            }

            if (needTax == "1")
            {
                queryARCreditNote = queryARCreditNote + @" AND ""OCRD"".""U_BDO_NotInv"" ='N' ";
            }

            queryARCreditNote = queryARCreditNote + @"GROUP BY ""ORIN"".""DocEntry"",
                      ""ORIN"".""CardCode"",
                      ""ORIN"".""DocNum"",
                      ""@BDO_TAXS"".""DocEntry"",
                      ""@BDO_TAXS"".""U_status"",
                      ""@BDO_TAXS"".""U_declNumber"",
                      ""@BDO_TAXS"".""U_series"",
                      ""@BDO_TAXS"".""U_number"",
                      ""@BDO_TAXS"".""U_invID"",
                      ""@BDO_TAXS"".""U_opDate"",
                      ""@BDO_TAXS"".""U_sentDate"",
                      ""@BDO_TAXS"".""U_comment"",
                      ""@BDO_TAXS"".""U_confDate"",
                      ""@BDO_TAXS"".""U_corrType"",
                      ""@BDO_TAXS"".""U_corrDoc"",
                      ""OCRD"".""LicTradNum"",
                      ""OCRD"".""U_BDO_NotInv"",
                      ""OCRD"".""CardName"",
                      ""ORIN"".""DocDate"", 
                      ""DocType"", 
                      ""RIN1"".""BaseEntry"" ";

            //გაცემული ავანსები

            query = query + queryARCreditNote + " UNION ALL ";

            string queryARDownPaymentVAT = @"SELECT
                 CAST(""@BDOSARDV"".""DocEntry"" AS nvarchar(30)) AS ""InvoiceEntry"",
	             '' AS ""CreditMemoEntry"",
            	 ""@BDOSARDV"".""U_DocDate"",
            	 ""@BDOSARDV"".""U_cardCode"",
            	 ""@BDOSARDV"".""DocNum"",
            	 CAST(""@BDO_TAXS"".""DocEntry"" AS nvarchar(30)) AS ""DocEntry"",
            	 ""@BDO_TAXS"".""U_status"",
            	 ""@BDO_TAXS"".""U_declNumber"",
            	 ""@BDO_TAXS"".""U_series"",
            	 ""@BDO_TAXS"".""U_number"",
            	 ""@BDO_TAXS"".""U_invID"",
            	 ""@BDO_TAXS"".""U_opDate"",
            	 ""@BDO_TAXS"".""U_sentDate"",
            	 ""@BDO_TAXS"".""U_confDate"",
            	 ""@BDO_TAXS"".""U_comment"",
            	 ""@BDO_TAXS"".""U_corrType"",
            	 CAST(""@BDO_TAXS"".""U_corrDoc"" as nvarchar(30)) as ""U_corrDoc"",
            	 SUM(""@BDOSARDV"".""U_GrsAmnt"") AS ""U_amount"",
            	 SUM(""@BDOSARDV"".""U_VatAmount"") AS ""U_amountTX"",
            	 ""OCRD"".""LicTradNum"",
""OCRD"".""U_BDO_NotInv"",
                      ""OCRD"".""CardName"",
            	 'ARDownPaymentVAT' AS ""DocType"",
            	 ""@BDOSARDV"".""DocEntry"" AS ""BaseDoc"" 
            FROM ""@BDOSARDV"" 
            
            LEFT JOIN ""OCRD"" ON ""@BDOSARDV"".""U_cardCode"" = ""OCRD"".""CardCode"" 
            LEFT JOIN (SELECT
            	 ""@BDO_TXS1"".""DocEntry"" AS ""DocEntry"",
            	 ""@BDO_TXS1"".""U_baseDoc"" AS ""U_baseDoc"",
            	 ""@BDO_TAXS"".""U_status"",
            	 ""@BDO_TAXS"".""U_declNumber"",
            	 ""@BDO_TAXS"".""U_series"",
            	 ""@BDO_TAXS"".""U_number"",
            	 ""@BDO_TAXS"".""U_invID"",
            	 ""@BDO_TAXS"".""U_opDate"",
            	 ""@BDO_TAXS"".""U_sentDate"",
            	 ""@BDO_TAXS"".""U_confDate"",
            	 ""@BDO_TAXS"".""U_comment"",
            	 ""@BDO_TAXS"".""U_corrType"",
            	 CASE WHEN ""@BDO_TAXS"".""U_corrInv"" = 'Y' 
            	THEN ""@BDO_TAXS"".""U_corrDoc"" 
            	ELSE '0' 
            	END AS ""U_corrDoc"" 
            	FROM ""@BDO_TXS1"" 
            	INNER JOIN ""@BDO_TAXS"" ON ""@BDO_TXS1"".""DocEntry"" = ""@BDO_TAXS"".""DocEntry"" 
            	WHERE ""@BDO_TXS1"".""U_baseDocT"" = 'ARDownPaymentVAT' 
            	AND ""@BDO_TAXS"".""U_elctrnic"" = 'Y' 
            	AND ""@BDO_TAXS"".""U_downPaymnt"" = 'Y' 
            	AND ""@BDO_TAXS"".""Canceled"" = 'N' 
            	AND ""@BDO_TAXS"".""U_status"" NOT IN ('canceled',
            	 'removed',
            	 'paper')) AS ""@BDO_TAXS"" ON ""@BDOSARDV"".""DocEntry"" = ""@BDO_TAXS"".""U_baseDoc"" 
            WHERE ""Canceled"" = 'N'";
            //AND ""Posted"" = 'Y'

            if (startDateOpStr != "")
            {
                queryARDownPaymentVAT = queryARDownPaymentVAT + @" AND ""@BDOSARDV"".""U_DocDate"" >= '" + startDateOpStr + "' ";
            }

            if (endDateOpStr != "")
            {
                queryARDownPaymentVAT = queryARDownPaymentVAT + @" AND ""@BDOSARDV"".""U_DocDate"" <= '" + endDateOpStr + "' ";
            }

            if (startDateStr != "")
            {
                queryARDownPaymentVAT = queryARDownPaymentVAT + @" AND ""@BDO_TAXS"".""U_sentDate"" >= '" + startDateStr + "' ";
            }

            if (endDateStr != "")
            {
                queryARDownPaymentVAT = queryARDownPaymentVAT + @" AND ""@BDO_TAXS"".""U_sentDate"" <= '" + endDateStr + "' ";
            }

            if (CardCode != "")
            {
                queryARDownPaymentVAT = queryARDownPaymentVAT + @" AND ""@BDOSARDV"".""U_cardCode"" = '" + CardCode + "' ";
            }

            if (Attach != "0")
            {
                if (Attach == "1")
                {
                    queryARDownPaymentVAT = queryARDownPaymentVAT + @" AND ""@BDO_TAXS"".""DocEntry"" > 0 ";
                }
                else if (Attach == "2")
                {
                    queryARDownPaymentVAT = queryARDownPaymentVAT + @" AND ""@BDO_TAXS"".""DocEntry"" IS NULL ";
                }
            }

            if (needTax == "1")
            {
                queryARDownPaymentVAT = queryARDownPaymentVAT + @" AND ""OCRD"".""U_BDO_NotInv"" ='N' ";
            }

            queryARDownPaymentVAT = queryARDownPaymentVAT + @"GROUP BY ""@BDOSARDV"".""DocEntry"",
            	 ""@BDOSARDV"".""U_cardCode"",
            	 ""@BDOSARDV"".""DocNum"",
            	 ""@BDO_TAXS"".""DocEntry"",
            	 ""@BDO_TAXS"".""U_status"",
            	 ""@BDO_TAXS"".""U_declNumber"",
            	 ""@BDO_TAXS"".""U_series"",
            	 ""@BDO_TAXS"".""U_number"",
            	 ""@BDO_TAXS"".""U_invID"",
            	 ""@BDO_TAXS"".""U_opDate"",
            	 ""@BDO_TAXS"".""U_sentDate"",
            	 ""@BDO_TAXS"".""U_comment"",
            	 ""@BDO_TAXS"".""U_confDate"",
            	 ""@BDO_TAXS"".""U_corrType"",
                 ""@BDO_TAXS"".""U_corrDoc"",
            	 ""OCRD"".""LicTradNum"",
                ""OCRD"".""U_BDO_NotInv"",
                      ""OCRD"".""CardName"",
            	 ""@BDOSARDV"".""U_DocDate"",
            	 ""OCRD"".""CardCode""";

            //სორტირება თარიღის მიხედვით
            //query = query + " ORDER BY  \"CardCode\",\"OINV\".\"DocDate\",\"BaseDoc\",\"DocType\" DESC ";
            query = query + queryARDownPaymentVAT + @" ORDER BY  ""CardCode"", ""BaseDoc"", ""DocType"" DESC ";

            try
            {
                oRecordSet.DoQuery(query);

                oDataTable.Rows.Clear();

                Dictionary<string, string> rsSettings = CompanyDetails.getRSSettings( out errorText);
                if (errorText != null)
                {
                    oForm.Update();
                    Program.uiApp.StatusBar.SetSystemMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return;
                }

                string su = rsSettings["SU"];
                string sp = rsSettings["SP"];

                TaxInvoice oTaxInvoice = new TaxInvoice(rsSettings["ProtocolType"]);

                bool chek_service_user = oTaxInvoice.check_usr( su, sp, out errorText);
                if (chek_service_user == false)
                {
                    oForm.Update();
                    errorText = BDOSResources.getTranslate("ServiceUserPasswordNotCorrect");
                    Program.uiApp.StatusBar.SetSystemMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return;
                }

                int rowIndex = 0;
                string declNum = null;
                string docEntry = "";
                string errorTextWb = null;
                string errorTextGoods = null;
                string statusRS = "";
                string number = null;
                string corrInvID = null;
                string corrDocEntry = "";

                while (!oRecordSet.EoF)
                {
                    declNum = oRecordSet.Fields.Item("U_declNumber").Value;
                    docEntry = oRecordSet.Fields.Item("DocEntry").Value;
                    corrDocEntry = oRecordSet.Fields.Item("U_corrDoc").Value;
                    corrDocEntry = corrDocEntry == "0" ? "" : corrDocEntry;
                    number = oRecordSet.Fields.Item("U_number").Value;

                    if (download == false & String.IsNullOrEmpty(number) == false)
                    {
                        BDO_TaxInvoiceSent.operationRS( oTaxInvoice, "checkSync", Convert.ToInt32(docEntry), -1, new DateTime(), out errorText, out errorTextWb, out errorTextGoods);

                        statusRS = BDO_TaxInvoiceSent.getStatusValueByStatusNumber(statusRS, false, false);
                    }
                    statusRS = download == false ? statusRS : oRecordSet.Fields.Item("U_status").Value;

                    //----------------------------->კორექტირების შევსება<-----------------------------
                    string CreditMemoEntry = oRecordSet.Fields.Item("CreditMemoEntry").Value;
                    string DocType = oRecordSet.Fields.Item("DocType").Value;

                    if (CreditMemoEntry != "" && DocType != "ARDownPaymentVAT")
                    {
                        string BaseDoc = oRecordSet.Fields.Item("BaseDoc").Value.ToString();
                        DateTime DocDate = oRecordSet.Fields.Item("DocDate").Value;
                        string queryCardCode = oRecordSet.Fields.Item("CardCode").Value;

                        DataTable сreditNotes = BDO_TaxInvoiceSent.getCorrDocs( Convert.ToInt32(CreditMemoEntry), DocDate, DocDate, Convert.ToInt32(BaseDoc), queryCardCode, out errorText);
                        if (сreditNotes != null)
                        {
                            int rowCount = сreditNotes.Rows.Count;
                            Dictionary<string, object> taxDocInfo = null;

                            if (rowCount == 1) // ეს ნიშნავს რომ რეალიზაციაზე მარტო ერთი сreditNote არის მიბმული. ამიტომ უნდა დავაკორექტიროთ რეალიზაციის ა/ფ.
                            {
                                taxDocInfo = BDO_TaxInvoiceSent.getTaxInvoiceSentDocumentInfo( Convert.ToInt32(BaseDoc), "ARInvoice", queryCardCode, out errorText);
                                if (taxDocInfo != null)
                                {
                                    corrDocEntry = taxDocInfo["docEntry"].ToString(); //კორექტირების ა/ფ-ის Entry
                                    if (corrDocEntry == "0")
                                    {
                                        corrDocEntry = "";
                                    }
                                    corrInvID = taxDocInfo["invID"].ToString(); //კორექტირების ა/ფ-ის ID
                                }
                            }
                            else
                            {
                                DataRow taxDataRow = сreditNotes.Rows[1]; //ვიღებთ მეორე სტრიქონს, იმიტომ რომ 1 სტრიქონში ამოდის თვითონ ეს сreditNote 

                                SAPbobsCOM.Documents oCreditNote = (SAPbobsCOM.Documents)taxDataRow["creditNote"];
                                corrDocEntry = taxDataRow["corrDocEntry"].ToString(); //კორექტირების ა/ფ-ის Entry
                                if (corrDocEntry == "0")
                                {
                                    corrDocEntry = "";
                                }
                                corrInvID = taxDataRow["invID"].ToString(); //კორექტირების ა/ფ-ის ID
                            }
                        }
                    }
                    else if (DocType == "ARDownPaymentVAT" && string.IsNullOrEmpty(corrDocEntry) == true)
                    {
                        bool primary;
                        int invoiceEntry = Convert.ToInt32(oRecordSet.Fields.Item("InvoiceEntry").Value);
                        DataTable confirmedInvoices;
                        DateTime docDate = oRecordSet.Fields.Item("DocDate").Value;

                        if (BDOSARDownPaymentVATAccrual.checkDocumentForTaxInvoice( invoiceEntry, docDate, docDate, out primary, out confirmedInvoices, out errorText) == true)
                        {
                            if (primary == false)
                            {
                                DataRow taxDataRow = confirmedInvoices.Rows[0];
                                corrDocEntry = taxDataRow["InvDocEntry"].ToString();
                            }
                        }
                    }
                    //----------------------------->კორექტირების შევსება<-----------------------------

                    oDataTable.Rows.Add();

                    oDataTable.SetValue("LineNum", rowIndex, rowIndex + 1);
                    oDataTable.SetValue("Status", rowIndex, statusRS);
                    oDataTable.SetValue("StatusDoc", rowIndex, oRecordSet.Fields.Item("U_status").Value);
                    oDataTable.SetValue("DeclStatus", rowIndex, (String.IsNullOrEmpty(declNum) == true) ? BDOSResources.getTranslate("WithoutDeclaration") : BDOSResources.getTranslate("WithDeclaration"));
                    oDataTable.SetValue("DeclNum", rowIndex, oRecordSet.Fields.Item("U_declNumber").Value);
                    oDataTable.SetValue("Document", rowIndex, docEntry);
                    oDataTable.SetValue("TxSerie", rowIndex, oRecordSet.Fields.Item("U_series").Value);
                    oDataTable.SetValue("TxNum", rowIndex, number);
                    oDataTable.SetValue("TxID", rowIndex, oRecordSet.Fields.Item("U_invID").Value);
                    oDataTable.SetValue("IsVATPayer", rowIndex, oRecordSet.Fields.Item("BDO_NotInv").Value);
                    oDataTable.SetValue("CardCode", rowIndex, oRecordSet.Fields.Item("CardCode").Value);
                    oDataTable.SetValue("CardName", rowIndex, oRecordSet.Fields.Item("CardName").Value);
                    oDataTable.SetValue("OpDate", rowIndex, oRecordSet.Fields.Item("DocDate").Value.ToString("yyyyMMdd") == "18991230" ? "" : oRecordSet.Fields.Item("DocDate").Value.ToString("yyyyMMdd"));
                    oDataTable.SetValue("RegDate", rowIndex, oRecordSet.Fields.Item("U_sentDate").Value.ToString("yyyyMMdd") == "18991230" ? "" : oRecordSet.Fields.Item("U_sentDate").Value.ToString("yyyyMMdd"));
                    oDataTable.SetValue("ConfDate", rowIndex, oRecordSet.Fields.Item("U_confDate").Value.ToString("yyyyMMdd") == "18991230" ? "" : oRecordSet.Fields.Item("U_confDate").Value.ToString("yyyyMMdd"));
                    oDataTable.SetValue("Sum", rowIndex, oRecordSet.Fields.Item("U_amount").Value);
                    oDataTable.SetValue("VatSum", rowIndex, oRecordSet.Fields.Item("U_amountTX").Value);
                    oDataTable.SetValue("Comment", rowIndex, oRecordSet.Fields.Item("U_comment").Value);
                    oDataTable.SetValue("TxChkBx", rowIndex, "N");
                    oDataTable.SetValue("Vatno", rowIndex, oRecordSet.Fields.Item("LicTradNum").Value);
                    oDataTable.SetValue("InvoiceEntry", rowIndex, oRecordSet.Fields.Item("InvoiceEntry").Value);
                    if (DocType == "ARCreditNote" && oRecordSet.Fields.Item("BaseDoc").Value != 0)
                        oDataTable.SetValue("BaseARInvoice", rowIndex, oRecordSet.Fields.Item("BaseDoc").Value);
                    oDataTable.SetValue("CorrDoc", rowIndex, corrDocEntry);
                    //oDataTable.SetValue("CreditMemoEntry", rowIndex, oRecordSet.Fields.Item("CreditMemoEntry").Value);
                    oDataTable.SetValue("CorrType", rowIndex, oRecordSet.Fields.Item("U_corrType").Value);
                    oDataTable.SetValue("DocNum", rowIndex, oRecordSet.Fields.Item("DocNum").Value);
                    oDataTable.SetValue("DocType", rowIndex, DocType);

                    oRecordSet.MoveNext();
                    rowIndex++;
                }

                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("TxTable2").Specific;
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
            }
        }
        //----------------------------->Tax Invoice Sent<-----------------------------


        //----------------------------->Tax Invoice Received<-----------------------------
        public static void rsOperationTaxInvoiceReceived( SAPbouiCOM.Form oForm,  int oOperation, out string errorText)
        {
            errorText = null;
            oForm.Freeze(true);
            bool opSuccess = true;

            Dictionary<string, string> rsSettings = CompanyDetails.getRSSettings( out errorText);
            if (errorText != null)
            {
                oForm.Update();
                oForm.Freeze(false);
                Program.uiApp.StatusBar.SetSystemMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return;
            }

            string su = rsSettings["SU"];
            string sp = rsSettings["SP"];

            TaxInvoice oTaxInvoice = new TaxInvoice(rsSettings["ProtocolType"]);

            bool chek_service_user = oTaxInvoice.check_usr( su, sp, out errorText);
            if (chek_service_user == false)
            {
                oForm.Update();
                oForm.Freeze(false);
                errorText = BDOSResources.getTranslate("ServiceUserPasswordNotCorrect");
                Program.uiApp.StatusBar.SetSystemMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return;
            }
            //oForm.Freeze(true);
            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("TxTable").Specific));
            int rowCount = oMatrix.RowCount;
            for (int row = 1; row <= rowCount; row++)
            {
                SAPbouiCOM.CheckBox Edtfield = oMatrix.Columns.Item("TxChkBx").Cells.Item(row).Specific;
                bool checkedLine = (Edtfield.Checked);

                if (checkedLine)
                {
                    int TxId = Convert.ToInt32(oMatrix.Columns.Item("TxID").Cells.Item(row).Specific.Value);
                    int docEntry = Convert.ToInt32(oMatrix.Columns.Item("Document").Cells.Item(row).Specific.Value);
                    string statusDoc = oMatrix.Columns.Item("StatusDoc").Cells.Item(row).Specific.Value;
                    //oMatrix.Columns.Item("TxChkBx").Cells.Item(row).Specific.Checked = false;
                    string opText = "";
                    string statusRS;

                    if (oOperation == 0) //დადასტურება
                    {
                        opText = BDOSResources.getTranslate("RSConfirm");

                        if (statusDoc == "received" || statusDoc == "correctionReceived" || statusDoc == "cancellationProcess")
                        {
                            BDO_TaxInvoiceReceived.operationRS( oTaxInvoice, "confirmation", Convert.ToInt32(docEntry), -1, new DateTime(), null, out statusRS, out errorText);
                        }
                        else
                        {
                            errorText = BDOSResources.getTranslate("Operation") + " " + opText + " " + BDOSResources.getTranslate("NotBeComplete") + " " + BDOSResources.getTranslate("DocumentStatusMustBe") + " " + BDOSResources.getTranslate("Received") + " " + BDOSResources.getTranslate("Or") + " " + BDOSResources.getTranslate("CorrectionReceived") + " " + BDOSResources.getTranslate("Or") + " " + BDOSResources.getTranslate("CancellationProcess") + "!";
                            opSuccess = false;
                        }
                    }
                    else if (oOperation == 1) //სტატუსების განახლება
                    {
                        opText = BDOSResources.getTranslate("RSUpdateStatus");
                        BDO_TaxInvoiceReceived.operationRS( oTaxInvoice, "updateStatus", Convert.ToInt32(docEntry), -1, new DateTime(), null, out statusRS, out errorText);
                    }
                    else if (oOperation == 2) //უარყოფა
                    {
                        opText = BDOSResources.getTranslate("RSDeny");
                        if (statusDoc == "received" || statusDoc == "correctionReceived")
                        {
                            BDO_TaxInvoiceReceived.operationRS( oTaxInvoice, "deny", Convert.ToInt32(docEntry), -1, new DateTime(), null, out statusRS, out errorText);
                        }
                        else
                        {
                            errorText = BDOSResources.getTranslate("Operation") + " " + opText + " " + BDOSResources.getTranslate("NotBeComplete") + " " + BDOSResources.getTranslate("DocumentStatusMustBe") + " " + BDOSResources.getTranslate("Received") + " " + BDOSResources.getTranslate("Or") + " " + BDOSResources.getTranslate("CorrectionReceived") + "!";
                            opSuccess = false;
                        }
                    }
                    if (errorText != null)
                    {
                        Program.uiApp.StatusBar.SetSystemMessage(errorText + " " + BDOSResources.getTranslate("TableRow") + ": " + row.ToString() + ", " + BDOSResources.getTranslate("TaxInvoiceID") + " : " + TxId, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);

                        opSuccess = false;
                        continue;
                    }
                    else
                    {
                        Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("Operation") + "  " + opText + BDOSResources.getTranslate("DoneSuccessfully") + " " + BDOSResources.getTranslate("TableRow") + " : " + row.ToString() + ", " + BDOSResources.getTranslate("TaxInvoiceID") + " : " + TxId, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                        updateMatrixRowTaxInvoiceReceived( oMatrix, row, docEntry);
                    }
                }
            }
            if (opSuccess == false)
            {
                Program.uiApp.MessageBox(BDOSResources.getTranslate("OperationForSomeDocumentFinishedWithErrorSeeLog"));
            }
            oForm.Update();
            oForm.Freeze(false);
        }

        private static void updateMatrixRowTaxInvoiceReceived( SAPbouiCOM.Matrix oMatrix, int row, int docEntry)
        {
            NumberFormatInfo Nfi = new NumberFormatInfo() { NumberDecimalSeparator = "." };

            SAPbobsCOM.CompanyService oCompanyService = null;
            SAPbobsCOM.GeneralService oGeneralService = null;
            SAPbobsCOM.GeneralData oGeneralData = null;
            SAPbobsCOM.GeneralDataParams oGeneralParams = null;
            oCompanyService = Program.oCompany.GetCompanyService();
            oGeneralService = oCompanyService.GetGeneralService("UDO_F_BDO_TAXR_D");
            oGeneralData = ((SAPbobsCOM.GeneralData)(oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)));

            oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
            oGeneralParams.SetProperty("DocEntry", docEntry);
            oGeneralData = oGeneralService.GetByParams(oGeneralParams);

            string declNumber = oGeneralData.GetProperty("U_declNumber");
            string corrInv = oGeneralData.GetProperty("U_corrInv");
            double amount = oGeneralData.GetProperty("U_amount"); //თანხა დღგ-ის ჩათვლით
            double amountTX = oGeneralData.GetProperty("U_amountTX"); //დღგ-ის თანხა
            string corrDocEntry = oGeneralData.GetProperty("U_corrDTxt");

            //if (corrInv == "Y")
            //{
            //    amount = oGeneralData.GetProperty("U_amtACor"); //თანხა დღგ-ის ჩათვლით
            //    amountTX = oGeneralData.GetProperty("U_amtTXACr"); //დღგ-ის თანხა               
            //}

            string opDate = oGeneralData.GetProperty("U_opDate").ToString("yyyyMMdd") == "18991230" ? "" : oGeneralData.GetProperty("U_opDate").ToString("yyyyMMdd");
            string recvDate = oGeneralData.GetProperty("U_recvDate").ToString("yyyyMMdd") == "18991230" ? "" : oGeneralData.GetProperty("U_recvDate").ToString("yyyyMMdd");
            string confDate = oGeneralData.GetProperty("U_confDate").ToString("yyyyMMdd") == "18991230" ? "" : oGeneralData.GetProperty("U_confDate").ToString("yyyyMMdd");

            oMatrix.Columns.Item("LineNum").Cells.Item(row).Specific.Value = row;
            //oMatrix.Columns.Item("SrvStatus").Cells.Item(row).Specific.Value = oGeneralData.GetProperty("U_status"); ???
            oMatrix.Columns.Item("Status").Cells.Item(row).Specific.Select(oGeneralData.GetProperty("U_status"), SAPbouiCOM.BoSearchKey.psk_ByValue);
            oMatrix.Columns.Item("StatusDoc").Cells.Item(row).Specific.Select(oGeneralData.GetProperty("U_status"), SAPbouiCOM.BoSearchKey.psk_ByValue);
            oMatrix.Columns.Item("TxID").Cells.Item(row).Specific.Value = oGeneralData.GetProperty("U_invID");
            oMatrix.Columns.Item("TxSerie").Cells.Item(row).Specific.Value = oGeneralData.GetProperty("U_series");
            oMatrix.Columns.Item("TxNum").Cells.Item(row).Specific.Value = oGeneralData.GetProperty("U_number");
            oMatrix.Columns.Item("Document").Cells.Item(row).Specific.Value = docEntry;
            oMatrix.Columns.Item("CorrDoc").Cells.Item(row).Specific.Value = corrDocEntry;
            oMatrix.Columns.Item("DeclNum").Cells.Item(row).Specific.Value = declNumber;
            oMatrix.Columns.Item("DeclStatus").Cells.Item(row).Specific.Value = (String.IsNullOrEmpty(declNumber) == true) ? BDOSResources.getTranslate("WithoutDeclaration") : BDOSResources.getTranslate("WithDeclaration");
            oMatrix.Columns.Item("CardCode").Cells.Item(row).Specific.Value = oGeneralData.GetProperty("U_cardCode");
            oMatrix.Columns.Item("OpDate").Cells.Item(row).Specific.Value = opDate;
            oMatrix.Columns.Item("RegDate").Cells.Item(row).Specific.Value = confDate;
            oMatrix.Columns.Item("ConfDate").Cells.Item(row).Specific.Value = recvDate;
            oMatrix.Columns.Item("Sum").Cells.Item(row).Specific.Value = amount.ToString(Nfi);
            oMatrix.Columns.Item("VatSum").Cells.Item(row).Specific.Value = amountTX.ToString(Nfi);
            oMatrix.Columns.Item("Comment").Cells.Item(row).Specific.Value = oGeneralData.GetProperty("U_comment");
            //oMatrix.Columns.Item("TxChkBx").Cells.Item(row).Specific.Checked = false;
            //oMatrix.Columns.Item("Vatno").Cells.Item(row).Specific.Value = oGeneralData.GetProperty("U_cardCodeT");  ???
            oMatrix.CommonSetting.SetCellBackColor(row, 2, FormsB1.getLongIntRGB(231, 231, 231));
            oMatrix.CommonSetting.SetCellBackColor(row, 3, FormsB1.getLongIntRGB(231, 231, 231));
        }

        public static void updateTaxInvoiceReceived(  SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            oForm.Freeze(true);
            bool opSuccess = true;

            Dictionary<string, string> rsSettings = CompanyDetails.getRSSettings( out errorText);
            if (errorText != null)
            {
                oForm.Update();
                oForm.Freeze(false);
                Program.uiApp.StatusBar.SetSystemMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return;
            }

            string su = rsSettings["SU"];
            string sp = rsSettings["SP"];

            TaxInvoice oTaxInvoice = new TaxInvoice(rsSettings["ProtocolType"]);

            bool chek_service_user = oTaxInvoice.check_usr( su, sp, out errorText);
            if (chek_service_user == false)
            {
                oForm.Update();
                oForm.Freeze(false);
                errorText = BDOSResources.getTranslate("ServiceUserPasswordNotCorrect");
                Program.uiApp.StatusBar.SetSystemMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return;
            }

            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("TxTable").Specific));

            oMatrix.FlushToDataSource();

            int rowCount = oMatrix.RowCount;
            string statusRS = null;

            for (int row = 1; row <= rowCount; row++)
            {
                SAPbouiCOM.CheckBox Edtfield = oMatrix.Columns.Item("TxChkBx").Cells.Item(row).Specific;
                bool checkedLine = (Edtfield.Checked);

                if (checkedLine)
                {
                    string StatusDoc = oMatrix.Columns.Item("StatusDoc").Cells.Item(row).Specific.Value;
                    //oMatrix.Columns.Item("TxChkBx").Cells.Item(row).Specific.Checked = false;

                    if (StatusDoc == "correctionConfirmed" || StatusDoc == "confirmed")
                    {
                        errorText = BDOSResources.getTranslate("UpdateConfirmedTaxInvoiceNotAllowed") + "! ";
                        Program.uiApp.StatusBar.SetSystemMessage(errorText + BDOSResources.getTranslate("TableRow") + " " + row.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                        opSuccess = false;
                        continue;
                    }

                    int docEntry = Convert.ToInt32(oMatrix.Columns.Item("Document").Cells.Item(row).Specific.Value);
                    BDO_TaxInvoiceReceived.operationRS( oTaxInvoice, "update", docEntry, -1, new DateTime(), null, out statusRS, out errorText);
                    if (errorText != null)
                    {
                        Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("ErrorWhileDocumentEdit") + " " + BDOSResources.getTranslate("TableRow") + " : " + row.ToString() + "! მიზეზით : " + errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        opSuccess = false;
                        continue;
                    }
                    else
                    {
                        updateMatrixRowTaxInvoiceReceived( oMatrix, row, docEntry);
                    }
                }
            }

            if (opSuccess == false)
            {
                Program.uiApp.MessageBox(BDOSResources.getTranslate("OperationForSomeDocumentFinishedWithErrorSeeLog"));
            }
            else
            {
                Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("TaxInvoicesUpdated"), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
            oForm.Update();
            oForm.Freeze(false);
        }

        public static void downloadTaxInvoiceReceived(  SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            int answer = Program.uiApp.MessageBox(BDOSResources.getTranslate("DoYouWantForDocuementsExecuteUpdate"), 1, BDOSResources.getTranslate("Yes"), BDOSResources.getTranslate("No"), "");
            bool isUpdate = answer == 1 ? true : false;

            Dictionary<string, string> rsSettings = CompanyDetails.getRSSettings( out errorText);
            if (errorText != null)
            {
                Program.uiApp.StatusBar.SetSystemMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return;
            }

            string su = rsSettings["SU"];
            string sp = rsSettings["SP"];

            TaxInvoice oTaxInvoice = new TaxInvoice(rsSettings["ProtocolType"]);

            bool chek_service_user = oTaxInvoice.check_usr( su, sp, out errorText);
            if (chek_service_user == false)
            {
                errorText = BDOSResources.getTranslate("ServiceUserPasswordNotCorrect");
                Program.uiApp.StatusBar.SetSystemMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return;
            }

            DateTime OperationPeriodStart = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            DateTime OperationPeriodEnd = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month));
            OperationPeriodEnd = OperationPeriodEnd.AddSeconds(24 * 3600 - 1);

            string startDateStr = oForm.DataSources.UserDataSources.Item("StartDate").ValueEx;
            DateTime startDate = FormsB1.DateFormats(startDateStr, "yyyyMMdd");

            string endDateStr = oForm.DataSources.UserDataSources.Item("EndDate").ValueEx;
            DateTime endDate = FormsB1.DateFormats(endDateStr, "yyyyMMdd") == new DateTime() ? OperationPeriodEnd : FormsB1.DateFormats(endDateStr, "yyyyMMdd").AddDays(1).AddSeconds(-1);

            string startDateOpStr = oForm.DataSources.UserDataSources.Item("StartDatOp").ValueEx;
            DateTime startDateOp = FormsB1.DateFormats(startDateOpStr, "yyyyMMdd") == new DateTime() ? new DateTime() : FormsB1.DateFormats(startDateOpStr, "yyyyMMdd");

            string endDateOpStr = oForm.DataSources.UserDataSources.Item("EndDateOp").ValueEx;
            DateTime endDateOp = FormsB1.DateFormats(endDateOpStr, "yyyyMMdd") == new DateTime() ? OperationPeriodEnd : FormsB1.DateFormats(endDateOpStr, "yyyyMMdd").AddDays(1).AddSeconds(-1);

            SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("TxTable");
            oDataTable.Rows.Clear();

            SAPbobsCOM.BusinessPartners oBP;
            oBP = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);
            oBP.GetByKey(oForm.DataSources.UserDataSources.Item("CCode").Value);

            string sa_identNum = oBP.UserFields.Fields.Item("LicTradNum").Value;

            DataTable TaxDataTable = oTaxInvoice.get_buyer_invoices(startDate, endDate, startDateOp, endDateOp, "", sa_identNum, "", "", out errorText);

            int rowCount = TaxDataTable.Rows.Count;

            for (int row = 0; row < rowCount; row++)
            {
                DataRow TaxDataRow = TaxDataTable.Rows[row];
                BDO_TaxInvoiceReceived.createDocumentTaxInvoiceType( oTaxInvoice, isUpdate, TaxDataRow, out errorText);

                if (errorText != null)
                {
                    Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("ErrorDocumentAddEdit") + " : " + errorText + " " + BDOSResources.getTranslate("TableRow") + " : " + row.ToString());
                }
            }

            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("TxTable").Specific;
            oForm.Freeze(true);
            oMatrix.Clear();
            oMatrix.LoadFromDataSource();
            oMatrix.AutoResizeColumns();
            oForm.Update();
            oForm.Freeze(false);

            Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("RSTaxInvoicesDownloaded"), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
        }

        public static void addDeclTaxInvoiceReceived(  SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            oForm.Freeze(true);
            bool opSuccess = true;

            oForm.Items.Item("DeclNum").Specific.Value = "";
            int seqNum = 0;

            Dictionary<string, string> rsSettings = CompanyDetails.getRSSettings( out errorText);
            if (errorText != null)
            {
                oForm.Update();
                oForm.Freeze(false);
                Program.uiApp.StatusBar.SetSystemMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return;
            }

            string su = rsSettings["SU"];
            string sp = rsSettings["SP"];

            TaxInvoice oTaxInvoice = new TaxInvoice(rsSettings["ProtocolType"]);

            bool chek_service_user = oTaxInvoice.check_usr( su, sp, out errorText);
            if (chek_service_user == false)
            {
                oForm.Update();
                oForm.Freeze(false);
                errorText = BDOSResources.getTranslate("ServiceUserPasswordNotCorrect");
                Program.uiApp.StatusBar.SetSystemMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return;
            }

            int DeclYear = Convert.ToInt32(oForm.Items.Item("DeclYear").Specific.Value) * -1 + DateTime.Now.Year;
            int DeclMonth = Convert.ToInt32(oForm.Items.Item("DeclMonth").Specific.Value) + 1;
            string DeclDateString = new DateTime(DeclYear, DeclMonth, 1).ToString("yyyyMM");
            DateTime DeclDate = new DateTime(DeclYear, DeclMonth, 1);

            //დეკლარაციის ნომრების მიღება
            DataTable TaxDeclTable = oTaxInvoice.get_seq_nums(DeclDateString, out errorText);

            for (int i = 0; i < TaxDeclTable.Rows.Count; i++)
            {
                DataRow TaxDeclRow = TaxDeclTable.Rows[i];
                seqNum = Convert.ToInt32(TaxDeclRow.ItemArray[0]);
                oForm.Items.Item("DeclNum").Specific.Value = seqNum.ToString();
            }

            if (seqNum == 0)
            {
                oForm.Update();
                oForm.Freeze(false);

                errorText = BDOSResources.getTranslate("CantReceiveDeclarationDate") + " " + DeclDate;

                Program.uiApp.StatusBar.SetSystemMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return;
            }
            //oForm.Freeze(true);
            //დეკლარაციაში დამატება
            string statusRS = null;
            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("TxTable").Specific));
            int rowCount = oMatrix.RowCount;
            for (int row = 1; row <= rowCount; row++)
            {
                SAPbouiCOM.CheckBox Edtfield = oMatrix.Columns.Item("TxChkBx").Cells.Item(row).Specific;
                bool checkedLine = (Edtfield.Checked);

                if (checkedLine)
                {
                    int docEntry = Convert.ToInt32(oMatrix.Columns.Item("Document").Cells.Item(row).Specific.Value);
                    string statusDoc = oMatrix.Columns.Item("StatusDoc").Cells.Item(row).Specific.Value;
                    string declNum = oMatrix.Columns.Item("DeclNum").Cells.Item(row).Specific.Value;
                    //oMatrix.Columns.Item("TxChkBx").Cells.Item(row).Specific.Checked = false;

                    BDO_TaxInvoiceReceived.operationRS( oTaxInvoice, "addToTheDeclaration", Convert.ToInt32(docEntry), seqNum, DeclDate, null, out statusRS, out errorText);
                    if (errorText != null)
                    {
                        Program.uiApp.StatusBar.SetSystemMessage(errorText + "! " + BDOSResources.getTranslate("TableRow") + " : " + row.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        opSuccess = false;
                        continue;
                    }
                    else
                    {
                        Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("Operation") + " " + BDOSResources.getTranslate("RSAddDeclaration") + " " + BDOSResources.getTranslate("DoneSuccessfully") + " " + BDOSResources.getTranslate("TableRow") + " : " + row.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                        updateMatrixRowTaxInvoiceReceived( oMatrix, row, docEntry);
                    }
                }
            }
            if (opSuccess == false)
            {
                Program.uiApp.MessageBox(BDOSResources.getTranslate("OperationForSomeDocumentFinishedWithErrorSeeLog"));
            }
            oForm.Update();
            oForm.Freeze(false);
        }

        public static void receiveVATTaxInvoiceReceived(  SAPbouiCOM.Form oForm, out string errorText)
        {
            bool opSuccess = true;
            errorText = null;

            int declYear = Convert.ToInt32(oForm.Items.Item("DeclYear").Specific.Value) * -1 + DateTime.Now.Year;
            int declMonth = Convert.ToInt32(oForm.Items.Item("DeclMonth").Specific.Value) + 1;
            DateTime declDate = new DateTime(declYear, declMonth, 1);

            //ჩათვლა
            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("TxTable").Specific));
            int rowCount = oMatrix.RowCount;
            for (int row = 1; row <= rowCount; row++)
            {
                SAPbouiCOM.CheckBox Edtfield = oMatrix.Columns.Item("TxChkBx").Cells.Item(row).Specific;
                bool checkedLine = (Edtfield.Checked);

                if (checkedLine)
                {
                    int docEntry = Convert.ToInt32(oMatrix.Columns.Item("Document").Cells.Item(row).Specific.Value);
                    BDO_TaxInvoiceReceived.receiveVAT( docEntry, declDate, out errorText);

                    if (errorText != null)
                    {
                        Program.uiApp.StatusBar.SetSystemMessage(errorText + "! " + BDOSResources.getTranslate("TableRow") + " : " + row.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        opSuccess = false;
                        continue;
                    }
                    else
                    {
                        Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("Operation") + " " + BDOSResources.getTranslate("ReceiveVat") + " " + BDOSResources.getTranslate("DoneSuccessfully") + " " + BDOSResources.getTranslate("TableRow") + " : " + row.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                    }
                }
            }

            if (opSuccess == false)
            {
                Program.uiApp.MessageBox(BDOSResources.getTranslate("OperationForSomeDocumentFinishedWithErrorSeeLog"));
            }
        }

        public static void fillFromBaseTaxInvoiceReceived(  SAPbouiCOM.Form oForm, bool download, out string errorText)
        {
            errorText = null;
            oForm.Freeze(true);

            SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("TxTable");
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            DateTime OperationPeriodStart = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            DateTime OperationPeriodEnd = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month));
            OperationPeriodEnd = OperationPeriodEnd.AddSeconds(24 * 3600 - 1);

            string startDateStr = oForm.DataSources.UserDataSources.Item("StartDate").ValueEx;
            string endDateStr = oForm.DataSources.UserDataSources.Item("EndDate").ValueEx;

            string startDateOpStr = oForm.DataSources.UserDataSources.Item("StartDatOp").ValueEx;
            string endDateOpStr = oForm.DataSources.UserDataSources.Item("EndDateOp").ValueEx;

            string CardCode = oForm.DataSources.UserDataSources.Item("CCode").ValueEx;
            string Attach = oForm.DataSources.UserDataSources.Item("Attach").ValueEx;

            string query = "SELECT \"@BDO_TAXR\".*, " +
                           "\"CardName\", " +
                           "\"LicTradNum\", " +
                           "\"U_BDO_NotInv\" " +
                           "FROM \"@BDO_TAXR\" " +
                           "LEFT JOIN \"OCRD\" " +
                           "ON \"@BDO_TAXR\".\"U_cardCode\" = \"OCRD\".\"CardCode\" " +
                           "WHERE \"Canceled\" = 'N' ";

            //ფილტრი თარიღის მიხედვით (გადასაცემია პარამეტრი)
            if (startDateStr != "")
            {
                query = query + " AND \"U_recvDate\">='" + startDateStr + "'";
            }

            if (endDateStr != "")
            {
                query = query + " AND \"U_recvDate\"<='" + endDateStr + "' ";
            }

            if (startDateOpStr != "")
            {
                query = query + " AND \"U_opDate\">='" + startDateOpStr + "' ";
            }

            if (endDateOpStr != "")
            {
                query = query + " AND \"U_opDate\"<='" + endDateOpStr + "' ";
            }

            if (CardCode != "")
            {
                query = query + " AND\"@BDO_TAXR\".\"U_cardCode\"='" + CardCode + "' ";
            }

            if (Attach != "" && Convert.ToInt32(Attach) > 0)
            {
                Attach = (Convert.ToInt32(Attach) - 1).ToString();
                query = query + " AND\"@BDO_TAXR\".\"U_LinkStatus\"='" + Attach + "' ";
            }

            query = query + @" ORDER BY ""U_recvDate"", ""U_invID"" ";

            int rowIndex = 0;
            string declNum = null;
            string docEntry = "";
            string corrDocEntry = null;
            string statusRS = "";
            string number = null;
            oRecordSet.DoQuery(query);

            oDataTable.Rows.Clear();

            Dictionary<string, string> rsSettings = CompanyDetails.getRSSettings( out errorText);
            if (errorText != null)
            {
                oForm.Update();
                oForm.Freeze(false);
                Program.uiApp.StatusBar.SetSystemMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return;
            }

            string su = rsSettings["SU"];
            string sp = rsSettings["SP"];

            TaxInvoice oTaxInvoice = new TaxInvoice(rsSettings["ProtocolType"]);

            bool chek_service_user = oTaxInvoice.check_usr( su, sp, out errorText);
            if (chek_service_user == false)
            {
                oForm.Update();
                oForm.Freeze(false);
                errorText = BDOSResources.getTranslate("ServiceUserPasswordNotCorrect");
                Program.uiApp.StatusBar.SetSystemMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return;
            }

            while (!oRecordSet.EoF)
            {
                declNum = oRecordSet.Fields.Item("U_declNumber").Value;
                docEntry = oRecordSet.Fields.Item("DocEntry").Value.ToString();
                corrDocEntry = oRecordSet.Fields.Item("U_corrDTxt").Value;
                number = oRecordSet.Fields.Item("U_number").Value;

                if (download == false & String.IsNullOrEmpty(number) == false)
                {
                    BDO_TaxInvoiceReceived.operationRS( oTaxInvoice, "checkSync", Convert.ToInt32(docEntry), -1, new DateTime(), null, out statusRS, out errorText);
                    statusRS = BDO_TaxInvoiceReceived.getStatusValueByStatusNumber(statusRS);
                    //if (BDO_TaxInvoiceReceived.checkSync( null, docEntry, out statusRS, out errorText) == false)
                    //{
                    //    if (errorText == null)
                    //    {
                    //        errorText = BDOSResources.getTranslate("SynchronisationViolatedUpdateStatus");
                    //    }
                    //}
                }
                statusRS = download == false ? statusRS : oRecordSet.Fields.Item("U_status").Value;

                string opDate = oRecordSet.Fields.Item("U_opDate").Value.ToString("yyyyMMdd") == "18991230" ? "" : oRecordSet.Fields.Item("U_opDate").Value.ToString("yyyyMMdd");
                string recvDate = oRecordSet.Fields.Item("U_recvDate").Value.ToString("yyyyMMdd") == "18991230" ? "" : oRecordSet.Fields.Item("U_recvDate").Value.ToString("yyyyMMdd");
                string confDate = oRecordSet.Fields.Item("U_confDate").Value.ToString("yyyyMMdd") == "18991230" ? "" : oRecordSet.Fields.Item("U_confDate").Value.ToString("yyyyMMdd");

                oDataTable.Rows.Add();
                oDataTable.SetValue(0, rowIndex, rowIndex + 1);//LineNum             
                //oDataTable.SetValue(1, rowIndex, "");//SrvStatus
                oDataTable.SetValue(2, rowIndex, statusRS);//Status
                oDataTable.SetValue(3, rowIndex, oRecordSet.Fields.Item("U_status").Value);//StatusDoc
                oDataTable.SetValue(4, rowIndex, (String.IsNullOrEmpty(declNum) == true) ? BDOSResources.getTranslate("WithoutDeclaration") : BDOSResources.getTranslate("WithDeclaration"));//DeclStatus
                oDataTable.SetValue(5, rowIndex, oRecordSet.Fields.Item("U_declNumber").Value);//DeclNum
                oDataTable.SetValue(6, rowIndex, docEntry);//Document
                oDataTable.SetValue(7, rowIndex, corrDocEntry);//corrDoc
                oDataTable.SetValue(8, rowIndex, oRecordSet.Fields.Item("U_series").Value);//TxSerie
                oDataTable.SetValue(9, rowIndex, number);//TxNum
                oDataTable.SetValue(10, rowIndex, oRecordSet.Fields.Item("U_invID").Value);//TxID
                oDataTable.SetValue(11, rowIndex, oRecordSet.Fields.Item("U_cardCode").Value);//CardCode
                oDataTable.SetValue(12, rowIndex, opDate);//OpDate
                oDataTable.SetValue(13, rowIndex, recvDate);//RegDate
                oDataTable.SetValue(14, rowIndex, confDate);//ConfDate
                oDataTable.SetValue(15, rowIndex, oRecordSet.Fields.Item("U_amount").Value);//Sum
                oDataTable.SetValue(16, rowIndex, oRecordSet.Fields.Item("U_amountTX").Value);//VatSum
                oDataTable.SetValue(17, rowIndex, oRecordSet.Fields.Item("U_comment").Value);//Comment
                oDataTable.SetValue(18, rowIndex, "N");//TxChkBx
                oDataTable.SetValue(19, rowIndex, oRecordSet.Fields.Item("LicTradNum").Value);//Vatno              
                oDataTable.SetValue(20, rowIndex, oRecordSet.Fields.Item("U_BDO_NotInv").Value);//Vatno              
                oDataTable.SetValue(21, rowIndex, oRecordSet.Fields.Item("CardName").Value);//Vatno              

                oRecordSet.MoveNext();
                rowIndex++;
            }

            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("TxTable").Specific;
            //oForm.Freeze(true);
            oMatrix.Clear();
            oMatrix.LoadFromDataSource();

            oDataTable = oForm.DataSources.DataTables.Item("TxTable");

            for (int i = 0; i < oMatrix.RowCount; i++)
            {
                if (oDataTable.GetValue("Status", i) != oDataTable.GetValue("StatusDoc", i))
                {
                    oMatrix.CommonSetting.SetCellBackColor(i + 1, 2, FormsB1.getLongIntRGB(255, 48, 48));
                    oMatrix.CommonSetting.SetCellBackColor(i + 1, 3, FormsB1.getLongIntRGB(255, 48, 48));
                }
                else
                {
                    oMatrix.CommonSetting.SetCellBackColor(i + 1, 2, FormsB1.getLongIntRGB(231, 231, 231));
                    oMatrix.CommonSetting.SetCellBackColor(i + 1, 3, FormsB1.getLongIntRGB(231, 231, 231));
                }
            }

            oMatrix.AutoResizeColumns();
            oForm.Update();
            oForm.Freeze(false);
        }

        //----------------------------->Tax Invoice Received<-----------------------------

        public static void matrixColumnSetCfl(SAPbouiCOM.Form oForm, SAPbouiCOM.ItemEvent pVal, out string errorText)
        {
            errorText = null;

            try
            {
                if (pVal.ItemUID == "TxTable2")
                {
                    if (pVal.ColUID == "InvEntry")
                    {
                        if ((pVal.EventType == SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED & pVal.BeforeAction == true) || (pVal.EventType == SAPbouiCOM.BoEventTypes.et_GOT_FOCUS & pVal.BeforeAction == false))
                        {
                            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("TxTable2").Specific));

                            SAPbouiCOM.ComboBox oComboBox = oMatrix.Columns.Item("DocType").Cells.Item(pVal.Row).Specific;
                            SAPbouiCOM.Column oColumn;

                            if (oComboBox.Value == "ARInvoice") //რეალიზაცია
                            {
                                oColumn = oMatrix.Columns.Item(pVal.ColUID);
                                SAPbouiCOM.LinkedButton oLink = oColumn.ExtendedObject;
                                oLink.LinkedObjectType = "13"; //SAPbouiCOM.BoLinkedObject.lf_Invoice
                            }
                            else if (oComboBox.Value == "ARCreditNote") //კორექტირება
                            {
                                oColumn = oMatrix.Columns.Item(pVal.ColUID);
                                SAPbouiCOM.LinkedButton oLink = oColumn.ExtendedObject;
                                oLink.LinkedObjectType = "14"; //SAPbouiCOM.BoLinkedObject.lf_InvoiceCreditMemo
                            }
                            else if (oComboBox.Value == "ARDownPaymentVAT") //ავანსი
                            {
                                oColumn = oMatrix.Columns.Item(pVal.ColUID);
                                SAPbouiCOM.LinkedButton oLink = oColumn.ExtendedObject;
                                oLink.LinkedObjectType = "UDO_F_BDO_ARDPV_D"; //A/R Down Payment Invoice
                            }
                        }
                    }
                    else
                    {

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

        public static void addMenus( out string errorText)
        {
            errorText = null;

            SAPbouiCOM.MenuItem menuItem;
            SAPbouiCOM.MenuItem fatherMenuItem;
            SAPbouiCOM.MenuCreationParams oCreationPackage;

            try
            {
                oCreationPackage = (SAPbouiCOM.MenuCreationParams)Program.uiApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                fatherMenuItem = Program.uiApp.Menus.Item("1536");
                
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                oCreationPackage.UniqueID = "TAX1536";
                oCreationPackage.String = BDOSResources.getTranslate("Tax");
                oCreationPackage.Position = -1;

                menuItem = fatherMenuItem.SubMenus.AddEx(oCreationPackage);
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }


            try
            {
                fatherMenuItem = Program.uiApp.Menus.Item("TAX1536");
                
                // Add a pop-up menu item
                oCreationPackage = (SAPbouiCOM.MenuCreationParams)Program.uiApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                oCreationPackage.Checked = false;
                oCreationPackage.Enabled = true;
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "BDOSTAXJ";
                oCreationPackage.String = BDOSResources.getTranslate("TaxInvoceJournal");
                oCreationPackage.Position = -1;

                menuItem = fatherMenuItem.SubMenus.AddEx(oCreationPackage);
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
        }

        public static void createForm(  out string errorText)
        {
            errorText = null;

            //ფორმის აუცილებელი თვისებები
            Dictionary<string, object> formProperties = new Dictionary<string, object>();
            formProperties.Add("UniqueID", "BDOSTaxRecvForm");
            formProperties.Add("BorderStyle", SAPbouiCOM.BoFormBorderStyle.fbs_Sizable);
            formProperties.Add("Title", BDOSResources.getTranslate("TaxInvoceJournal"));
            formProperties.Add("Left", 558);
            formProperties.Add("ClientWidth", 750);
            formProperties.Add("Top", 335);
            formProperties.Add("ClientHeight", 400);

            SAPbouiCOM.Form oForm;
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

                    //bool multiSelection = false;
                    //string objectType = "13"; //SAPbouiCOM.BoLinkedObject.lf_Invoice
                    //string uniqueID_lf_InvoiceCFL = "Invoice_CFL";
                    //FormsB1.addChooseFromList( oForm, multiSelection, objectType, uniqueID_lf_InvoiceCFL);
                    //objectType = "14"; //SAPbouiCOM.BoLinkedObject.lf_InvoiceCreditMemo 
                    //string uniqueID_lf_InvoiceCreditMemoCFL = "InvoiceCreditMemo_CFL";
                    //FormsB1.addChooseFromList( oForm, multiSelection, objectType, uniqueID_lf_InvoiceCreditMemoCFL);
                    //objectType = "UDO_F_BDO_ARDPV_D"; //A/R Down Payment Invoice
                    //string uniqueID_lf_DownPaymentInvoiceCFL = "DownPaymentInvoice_CFL";
                    //FormsB1.addChooseFromList( oForm, multiSelection, objectType, uniqueID_lf_DownPaymentInvoiceCFL);

                    //ფორმის ელემენტების თვისებები
                    Dictionary<string, object> formItems = null;

                    string itemName = "";
                    int left = 6;
                    int Top = 5;
                    int leftSC = 400;
                    List<string> listValidValues;

                    Top = Top + 20;
                    left = 6;

                    //რიგი 1 წარწერები შერჩევა და დეკლარაციაში დამატება ჩანს ორივე ჩანართზე
                    formItems = new Dictionary<string, object>();
                    itemName = "FilterSt";
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 130);
                    formItems.Add("Top", Top);
                    formItems.Add("Caption", BDOSResources.getTranslate("Filter"));
                    formItems.Add("UID", itemName);
                    formItems.Add("TextStyle", 4);
                    formItems.Add("FontSize", 10);
                    formItems.Add("FromPane", 1);
                    formItems.Add("ToPane", 2);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "DeclSt";
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", leftSC);
                    formItems.Add("Width", 170);
                    formItems.Add("Top", Top);
                    formItems.Add("Caption", BDOSResources.getTranslate("ReceiveAddDeclaration"));
                    formItems.Add("UID", itemName);
                    formItems.Add("TextStyle", 4);
                    formItems.Add("FontSize", 10);
                    formItems.Add("FromPane", 1);
                    formItems.Add("ToPane", 2);
                    //

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    DateTime OperationPeriodStart;
                    DateTime OperationPeriodEnd;

                    OperationPeriodStart = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
                    OperationPeriodEnd = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month));
                    OperationPeriodEnd = OperationPeriodEnd.AddSeconds(24 * 3600 - 1);

                    //რიგი 2
                    Top = Top + 20;
                    left = 6;

                    //თარიღები
                    formItems = new Dictionary<string, object>();
                    itemName = "DateFromOp";
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 130);
                    formItems.Add("Top", Top);
                    formItems.Add("Caption", BDOSResources.getTranslate("OperationPeriod"));
                    formItems.Add("UID", itemName);
                    formItems.Add("FromPane", 1);
                    formItems.Add("ToPane", 1);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    //გაცემულები Pane "2"
                    formItems = new Dictionary<string, object>();
                    itemName = "DateFrmOp2";
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 130);
                    formItems.Add("Top", Top);
                    formItems.Add("Caption", BDOSResources.getTranslate("OperationPeriod"));
                    formItems.Add("UID", itemName);
                    formItems.Add("FromPane", 2);
                    formItems.Add("ToPane", 2);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    //გაცემულები Pane "2"

                    left = left + 130 + 10;

                    string startOfMonthStr = OperationPeriodStart.ToString("yyyyMMdd");
                    formItems = new Dictionary<string, object>();
                    itemName = "StartDatOp";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_DATE);
                    formItems.Add("Length", 1);
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
                    formItems.Add("ValueEx", startOfMonthStr);
                    formItems.Add("FromPane", 1);
                    formItems.Add("ToPane", 1);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    //გაცემულები Pane "2"
                    formItems = new Dictionary<string, object>();
                    itemName = "StrtDatOp2";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_DATE);
                    formItems.Add("Length", 1);
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
                    formItems.Add("ValueEx", startOfMonthStr);
                    formItems.Add("FromPane", 2);
                    formItems.Add("ToPane", 2);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }
                    //გაცემულები Pane "2"

                    left = left + 100 + 10;

                    string endOfMonthStr = OperationPeriodEnd.ToString("yyyyMMdd");
                    formItems = new Dictionary<string, object>();
                    itemName = "EndDateOp";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_DATE);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("Length", 1);
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
                    formItems.Add("ValueEx", endOfMonthStr);
                    formItems.Add("FromPane", 1);
                    formItems.Add("ToPane", 1);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    //გაცემულები Pane "2"
                    endOfMonthStr = OperationPeriodEnd.ToString("yyyyMMdd");
                    formItems = new Dictionary<string, object>();
                    itemName = "EndDateOp2";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_DATE);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("Length", 1);
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
                    formItems.Add("ValueEx", endOfMonthStr);
                    formItems.Add("FromPane", 2);
                    formItems.Add("ToPane", 2);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }


                    //გაცემულები Pane "2"


                    formItems = new Dictionary<string, object>(); //დეკლარაციის წელი

                    listValidValues = new List<string>();
                    int CurrYear = DateTime.Now.Year;
                    int FirstYear = CurrYear - 5;
                    for (int yrCnt = CurrYear; yrCnt >= FirstYear; yrCnt--)
                    {
                        listValidValues.Add(yrCnt.ToString());
                    }

                    itemName = "DeclYear";
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                    formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("Left", leftSC);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("ValidValues", listValidValues);
                    formItems.Add("FromPane", 1);
                    formItems.Add("ToPane", 1);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    SAPbouiCOM.ComboBox oCombobox = (SAPbouiCOM.ComboBox)oForm.Items.Item("DeclYear").Specific;
                    oCombobox.Select("0");

                    formItems = new Dictionary<string, object>(); //დეკლარაციის ნომერი
                    itemName = "DeclNumSt";
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", leftSC + 100 + 10);
                    formItems.Add("Width", 10);
                    formItems.Add("Top", Top);
                    formItems.Add("Caption", "#");
                    formItems.Add("UID", itemName);
                    formItems.Add("FromPane", 1);
                    formItems.Add("ToPane", 1);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "DeclNum";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("Length", 20);
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Left", leftSC + 100 + 10 + 5 + 10);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("FromPane", 1);
                    formItems.Add("ToPane", 1);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    //დეკლარაცია - გაცემულები Pane "2"
                    formItems = new Dictionary<string, object>(); //დეკლარაციის წელი

                    listValidValues = new List<string>();
                    for (int yrCnt = CurrYear; yrCnt >= FirstYear; yrCnt--)
                    {
                        listValidValues.Add(yrCnt.ToString());
                    }

                    itemName = "DeclYear2";
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                    formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("Left", leftSC);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("ValidValues", listValidValues);
                    formItems.Add("FromPane", 2);
                    formItems.Add("ToPane", 2);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    oCombobox = (SAPbouiCOM.ComboBox)oForm.Items.Item("DeclYear2").Specific;
                    oCombobox.Select("0");

                    formItems = new Dictionary<string, object>(); //დეკლარაციის ნომერი
                    itemName = "DeclNumSt2";
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", leftSC + 100 + 10);
                    formItems.Add("Width", 10);
                    formItems.Add("Top", Top);
                    formItems.Add("Caption", "#");
                    formItems.Add("UID", itemName);
                    formItems.Add("FromPane", 2);
                    formItems.Add("ToPane", 2);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "DeclNum2";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("Length", 20);
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Left", leftSC + 100 + 10 + 5 + 10);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("FromPane", 2);
                    formItems.Add("ToPane", 2);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }
                    //დეკლარაცია - გაცემულები - Pane "2"

                    //რიგი 3
                    Top = Top + 20;
                    left = 6;

                    //თარიღები
                    formItems = new Dictionary<string, object>();
                    itemName = "DateFrom";
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 130);
                    formItems.Add("Top", Top);
                    formItems.Add("Caption", BDOSResources.getTranslate("RegistrationPeriod"));
                    formItems.Add("UID", itemName);
                    formItems.Add("FromPane", 1);
                    formItems.Add("ToPane", 1);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    //გაცემულები Pane "2"
                    formItems = new Dictionary<string, object>();
                    itemName = "DateFrom2";
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 130);
                    formItems.Add("Top", Top);
                    formItems.Add("Caption", BDOSResources.getTranslate("TaxInvoicePeriod"));
                    formItems.Add("UID", itemName);
                    formItems.Add("FromPane", 2);
                    formItems.Add("ToPane", 2);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }
                    //გაცემულები Pane "2"

                    left = left + 130 + 10;

                    startOfMonthStr = OperationPeriodStart.ToString("yyyyMMdd");
                    formItems = new Dictionary<string, object>();
                    itemName = "StartDate";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_DATE);
                    formItems.Add("Length", 1);
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
                    formItems.Add("ValueEx", DateTime.Today.ToString("yyyyMMdd"));
                    formItems.Add("FromPane", 1);
                    formItems.Add("ToPane", 1);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    //გაცემულები Pane "2"           
                    formItems = new Dictionary<string, object>();
                    itemName = "StartDate2";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_DATE);
                    formItems.Add("Length", 1);
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
                    // formItems.Add("ValueEx", startOfMonthStr);
                    formItems.Add("FromPane", 2);
                    formItems.Add("ToPane", 2);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }
                    //გაცემულები Pane "2"

                    left = left + 100 + 10;


                    endOfMonthStr = OperationPeriodEnd.ToString("yyyyMMdd");
                    formItems = new Dictionary<string, object>();
                    itemName = "EndDate";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_DATE);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("Length", 1);
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
                    formItems.Add("ValueEx", DateTime.Today.ToString("yyyyMMdd"));
                    formItems.Add("FromPane", 1);
                    formItems.Add("ToPane", 1);


                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    //გაცემულები Pane "2"
                    formItems = new Dictionary<string, object>();
                    itemName = "EndDate2";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_DATE);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("Length", 1);
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
                    //formItems.Add("ValueEx", endOfMonthStr);
                    formItems.Add("FromPane", 2);
                    formItems.Add("ToPane", 2);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    //გაცემულები Pane "2"

                    formItems = new Dictionary<string, object>(); //დეკლარაციის თვე

                    listValidValues = new List<string>();
                    int CurrMonth = DateTime.Now.Month - 1;
                    listValidValues.Add(BDOSResources.getTranslate("January"));
                    listValidValues.Add(BDOSResources.getTranslate("February"));
                    listValidValues.Add(BDOSResources.getTranslate("March"));
                    listValidValues.Add(BDOSResources.getTranslate("April"));
                    listValidValues.Add(BDOSResources.getTranslate("May"));
                    listValidValues.Add(BDOSResources.getTranslate("June"));
                    listValidValues.Add(BDOSResources.getTranslate("July"));
                    listValidValues.Add(BDOSResources.getTranslate("August"));
                    listValidValues.Add(BDOSResources.getTranslate("Septempet"));
                    listValidValues.Add(BDOSResources.getTranslate("Octomber"));
                    listValidValues.Add(BDOSResources.getTranslate("November"));
                    listValidValues.Add(BDOSResources.getTranslate("December"));

                    itemName = "DeclMonth";
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                    formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("Left", leftSC);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("ValidValues", listValidValues);
                    formItems.Add("FromPane", 1);
                    formItems.Add("ToPane", 1);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    SAPbouiCOM.ComboBox oComboboxM = (SAPbouiCOM.ComboBox)oForm.Items.Item("DeclMonth").Specific;
                    oComboboxM.Select(CurrMonth.ToString());

                    formItems = new Dictionary<string, object>();
                    Dictionary<string, object> fieldskeysMap;
                    fieldskeysMap = new Dictionary<string, object>();
                    listValidValues = new List<string>();
                    listValidValues.Add(BDOSResources.getTranslate("RSAddDeclaration"));
                    listValidValues.Add(BDOSResources.getTranslate("ReceiveVat"));


                    formItems = new Dictionary<string, object>(); //დეკლარაციაში დამატება
                    itemName = "addDecl";
                    formItems.Add("Caption", BDOSResources.getTranslate("Add"));
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON_COMBO);
                    formItems.Add("Left", leftSC + 100 + 10 + 10 + 5);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("ValidValues", listValidValues);
                    formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
                    formItems.Add("AffectsFormMode", false);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("FromPane", 1);
                    formItems.Add("ToPane", 1);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }


                    //დეკლ.დამატება - გამავალი - Pane "2"
                    formItems = new Dictionary<string, object>(); //დეკლარაციის თვე

                    listValidValues = new List<string>();
                    CurrMonth = DateTime.Now.Month - 1;
                    listValidValues.Add(BDOSResources.getTranslate("January"));
                    listValidValues.Add(BDOSResources.getTranslate("February"));
                    listValidValues.Add(BDOSResources.getTranslate("March"));
                    listValidValues.Add(BDOSResources.getTranslate("April"));
                    listValidValues.Add(BDOSResources.getTranslate("May"));
                    listValidValues.Add(BDOSResources.getTranslate("June"));
                    listValidValues.Add(BDOSResources.getTranslate("July"));
                    listValidValues.Add(BDOSResources.getTranslate("August"));
                    listValidValues.Add(BDOSResources.getTranslate("Septempet"));
                    listValidValues.Add(BDOSResources.getTranslate("Octomber"));
                    listValidValues.Add(BDOSResources.getTranslate("November"));
                    listValidValues.Add(BDOSResources.getTranslate("December"));

                    itemName = "DeclMonth2";
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                    formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("Left", leftSC);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("ValidValues", listValidValues);
                    formItems.Add("FromPane", 2);
                    formItems.Add("ToPane", 2);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    oComboboxM = (SAPbouiCOM.ComboBox)oForm.Items.Item("DeclMonth2").Specific;
                    oComboboxM.Select(CurrMonth.ToString());

                    formItems = new Dictionary<string, object>();
                    fieldskeysMap = new Dictionary<string, object>();
                    listValidValues = new List<string>();
                    listValidValues.Add(BDOSResources.getTranslate("RSAddDeclaration"));

                    formItems = new Dictionary<string, object>(); //დეკლარაციაში დამატება
                    itemName = "addDecl2";
                    formItems.Add("Caption", BDOSResources.getTranslate("Add"));
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON_COMBO);
                    formItems.Add("Left", leftSC + 100 + 10 + 10 + 5);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("ValidValues", listValidValues);
                    formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
                    formItems.Add("AffectsFormMode", false);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("FromPane", 2);
                    formItems.Add("ToPane", 2);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }
                    //დეკლ.დამატება - გამავალი - Pane "2"

                    //რიგი 4
                    Top = Top + 20;
                    left = 6;

                    formItems = new Dictionary<string, object>();
                    itemName = "CardCodeSt";
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 130);
                    formItems.Add("Top", Top);
                    formItems.Add("Caption", BDOSResources.getTranslate("BPCardCode"));
                    formItems.Add("UID", itemName);
                    formItems.Add("FromPane", 1);
                    formItems.Add("ToPane", 1);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    //ბიზნეს პარტნიორი წარწერა - გაცემული - Pane 2
                    formItems = new Dictionary<string, object>();
                    itemName = "CCodeSt2";
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 130);
                    formItems.Add("Top", Top);
                    formItems.Add("Caption", BDOSResources.getTranslate("BPCardCode"));
                    formItems.Add("UID", itemName);
                    formItems.Add("FromPane", 2);
                    formItems.Add("ToPane", 2);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }
                    //ბიზნეს პარტნიორი წარწერა  - გაცემული - Pane 2

                    left = left + 130 + 10;

                    bool multiSelection = false;
                    string objectType = "2"; //Warehouse
                    string uniqueID_lf_BusinessPartnerCFL = "BP_CFL";
                    FormsB1.addChooseFromList( oForm, multiSelection, objectType, uniqueID_lf_BusinessPartnerCFL);

                    //პირობის დადება ბიზნესპარტნიორის არჩევის სიაზე
                    SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(uniqueID_lf_BusinessPartnerCFL);
                    SAPbouiCOM.Conditions oCons = oCFL.GetConditions();
                    SAPbouiCOM.Condition oCon = oCons.Add();
                    oCon.Alias = "CardType";
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCon.CondVal = "S"; //მომწოდებელი
                    oCFL.SetConditions(oCons);

                    formItems = new Dictionary<string, object>();
                    itemName = "CCode";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
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
                    formItems.Add("ChooseFromListUID", uniqueID_lf_BusinessPartnerCFL);
                    formItems.Add("ChooseFromListAlias", "CardCode");
                    formItems.Add("FromPane", 1);
                    formItems.Add("ToPane", 1);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    //ბიზნეს პარტნიორი ველები - გაცემული - Pane 2
                    string uniqueID_lf_BusinessPartnerCFL2 = "BP_CFL2";
                    FormsB1.addChooseFromList( oForm, multiSelection, objectType, uniqueID_lf_BusinessPartnerCFL2);

                    //პირობის დადება ბიზნესპარტნიორის არჩევის სიაზე
                    SAPbouiCOM.ChooseFromList oCFL2 = oForm.ChooseFromLists.Item(uniqueID_lf_BusinessPartnerCFL2);
                    SAPbouiCOM.Conditions oCons2 = oCFL2.GetConditions();
                    SAPbouiCOM.Condition oCon2 = oCons2.Add();
                    oCon2.Alias = "CardType";
                    oCon2.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCon2.CondVal = "C"; //კლიენტი
                    oCFL2.SetConditions(oCons2);

                    formItems = new Dictionary<string, object>();
                    itemName = "CCode2";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
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
                    formItems.Add("ChooseFromListUID", uniqueID_lf_BusinessPartnerCFL2);
                    formItems.Add("ChooseFromListAlias", "CardCode");
                    formItems.Add("FromPane", 2);
                    formItems.Add("ToPane", 2);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }
                    //ბიზნეს პარტნიორი ველები - გაცემული - Pane 2

                    formItems = new Dictionary<string, object>();
                    itemName = "CCode_LB"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    formItems.Add("Left", left - 20);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("LinkTo", "CCode");
                    formItems.Add("LinkedObjectType", objectType);
                    formItems.Add("FromPane", 1);
                    formItems.Add("ToPane", 1);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    //ბიზნეს პარტნიორი ველები - გაცემული - Pane 2
                    formItems = new Dictionary<string, object>();
                    itemName = "CCode_LB2"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    formItems.Add("Left", left - 20);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("LinkTo", "CCode2");
                    formItems.Add("LinkedObjectType", objectType);
                    formItems.Add("FromPane", 2);
                    formItems.Add("ToPane", 2);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }
                    //ბიზნეს პარტნიორი ველები - გაცემული - Pane 2

                    //რიგი 5
                    Top = Top + 20;
                    left = 6;

                    formItems = new Dictionary<string, object>();
                    itemName = "AttachST";
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 130);
                    formItems.Add("Top", Top);
                    formItems.Add("Caption", BDOSResources.getTranslate("LinkToDocument"));
                    formItems.Add("UID", itemName);
                    formItems.Add("FromPane", 1);
                    formItems.Add("ToPane", 1);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    //გაცემული - Pane 2
                    formItems = new Dictionary<string, object>();
                    itemName = "AttachST2";
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 130);
                    formItems.Add("Top", Top);
                    formItems.Add("Caption", BDOSResources.getTranslate("LinkToDocument"));
                    formItems.Add("UID", itemName);
                    formItems.Add("FromPane", 2);
                    formItems.Add("ToPane", 2);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    /////////
                    formItems = new Dictionary<string, object>();
                    itemName = "NeedTxST2";
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left + 130 + 10 + 100 + 10);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", Top);
                    formItems.Add("Caption", BDOSResources.getTranslate("NeedTax"));
                    formItems.Add("UID", itemName);
                    formItems.Add("FromPane", 2);
                    formItems.Add("ToPane", 2);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }
                    //გაცემული - Pane 2

                    left = left + 130 + 10;

                    List<string> ValidValues = new List<string>();
                    ValidValues.Add(BDOSResources.getTranslate("WithoutFilter"));

                    SAPbobsCOM.UserTable oUserTable = null;
                    oUserTable = Program.oCompany.UserTables.Item("BDO_TAXR");
                    SAPbobsCOM.ValidValues StatusValidValues = oUserTable.UserFields.Fields.Item("U_LinkStatus").ValidValues;

                    for (int i = 0; i < StatusValidValues.Count; i++)
                    {
                        ValidValues.Add(StatusValidValues.Item(i).Description);
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "Attach";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("Length", 30);
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("ValidValues", ValidValues);
                    formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("FromPane", 1);
                    formItems.Add("ToPane", 1);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    //გაცემული - Pane 2
                    ValidValues = new List<string>();
                    ValidValues.Add(BDOSResources.getTranslate("WithoutFilter"));
                    ValidValues.Add(BDOSResources.getTranslate("TaxInvoiceAlreadyCreated"));
                    ValidValues.Add(BDOSResources.getTranslate("NoTaxInvoiceSaved"));

                    formItems = new Dictionary<string, object>();
                    itemName = "Attach2";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("Length", 30);
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("ValidValues", ValidValues);
                    formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("FromPane", 2);
                    formItems.Add("ToPane", 2);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }
                    /////////////////////
                    ValidValues = new List<string>();
                    ValidValues.Add(BDOSResources.getTranslate("All"));
                    ValidValues.Add(BDOSResources.getTranslate("OnlyNeedTax"));
                    

                    formItems = new Dictionary<string, object>();
                    itemName = "needTax";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("Length", 30);
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Left", left + 100 + 10 + 100 + 10);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("ValidValues", ValidValues);
                    formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("FromPane", 2);
                    formItems.Add("ToPane", 2);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }
                    //გაცემული - Pane 2


                    //რიგი 6
                    Top = Top + 25;
                    left = 6;

                    formItems = new Dictionary<string, object>();
                    itemName = "TxCheck";
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 19);
                    formItems.Add("Top", Top + 1);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("Image", "HANA_CHECKBOX_CH");
                    formItems.Add("FromPane", 1);
                    formItems.Add("ToPane", 1);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    //გაცემული - Pane 2
                    formItems = new Dictionary<string, object>();
                    itemName = "TxCheck2";
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 19);
                    formItems.Add("Top", Top + 1);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("Image", "HANA_CHECKBOX_CH");
                    formItems.Add("FromPane", 2);
                    formItems.Add("ToPane", 2);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }
                    //გაცემული - Pane 2


                    left = left + 20 + 1;

                    formItems = new Dictionary<string, object>();
                    itemName = "TxUncheck";
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 19);
                    formItems.Add("Top", Top + 1);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("Image", "HANA_CHECKBOX_UH");
                    formItems.Add("FromPane", 1);
                    formItems.Add("ToPane", 1);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    //გაცემული - Pane 2
                    formItems = new Dictionary<string, object>();
                    itemName = "TxUncheck2";
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 19);
                    formItems.Add("Top", Top + 1);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("Image", "HANA_CHECKBOX_UH");
                    formItems.Add("FromPane", 2);
                    formItems.Add("ToPane", 2);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }
                    //გაცემული - Pane 2

                    left = left + 20 + 5;

                    formItems = new Dictionary<string, object>();
                    itemName = "fillFrmBs";
                    formItems.Add("Caption", BDOSResources.getTranslate("Fill"));
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("FromPane", 1);
                    formItems.Add("ToPane", 1);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    //გაცემულები - Pane 2
                    formItems = new Dictionary<string, object>();
                    itemName = "fillFrmBs2";
                    formItems.Add("Caption", BDOSResources.getTranslate("Fill"));
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("FromPane", 2);
                    formItems.Add("ToPane", 2);
                    //გაცემულები - Pane 2

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left = left + 100 + 2;

                    formItems = new Dictionary<string, object>();
                    itemName = "dwnldTax";
                    formItems.Add("Caption", BDOSResources.getTranslate("RSDownload"));
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("FromPane", 1);
                    formItems.Add("ToPane", 1);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    //გაცემული - ოპერაციები - Pane 2
                    formItems = new Dictionary<string, object>();

                    fieldskeysMap = new Dictionary<string, object>();
                    listValidValues = new List<string>();
                    listValidValues.Add(BDOSResources.getTranslate("RSSave"));
                    listValidValues.Add(BDOSResources.getTranslate("RSSend"));
                    listValidValues.Add(BDOSResources.getTranslate("RSDelete"));
                    listValidValues.Add(BDOSResources.getTranslate("RSCancel"));
                    listValidValues.Add(BDOSResources.getTranslate("RSUpdateStatus"));
                    listValidValues.Add(BDOSResources.getTranslate("RSSaveUnitedTaxInvoice"));

                    itemName = "TxOperRS2";
                    formItems.Add("Caption", BDOSResources.getTranslate("Operations"));
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON_COMBO);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("ValidValues", listValidValues);
                    formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
                    formItems.Add("AffectsFormMode", false);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("FromPane", 2);
                    formItems.Add("ToPane", 2);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    //გაცემული - ოპერაციები - Pane 2

                    left = left + 100 + 2;

                    formItems = new Dictionary<string, object>();
                    itemName = "updtTax";
                    formItems.Add("Caption", BDOSResources.getTranslate("RSUpdate"));
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("FromPane", 1);
                    formItems.Add("ToPane", 1);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left = left + 100 + 2;

                    formItems = new Dictionary<string, object>();

                    fieldskeysMap = new Dictionary<string, object>();
                    listValidValues = new List<string>();
                    listValidValues.Add(BDOSResources.getTranslate("RSConfirm"));
                    listValidValues.Add(BDOSResources.getTranslate("RSUpdateStatus"));
                    listValidValues.Add(BDOSResources.getTranslate("RSDeny"));

                    itemName = "TxOperRS";
                    formItems.Add("Caption", BDOSResources.getTranslate("Operations"));
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON_COMBO);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("ValidValues", listValidValues);
                    formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
                    formItems.Add("AffectsFormMode", false);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("FromPane", 1);
                    formItems.Add("ToPane", 1);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    //რიგი
                    Top = Top + 25;
                    left = 6;

                    itemName = "TxTable";
                    formItems = new Dictionary<string, object>();
                    formItems.Add("isDataSource", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 750);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 200);
                    formItems.Add("UID", itemName);
                    formItems.Add("FromPane", 1);
                    formItems.Add("ToPane", 1);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    //გაცემული Pane 2
                    itemName = "TxTable2";
                    formItems = new Dictionary<string, object>();
                    formItems.Add("isDataSource", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 750);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 200);
                    formItems.Add("UID", itemName);
                    formItems.Add("FromPane", 2);
                    formItems.Add("ToPane", 2);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }
                    //გაცემული Pane 2

                    SAPbouiCOM.LinkedButton oLink;
                    SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("TxTable").Specific));
                    oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single;
                    SAPbouiCOM.Columns oColumns = oMatrix.Columns;
                    SAPbouiCOM.Column oColumn;

                    SAPbouiCOM.DataTable oDataTable;
                    oDataTable = oForm.DataSources.DataTables.Add("TxTable");

                    oDataTable.Columns.Add("LineNum", SAPbouiCOM.BoFieldsType.ft_Text, 50); // 0 - ინდექსი გვჭირდება SetValue-ს პირველ პარამეტრად
                    oDataTable.Columns.Add("SrvStatus", SAPbouiCOM.BoFieldsType.ft_Text, 50); //1
                    oDataTable.Columns.Add("Status", SAPbouiCOM.BoFieldsType.ft_Text, 50); //2
                    oDataTable.Columns.Add("StatusDoc", SAPbouiCOM.BoFieldsType.ft_Text, 50); //3
                    oDataTable.Columns.Add("DeclStatus", SAPbouiCOM.BoFieldsType.ft_Text, 50); //4                
                    oDataTable.Columns.Add("DeclNum", SAPbouiCOM.BoFieldsType.ft_Text, 50); //5
                    oDataTable.Columns.Add("Document", SAPbouiCOM.BoFieldsType.ft_Text, 50); //6
                    oDataTable.Columns.Add("CorrDoc", SAPbouiCOM.BoFieldsType.ft_Text, 50); //7
                    oDataTable.Columns.Add("TxSerie", SAPbouiCOM.BoFieldsType.ft_Text, 50); //8
                    oDataTable.Columns.Add("TxNum", SAPbouiCOM.BoFieldsType.ft_Text, 50); //9
                    oDataTable.Columns.Add("TxID", SAPbouiCOM.BoFieldsType.ft_Text, 50); //10
                    oDataTable.Columns.Add("CardCode", SAPbouiCOM.BoFieldsType.ft_Text, 50); //11
                    oDataTable.Columns.Add("OpDate", SAPbouiCOM.BoFieldsType.ft_Date, 50); //12
                    oDataTable.Columns.Add("RegDate", SAPbouiCOM.BoFieldsType.ft_Date, 50); //13
                    oDataTable.Columns.Add("ConfDate", SAPbouiCOM.BoFieldsType.ft_Date, 50); //14
                    oDataTable.Columns.Add("Sum", SAPbouiCOM.BoFieldsType.ft_Sum, 50); //15
                    oDataTable.Columns.Add("VatSum", SAPbouiCOM.BoFieldsType.ft_Sum, 50); //16
                    oDataTable.Columns.Add("Comment", SAPbouiCOM.BoFieldsType.ft_Text, 50); //17
                    oDataTable.Columns.Add("TxChkBx", SAPbouiCOM.BoFieldsType.ft_Text, 50); //18
                    oDataTable.Columns.Add("VATno", SAPbouiCOM.BoFieldsType.ft_Text, 50); //19
                    oDataTable.Columns.Add("IsVATPayer", SAPbouiCOM.BoFieldsType.ft_Text, 50); //20
                    oDataTable.Columns.Add("CardName", SAPbouiCOM.BoFieldsType.ft_Text, 50); //21

                    oColumn = oColumns.Add("LineNum", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = "#";
                    oColumn.TitleObject.Sortable = true;
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind("TxTable", "LineNum");

                    oColumn = oColumns.Add("TxChkBx", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
                    oColumn.TitleObject.Caption = "";
                    oColumn.TitleObject.Sortable = true;
                    oColumn.Width = 100;
                    oColumn.Editable = true;
                    oColumn.DataBind.Bind("TxTable", "TxChkBx");
                    oColumn.ValOn = "Y";
                    oColumn.ValOff = "N";

                    oColumn = oColumns.Add("Status", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("TaxInvoceStatusRs");
                    oColumn.TitleObject.Sortable = true;
                    oColumn.Width = 200;
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind("TxTable", "Status");
                    oColumn.DisplayDesc = true;
                    oColumn.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;

                    oColumn.ValidValues.Add("empty", "");
                    oColumn.ValidValues.Add("paper", BDOSResources.getTranslate("Paper")); //ქაღალდის
                    oColumn.ValidValues.Add("received", BDOSResources.getTranslate("Received")); //მიღებული
                    oColumn.ValidValues.Add("confirmed", BDOSResources.getTranslate("Confirmed")); //დადასტურებული
                    oColumn.ValidValues.Add("incompleteReceived", BDOSResources.getTranslate("IncompleteReceived")); //არასრულად მიღებული
                    oColumn.ValidValues.Add("denied", BDOSResources.getTranslate("Denied")); //უარყოფილი
                    oColumn.ValidValues.Add("cancellationProcess", BDOSResources.getTranslate("CancellationProcess")); //გაუქმების პროცესში
                    oColumn.ValidValues.Add("canceled", BDOSResources.getTranslate("Canceled")); //გაუქმებული
                    oColumn.ValidValues.Add("correctionReceived", BDOSResources.getTranslate("CorrectionReceived")); //მიღებული კორექტირებული
                    oColumn.ValidValues.Add("correctionDenied", BDOSResources.getTranslate("CorrectionDenied")); //უარყოფილი კორექტირებული
                    oColumn.ValidValues.Add("correctionConfirmed", BDOSResources.getTranslate("CorrectionConfirmed")); //დადასტურებული კორექტირებული
                    oColumn.ValidValues.Add("attachedToTheDeclaration", BDOSResources.getTranslate("AttachedToTheDeclaration")); //დეკლარაციაზე მიბმული
                    oColumn.ValidValues.Add("removed", BDOSResources.getTranslate("Removed")); //წაშლილი
                    oColumn.ValidValues.Add("corrected", BDOSResources.getTranslate("Corrected")); //კორექტირებული
                    oColumn.ValidValues.Add("replaced", BDOSResources.getTranslate("Replaced")); //ჩანაცვლებული

                    oColumn = oColumns.Add("StatusDoc", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("TaxInvoceStatus");
                    oColumn.TitleObject.Sortable = true;
                    oColumn.Width = 200;
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind("TxTable", "StatusDoc");
                    oColumn.DisplayDesc = true;
                    oColumn.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;

                    oColumn.ValidValues.Add("empty", "");
                    oColumn.ValidValues.Add("paper", BDOSResources.getTranslate("Paper")); //ქაღალდის
                    oColumn.ValidValues.Add("received", BDOSResources.getTranslate("Received")); //მიღებული
                    oColumn.ValidValues.Add("confirmed", BDOSResources.getTranslate("Confirmed")); //დადასტურებული
                    oColumn.ValidValues.Add("incompleteReceived", BDOSResources.getTranslate("IncompleteReceived")); //არასრულად მიღებული
                    oColumn.ValidValues.Add("denied", BDOSResources.getTranslate("Denied")); //უარყოფილი
                    oColumn.ValidValues.Add("cancellationProcess", BDOSResources.getTranslate("CancellationProcess")); //გაუქმების პროცესში
                    oColumn.ValidValues.Add("canceled", BDOSResources.getTranslate("Canceled")); //გაუქმებული
                    oColumn.ValidValues.Add("correctionReceived", BDOSResources.getTranslate("CorrectionReceived")); //მიღებული კორექტირებული
                    oColumn.ValidValues.Add("correctionDenied", BDOSResources.getTranslate("CorrectionDenied")); //უარყოფილი კორექტირებული
                    oColumn.ValidValues.Add("correctionConfirmed", BDOSResources.getTranslate("CorrectionConfirmed")); //დადასტურებული კორექტირებული
                    oColumn.ValidValues.Add("attachedToTheDeclaration", BDOSResources.getTranslate("AttachedToTheDeclaration")); //დეკლარაციაზე მიბმული
                    oColumn.ValidValues.Add("removed", BDOSResources.getTranslate("Removed")); //წაშლილი
                    oColumn.ValidValues.Add("corrected", BDOSResources.getTranslate("Corrected")); //კორექტირებული
                    oColumn.ValidValues.Add("replaced", BDOSResources.getTranslate("Replaced")); //ჩანაცვლებული

                    oColumn = oColumns.Add("DeclStatus", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("AttachedToTheDeclaration");
                    oColumn.TitleObject.Sortable = true;
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind("TxTable", "DeclStatus");

                    oColumn = oColumns.Add("DeclNum", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("DeclarationNumber");
                    oColumn.TitleObject.Sortable = true;
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind("TxTable", "DeclNum");

                    oColumn = oColumns.Add("Document", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("TaxInvoice");
                    oColumn.TitleObject.Sortable = true;
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oLink = oColumn.ExtendedObject;
                    oLink.LinkedObjectType = "UDO_F_BDO_TAXR_D";
                    oColumn.DataBind.Bind("TxTable", "Document");

                    oColumn = oColumns.Add("CorrDoc", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("CorrectedTaxInvoice");
                    oColumn.TitleObject.Sortable = true;
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oLink = oColumn.ExtendedObject;
                    oLink.LinkedObjectType = "UDO_F_BDO_TAXR_D";
                    oColumn.DataBind.Bind("TxTable", "CorrDoc");

                    oColumn = oColumns.Add("TxSerie", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("Series");
                    oColumn.TitleObject.Sortable = true;
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind("TxTable", "TxSerie");

                    oColumn = oColumns.Add("TxNum", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("Number");
                    oColumn.TitleObject.Sortable = true;
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind("TxTable", "TxNum");

                    oColumn = oColumns.Add("TxID", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("TaxInvoiceID");
                    oColumn.TitleObject.Sortable = true;
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind("TxTable", "TxID");

                    oColumn = oColumns.Add("IsVATPayer", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("IsVATPayer");
                    oColumn.TitleObject.Sortable = true;
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oColumn.ValidValues.Add("N", BDOSResources.getTranslate("NeedTax"));
                    oColumn.ValidValues.Add("Y", BDOSResources.getTranslate("NotNeddTax"));
                    oColumn.DataBind.Bind("TxTable", "IsVATPayer");
                    oColumn.DisplayDesc = true;
                    oColumn.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;

                    oColumn = oColumns.Add("CardCode", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("BPCardCode");
                    oColumn.TitleObject.Sortable = true;
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oLink = oColumn.ExtendedObject;
                    oLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_BusinessPartner;
                    oColumn.DataBind.Bind("TxTable", "CardCode");

                    oColumn = oColumns.Add("CardName", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("BPName");
                    oColumn.TitleObject.Sortable = true;
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind("TxTable", "CardName");

                    oColumn = oColumns.Add("VATno", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("BPTin");
                    oColumn.TitleObject.Sortable = true;
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind("TxTable", "VATno");

                    oColumn = oColumns.Add("OpDate", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("OperationDate");
                    oColumn.TitleObject.Sortable = true;
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind("TxTable", "OpDate");

                    oColumn = oColumns.Add("RegDate", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("RegistrationDate");
                    oColumn.TitleObject.Sortable = true;
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind("TxTable", "RegDate");

                    oColumn = oColumns.Add("ConfDate", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("ConfirmationDate");
                    oColumn.TitleObject.Sortable = true;
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind("TxTable", "ConfDate");

                    oColumn = oColumns.Add("Sum", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("Amount");
                    oColumn.TitleObject.Sortable = true;
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind("TxTable", "Sum");

                    oColumn = oColumns.Add("VatSum", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("VatAmount");
                    oColumn.TitleObject.Sortable = true;
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind("TxTable", "VatSum");

                    oColumn = oColumns.Add("Comment", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("Comment");
                    oColumn.TitleObject.Sortable = true;
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind("TxTable", "Comment");

                    oMatrix.Clear();
                    oMatrix.LoadFromDataSource();
                    oMatrix.AutoResizeColumns();

                    //გაცემული Pane 2 - საერთო სვეტებს იგივე სახელები დარჩა
                    SAPbouiCOM.Matrix oMatrix2 = ((SAPbouiCOM.Matrix)(oForm.Items.Item("TxTable2").Specific));
                    oMatrix2.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single;

                    oColumns = oMatrix2.Columns;

                    //SAPbouiCOM.DataTable oDataTable2;
                    oDataTable = oForm.DataSources.DataTables.Add("TxTable2");

                    oDataTable.Columns.Add("LineNum", SAPbouiCOM.BoFieldsType.ft_Text, 50); // 0 - ინდექსი გვჭირდება SetValue-ს პირველ პარამეტრად
                    oDataTable.Columns.Add("SrvStatus", SAPbouiCOM.BoFieldsType.ft_Text, 50); //1
                    oDataTable.Columns.Add("Status", SAPbouiCOM.BoFieldsType.ft_Text, 50); //2
                    oDataTable.Columns.Add("StatusDoc", SAPbouiCOM.BoFieldsType.ft_Text, 50); //3
                    oDataTable.Columns.Add("DeclStatus", SAPbouiCOM.BoFieldsType.ft_Text, 50); //4                
                    oDataTable.Columns.Add("DeclNum", SAPbouiCOM.BoFieldsType.ft_Text, 50); //5
                    oDataTable.Columns.Add("Document", SAPbouiCOM.BoFieldsType.ft_Text, 50); //6
                    oDataTable.Columns.Add("TxSerie", SAPbouiCOM.BoFieldsType.ft_Text, 50); //7
                    oDataTable.Columns.Add("TxNum", SAPbouiCOM.BoFieldsType.ft_Text, 50); //8
                    oDataTable.Columns.Add("TxID", SAPbouiCOM.BoFieldsType.ft_Text, 50); //9
                    oDataTable.Columns.Add("IsVATPayer", SAPbouiCOM.BoFieldsType.ft_Text, 50); //11
                    oDataTable.Columns.Add("CardCode", SAPbouiCOM.BoFieldsType.ft_Text, 50); //10
                    oDataTable.Columns.Add("CardName", SAPbouiCOM.BoFieldsType.ft_Text, 50); //11
                    oDataTable.Columns.Add("OpDate", SAPbouiCOM.BoFieldsType.ft_Date, 50); //11
                    oDataTable.Columns.Add("RegDate", SAPbouiCOM.BoFieldsType.ft_Date, 50); //12
                    oDataTable.Columns.Add("ConfDate", SAPbouiCOM.BoFieldsType.ft_Date, 50); //13
                    oDataTable.Columns.Add("Sum", SAPbouiCOM.BoFieldsType.ft_Sum, 50); //14
                    oDataTable.Columns.Add("VatSum", SAPbouiCOM.BoFieldsType.ft_Sum, 50); //15
                    oDataTable.Columns.Add("Comment", SAPbouiCOM.BoFieldsType.ft_Text, 50); //16
                    oDataTable.Columns.Add("TxChkBx", SAPbouiCOM.BoFieldsType.ft_Text, 50); //17
                    oDataTable.Columns.Add("VATno", SAPbouiCOM.BoFieldsType.ft_Text, 50); //18
                    oDataTable.Columns.Add("InvoiceEntry", SAPbouiCOM.BoFieldsType.ft_Text, 50); //19
                    oDataTable.Columns.Add("BaseARInvoice", SAPbouiCOM.BoFieldsType.ft_Text, 50); //20
                    oDataTable.Columns.Add("CorrDoc", SAPbouiCOM.BoFieldsType.ft_Text, 50); //21
                    //oDataTable.Columns.Add("CreditMemoEntry", SAPbouiCOM.BoFieldsType.ft_Text, 50); //22
                    oDataTable.Columns.Add("CorrType", SAPbouiCOM.BoFieldsType.ft_Text, 50); //23
                    oDataTable.Columns.Add("DocNum", SAPbouiCOM.BoFieldsType.ft_Text, 50); //24
                    oDataTable.Columns.Add("DocType", SAPbouiCOM.BoFieldsType.ft_Text, 50); //25

                    oColumn = oColumns.Add("LineNum", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = "#";
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind("TxTable2", "LineNum");

                    oColumn = oColumns.Add("TxChkBx", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
                    oColumn.TitleObject.Caption = "";
                    oColumn.Width = 100;
                    oColumn.Editable = true;
                    oColumn.DataBind.Bind("TxTable2", "TxChkBx");
                    oColumn.ValOn = "Y";
                    oColumn.ValOff = "N";

                    oUserTable = null;
                    oUserTable = Program.oCompany.UserTables.Item("BDO_TAXS");
                    StatusValidValues = oUserTable.UserFields.Fields.Item("U_status").ValidValues;

                    oColumn = oColumns.Add("StatusDoc", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("TaxInvoceStatus");
                    oColumn.Width = 200;
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind("TxTable2", "StatusDoc");
                    oColumn.DisplayDesc = true;
                    oColumn.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
                    oColumn.ValidValues.Add("empty", "");
                    oColumn.ValidValues.Add("created", BDOSResources.getTranslate("Created"));
                    oColumn.ValidValues.Add("shipped", BDOSResources.getTranslate("Sent"));
                    oColumn.ValidValues.Add("confirmed", BDOSResources.getTranslate("Confirmed"));
                    oColumn.ValidValues.Add("removed", BDOSResources.getTranslate("deleted"));
                    oColumn.ValidValues.Add("incompleteShipped", BDOSResources.getTranslate("CreatedIncompletely"));
                    oColumn.ValidValues.Add("paper", BDOSResources.getTranslate("Paper"));
                    oColumn.ValidValues.Add("disturbedSynchronization", BDOSResources.getTranslate("SynchronizationViolated"));
                    oColumn.ValidValues.Add("denied", BDOSResources.getTranslate("Denied"));
                    oColumn.ValidValues.Add("cancellationProcess", BDOSResources.getTranslate("CancellationProcess"));
                    oColumn.ValidValues.Add("canceled", BDOSResources.getTranslate("Canceled"));
                    oColumn.ValidValues.Add("attachedToTheDeclaration", BDOSResources.getTranslate("RSAddDeclaration"));
                    oColumn.ValidValues.Add("correctionCreated", BDOSResources.getTranslate("CreatedCorrected"));
                    oColumn.ValidValues.Add("correctionShipped", BDOSResources.getTranslate("SentCorrected"));
                    oColumn.ValidValues.Add("correctionConfirmed", BDOSResources.getTranslate("ConfirmedCorrected"));
                    oColumn.ValidValues.Add("primary", BDOSResources.getTranslate("Primary"));
                    oColumn.ValidValues.Add("corrected", BDOSResources.getTranslate("Corrected"));

                    oColumn = oColumns.Add("DeclStatus", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("AttachedToTheDeclaration");
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind("TxTable2", "DeclStatus");

                    oColumn = oColumns.Add("DeclNum", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("DeclarationNumber");
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind("TxTable2", "DeclNum");

                    oColumn = oColumns.Add("DocType", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("DocType");
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind("TxTable2", "DocType");
                    oColumn.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
                    oColumn.DisplayDesc = true;
                    oColumn.ValidValues.Add("ARInvoice", BDOSResources.getTranslate("ARInvoice"));
                    oColumn.ValidValues.Add("ARCreditNote", BDOSResources.getTranslate("ARCreditNote"));
                    oColumn.ValidValues.Add("ARDownPaymentVAT", BDOSResources.getTranslate("ARDownPaymentVAT"));

                    oColumn = oColumns.Add("InvEntry", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("Invoice");
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind("TxTable2", "InvoiceEntry");

                    oColumn = oColumns.Add("DocNum", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("DocNum");
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind("TxTable2", "DocNum");

                    oColumn = oColumns.Add("InvEntryB", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("BaseARInvoice");
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oLink = oColumn.ExtendedObject;
                    oLink.LinkedObjectType = "13";
                    oColumn.DataBind.Bind("TxTable2", "BaseARInvoice");

                    oColumn = oColumns.Add("Document", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("TaxInvoice");
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oLink = oColumn.ExtendedObject;
                    oLink.LinkedObjectType = "UDO_F_BDO_TAXS_D";
                    oColumn.DataBind.Bind("TxTable2", "Document");

                    oColumn = oColumns.Add("TxSerie", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("Series");
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind("TxTable2", "TxSerie");

                    oColumn = oColumns.Add("TxNum", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("Number");
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind("TxTable2", "TxNum");

                    oColumn = oColumns.Add("TxID", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("TaxInvoiceID");
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind("TxTable2", "TxID");

                    oColumn = oColumns.Add("CorrDoc", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("CorrectedTaxInvoice");
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oLink = oColumn.ExtendedObject;
                    oLink.LinkedObjectType = "UDO_F_BDO_TAXS_D";
                    oColumn.DataBind.Bind("TxTable2", "CorrDoc");

                    oUserTable = Program.oCompany.UserTables.Item("BDO_TAXS");
                    oColumn = oColumns.Add("CorrType", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("Reason");
                    oColumn.Width = 200;
                    oColumn.Editable = true;
                    oColumn.DataBind.Bind("TxTable2", "CorrType");
                    oColumn.DisplayDesc = true;
                    oColumn.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
                    oColumn.ValidValues.Add("-1", "");
                    oColumn.ValidValues.Add("1", BDOSResources.getTranslate("CanceledTaxOperation")); //1 //გაუქმებულია დასაბეგრი ოპერაცია
                    oColumn.ValidValues.Add("2", BDOSResources.getTranslate("ChangedTaxOperationType")); //2 //შეცვლილია დასაბეგრი ოპერაციის სახე
                    oColumn.ValidValues.Add("3", BDOSResources.getTranslate("ChangedAgreementAmountPricesDecrease")); //3 //ფასების შემცირების ან სხვა მიზეზით შეცვლილია ოპერაციაზე ადრე შეთანხმებული კომპენსაციის თანხა
                    oColumn.ValidValues.Add("4", BDOSResources.getTranslate("ItemServiceReturnedToSeller")); //4 საქონელი (მომსახურება) სრულად ან ნაწილობრივ უბრუნდება გამყიდველს


                    oColumn = oColumns.Add("IsVATPayer", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("IsVATPayer");
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oColumn.ValidValues.Add("N", BDOSResources.getTranslate("NeedTax"));
                    oColumn.ValidValues.Add("Y", BDOSResources.getTranslate("NotNeedTax"));
                    oColumn.DataBind.Bind("TxTable2", "IsVATPayer");
                    oColumn.DisplayDesc = true;
                    oColumn.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;

                    oColumn = oColumns.Add("CardCode", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("BPCardCode");
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oLink = oColumn.ExtendedObject;
                    oLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_BusinessPartner;
                    oColumn.DataBind.Bind("TxTable2", "CardCode");

                    oColumn = oColumns.Add("CardName", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("BPName");
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind("TxTable2", "CardName");

                    oColumn = oColumns.Add("VATno", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("BPTin");
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind("TxTable2", "VATno");

                    oColumn = oColumns.Add("OpDate", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("OperationDate");
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind("TxTable2", "OpDate");

                    oColumn = oColumns.Add("RegDate", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("RegistrationDate");
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind("TxTable2", "RegDate");

                    oColumn = oColumns.Add("ConfDate", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("ConfirmationDate");
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind("TxTable2", "ConfDate");

                    oColumn = oColumns.Add("Sum", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("AmountWithVat");
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind("TxTable2", "Sum");

                    oColumn = oColumns.Add("VatSum", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("VatAmount");
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind("TxTable2", "VatSum");

                    oColumn = oColumns.Add("Comment", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("Comment");
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind("TxTable2", "Comment");

                    oMatrix2.Clear();
                    oMatrix2.LoadFromDataSource();
                    oMatrix2.AutoResizeColumns();
                    //გაცემული Pane 2

                }

                createFolder(  oForm, out errorText);

                oForm.Visible = true;
                oForm.Select();

                oForm.Freeze(true);
                oForm.Items.Item("Folder2").Click();
                oForm.Freeze(false);
            }
            GC.Collect();
        }

        public static void createFolder(  SAPbouiCOM.Form oForm, out string errorText)
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
                        folderName = BDOSResources.getTranslate("TaxInvoiceReceived");
                    }
                    else
                    {
                        folderName = BDOSResources.getTranslate("TaxInvoiceSent");
                    }

                    //SAPbouiCOM.Folder oFolder = (SAPbouiCOM.Folder)oForm.Items.Add("Folder" + i.ToString(), SAPbouiCOM.BoFormItemTypes.it_FOLDER).Specific;
                    //oFolder.Caption = "folder " + i.ToString();
                    //oFolder.DataBind.SetBound(true, "", "FolderDS");

                    Dictionary<string, object> formItems = new Dictionary<string, object>();
                    string itemName = "Folder" + i.ToString();
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_FOLDER);
                    formItems.Add("Bound", true);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", "FolderDS");
                    formItems.Add("Width", 200);
                    formItems.Add("Top", 5);
                    formItems.Add("Height", 10);
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

        public static void checkUncheckTaxes(SAPbouiCOM.Form oForm, string CheckOperation, string type, out string errorText)
        {
            errorText = null;
            try
            {
                oForm.Freeze(true);
                SAPbouiCOM.CheckBox oCheckBox;
                SAPbouiCOM.Matrix oMatrix;
                if (CheckOperation == "TxCheck" || CheckOperation == "TxUncheck")
                {
                    oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("TxTable").Specific));
                }
                else
                {
                    oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("TxTable2").Specific));
                }
                int rowCount = oMatrix.RowCount;
                for (int j = 1; j <= rowCount; j++)
                {
                    oCheckBox = oMatrix.Columns.Item("TxChkBx").Cells.Item(j).Specific;
                    oCheckBox.Checked = (CheckOperation == "TxCheck") || (CheckOperation == "TxCheck2");
                }
                oForm.Freeze(false);
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

        public static void chooseFromList(SAPbouiCOM.Form oForm, SAPbouiCOM.ItemEvent pVal, SAPbouiCOM.IChooseFromListEvent oCFLEvento, out string errorText)
        {
            errorText = null;
            try
            {
                string sCFL_ID = oCFLEvento.ChooseFromListUID;
                SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(sCFL_ID);

                if (pVal.BeforeAction == true)
                {
                    //if (sCFL_ID == "BP_CFL")
                    //{
                    //    SAPbouiCOM.Conditions oCons = oCFL.GetConditions();
                    //    SAPbouiCOM.Condition oCon = oCons.Add();
                    //    oCon.Alias = "CardType";
                    //    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    //    oCon.CondVal = "S"; //მომწოდებელი
                    //    oCFL.SetConditions(oCons);
                    //}

                    //if (sCFL_ID == "BP_CFL2")
                    //{
                    //    SAPbouiCOM.Conditions oCons = oCFL.GetConditions();
                    //    SAPbouiCOM.Condition oCon = oCons.Add();
                    //    oCon.Alias = "CardType";
                    //    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    //    oCon.CondVal = "C"; //მყიდველი
                    //    oCFL.SetConditions(oCons);
                    //}
                }
                else if (pVal.BeforeAction == false)
                {
                    SAPbouiCOM.DataTable oDataTable = null;
                    oDataTable = oCFLEvento.SelectedObjects;

                    if (oDataTable != null)
                    {
                        if (sCFL_ID == "BP_CFL")
                        {
                            SAPbouiCOM.DataTable oDataTableSelectedObjects = oCFLEvento.SelectedObjects;
                            string CardCode = oDataTableSelectedObjects.GetValue("CardCode", 0);

                            SAPbouiCOM.EditText BPEdit = oForm.Items.Item("CCode").Specific;
                            BPEdit.Value = CardCode;
                        }

                        if (sCFL_ID == "BP_CFL2")
                        {
                            SAPbouiCOM.DataTable oDataTableSelectedObjects = oCFLEvento.SelectedObjects;
                            string CardCode = oDataTableSelectedObjects.GetValue("CardCode", 0);

                            SAPbouiCOM.EditText BPEdit = oForm.Items.Item("CCode2").Specific;
                            BPEdit.Value = CardCode;
                        }
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

        public static object GetDocumentProperty( string UDOName, int key, string propertyName, out string errorText)
        {
            errorText = null;
            try
            {
                SAPbobsCOM.CompanyService oCompanyService = null;
                SAPbobsCOM.GeneralService oGeneralService = null;
                SAPbobsCOM.GeneralData oGeneralData = null;
                SAPbobsCOM.GeneralDataParams oGeneralParams = null;
                oCompanyService = Program.oCompany.GetCompanyService();
                oGeneralService = oCompanyService.GetGeneralService(UDOName);
                oGeneralData = ((SAPbobsCOM.GeneralData)(oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)));

                oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                oGeneralParams.SetProperty("DocEntry", key);
                oGeneralData = oGeneralService.GetByParams(oGeneralParams);

                return oGeneralData.GetProperty(propertyName);
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }

            return null;
        }

        public static void itemPressed(SAPbouiCOM.Form oForm, SAPbouiCOM.ItemEvent pVal, out string errorText)
        {
            errorText = null;

            try
            {
                if (pVal.ItemUID == "StartDatOp")
                {
                    SAPbouiCOM.EditText oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("StartDatOp").Specific;
                    DateTime StartDateOp = DateTime.ParseExact(oEditText.Value, "yyyyMMdd", null);
                    StartDateOp = new DateTime(StartDateOp.Year, StartDateOp.Month, 1);
                    oEditText.Value = StartDateOp.ToString("yyyyMMdd");
                }
                else if (pVal.ItemUID == "EndDateOp")
                {
                    SAPbouiCOM.EditText oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("EndDateOp").Specific;
                    DateTime EndDateOp = DateTime.ParseExact(oEditText.Value, "yyyyMMdd", null);
                    EndDateOp = new DateTime(EndDateOp.Year, EndDateOp.Month, 1).AddMonths(1).AddDays(-1);
                    oEditText.Value = EndDateOp.ToString("yyyyMMdd");
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

        public static void uiApp_ItemEvent(  string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
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

                if ((pVal.ItemUID == "StrtDatOp2" || pVal.ItemUID == "EndDateOp2") && (pVal.ItemChanged) && pVal.BeforeAction == false)
                {
                    string startDateStr = oForm.DataSources.UserDataSources.Item("StrtDatOp2").ValueEx.ToString();

                    DateTime startDate = DateTime.ParseExact(startDateStr, "yyyyMMdd", CultureInfo.InvariantCulture);
                    DateTime OperationPeriodStart = new DateTime(startDate.Year, startDate.Month, 1);   
                    oForm.DataSources.UserDataSources.Item("StrtDatOp2").ValueEx = OperationPeriodStart.ToString("yyyyMMdd");

                    string endDateStr = oForm.DataSources.UserDataSources.Item("StrtDatOp2").ValueEx.ToString();
                    DateTime endDate = DateTime.ParseExact(startDateStr, "yyyyMMdd", CultureInfo.InvariantCulture);
                    DateTime OperationPeriodEnd = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month));
                    OperationPeriodEnd = OperationPeriodEnd.AddSeconds(24 * 3600 - 1);

                    oForm.DataSources.UserDataSources.Item("EndDateOp2").ValueEx = OperationPeriodEnd.ToString("yyyyMMdd");

                    oForm.Freeze(false);
                }

                if ((pVal.ItemUID == "StrtDatOp" || pVal.ItemUID == "EndDateOp") && (pVal.ItemChanged) && pVal.BeforeAction == false)
                {
                    string startDateStr = oForm.DataSources.UserDataSources.Item("StrtDatOp").ValueEx.ToString();

                    DateTime startDate = DateTime.ParseExact(startDateStr, "yyyyMMdd", CultureInfo.InvariantCulture);
                    DateTime OperationPeriodStart = new DateTime(startDate.Year, startDate.Month, 1);
                    oForm.DataSources.UserDataSources.Item("StrtDatOp").ValueEx = OperationPeriodStart.ToString("yyyyMMdd");

                    string endDateStr = oForm.DataSources.UserDataSources.Item("StrtDatOp").ValueEx.ToString();
                    DateTime endDate = DateTime.ParseExact(startDateStr, "yyyyMMdd", CultureInfo.InvariantCulture);
                    DateTime OperationPeriodEnd = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month));
                    OperationPeriodEnd = OperationPeriodEnd.AddSeconds(24 * 3600 - 1);

                    oForm.DataSources.UserDataSources.Item("EndDateOp").ValueEx = OperationPeriodEnd.ToString("yyyyMMdd");

                    oForm.Freeze(false);
                }

                //ბიზნეს პარტნიორის არჩევა
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                {
                    SAPbouiCOM.ChooseFromListEvent oCFLEvento = null;
                    oCFLEvento = ((SAPbouiCOM.ChooseFromListEvent)(pVal));
                    chooseFromList(oForm, pVal, oCFLEvento, out errorText);
                }

                //გაცემულები
                //შევსება
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK & pVal.ItemUID == "fillFrmBs2" & pVal.BeforeAction == false)
                {
                    //SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);
                    fillFromBaseTaxInvoiceSent(  oForm, false, out errorText);
                }
                //გაცემულები

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK & pVal.ItemUID == "dwnldTax" & pVal.BeforeAction == false)
                {
                    downloadTaxInvoiceReceived(  oForm, out errorText);
                    fillFromBaseTaxInvoiceReceived(  oForm, true, out errorText);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK & pVal.ItemUID == "updtTax" & pVal.BeforeAction == false)
                {
                    updateTaxInvoiceReceived(  oForm, out errorText);
                    //fillFromBase(  oForm, out errorText);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK & pVal.ItemUID == "fillFrmBs" & pVal.BeforeAction == false)
                {
                    fillFromBaseTaxInvoiceReceived(  oForm, false, out errorText);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK & pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "TxCheck" || pVal.ItemUID == "TxUncheck")
                    {
                        //SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);
                        checkUncheckTaxes(oForm, pVal.ItemUID, "", out errorText);
                    }

                    if (pVal.ItemUID == "TxCheck2" || pVal.ItemUID == "TxUncheck2")
                    {
                        //SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);
                        checkUncheckTaxes(oForm, pVal.ItemUID, "2", out errorText);
                    }
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_COMBO_SELECT & pVal.BeforeAction == false)
                {
                    //SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                    if (pVal.ItemUID == "TxOperRS")
                    {
                        SAPbouiCOM.ButtonCombo oTxOperRS = ((SAPbouiCOM.ButtonCombo)(oForm.Items.Item("TxOperRS").Specific));
                        oTxOperRS.Caption = BDOSResources.getTranslate("Operations");
                        int oOperation = pVal.PopUpIndicator;
                        rsOperationTaxInvoiceReceived( oForm, oOperation, out errorText);
                    }

                    if (pVal.ItemUID == "TxOperRS2")
                    {
                        SAPbouiCOM.ButtonCombo oTxOperRS = ((SAPbouiCOM.ButtonCombo)(oForm.Items.Item("TxOperRS2").Specific));
                        oTxOperRS.Caption = BDOSResources.getTranslate("Operations");
                        int oOperation = pVal.PopUpIndicator;
                        rsOperationTaxInvoiceSent( oForm, oOperation, out errorText);
                    }

                    if (pVal.ItemUID == "addDecl")
                    {
                        SAPbouiCOM.ButtonCombo oaddDecl = ((SAPbouiCOM.ButtonCombo)(oForm.Items.Item("addDecl").Specific));
                        oaddDecl.Caption = BDOSResources.getTranslate("Add");
                        int oOperation = pVal.PopUpIndicator;

                        if (oOperation == 0)
                        {
                            addDeclTaxInvoiceReceived(  oForm, out errorText);
                        }
                        else
                        {
                            receiveVATTaxInvoiceReceived(  oForm, out errorText);
                        }
                    }
                    if (pVal.ItemUID == "addDecl2")
                    {
                        SAPbouiCOM.ButtonCombo oaddDecl = ((SAPbouiCOM.ButtonCombo)(oForm.Items.Item("addDecl2").Specific));
                        oaddDecl.Caption = BDOSResources.getTranslate("Add");
                        int oOperation = pVal.PopUpIndicator;

                        if (oOperation == 0)
                        {
                            addDeclTaxInvoiceSent(  oForm, out errorText);
                        }
                    }
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED & pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "StartDatOp" || pVal.ItemUID == "EndDateOp")
                    {
                        itemPressed(oForm, pVal, out errorText);
                    }
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_GOT_FOCUS)
                {
                    matrixColumnSetCfl(oForm, pVal, out errorText);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED)
                {
                    matrixColumnSetCfl(oForm, pVal, out errorText);
                }
            }
        }
    }
}
