using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using BDO_Localisation_AddOn.TBC_Integration_Services;
using BDO_Localisation_AddOn.BOG_Integration_Services;
using System.Text.RegularExpressions;
using BDO_Localisation_AddOn.BOG_Integration_Services.Model;
using System.Data;

namespace BDO_Localisation_AddOn
{
    static partial class BDOSInternetBanking
    {
        public static AccountMovementDetailIo[] oAccountMovementDetailIoStc = null;
        public static BaseQueryResultIo oBaseQueryResultIoStc = null;
        public static List<StatementDetail> oStatementDetailStc = null;
        public static int CurrentRowExportMTRForDetail;
        public static DataTable TableExportMTRForDetail;
        public static bool SelectAllImportPressed = false;
        public static List<int> RowsWithDifferentDates;
        public static List<int> RowsWithCorrectedDates;
        public static string blnkAgrOld;

        #region Import Data into Bank System
        public static void chooseFromListImport(SAPbouiCOM.Form oForm, SAPbouiCOM.IChooseFromListEvent oCFLEvento, SAPbouiCOM.ItemEvent pVal, out string errorText)
        {
            errorText = null;

            NumberFormatInfo Nfi = new NumberFormatInfo() { NumberDecimalSeparator = "." };

            try
            {
                string sCFL_ID = oCFLEvento.ChooseFromListUID;
                SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(sCFL_ID);

                if (pVal.BeforeAction == true)
                {
                }
                else if (!pVal.BeforeAction)
                {
                    SAPbouiCOM.DataTable oDataTable = null;
                    oDataTable = oCFLEvento.SelectedObjects;

                    if (oDataTable != null)
                    {
                        if (sCFL_ID == "HouseBankAccount_CFL")
                        {
                            string account = Convert.ToString(oDataTable.GetValue("Account", 0));
                            SAPbouiCOM.EditText oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("accountE").Specific;
                            try
                            {
                                oEditText.Value = account;
                            }
                            catch { }

                            oForm.Freeze(true);
                            setVisibleFormItemsImport(oForm, out errorText);
                            oForm.Update();
                            oForm.Freeze(false);
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

        public static void comboSelectImport(SAPbouiCOM.Form oForm, SAPbouiCOM.ItemEvent pVal)
        {
            try
            {
                string errorText;

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

                        if (selectedOperation != null)
                        {
                            string importType = oForm.DataSources.UserDataSources.Item("imptTypeCB").ValueEx;
                            bool batchPayment = false;
                            string batchName = null;

                            string account = oForm.DataSources.UserDataSources.Item("accountE").ValueEx;
                            string bankProgram = CommonFunctions.getBankProgram(null, account);

                            if (importType == "batchPayment" && selectedOperation == "import")
                            {
                                batchName = oForm.DataSources.UserDataSources.Item("batchNameE").ValueEx;
                                if (string.IsNullOrEmpty(batchName) && bankProgram == "TBC")
                                {
                                    errorText = BDOSResources.getTranslate("TheFollowingFieldIsMandatory") + " : \"" + BDOSResources.getTranslate("BatchName") + "\"";
                                    Program.uiApp.SetStatusBarMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                    return;
                                }
                                batchPayment = true;
                            }

                            List<int> docEntryList = new List<int>();
                            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("importMTR").Specific));
                            int row = 1;

                            while (row <= oMatrix.RowCount)
                            {
                                if (oMatrix.GetCellSpecific("CheckBox", row).Checked)
                                {
                                    int docEntry = Convert.ToInt32(oMatrix.GetCellSpecific("DocEntry", row).Value);
                                    docEntryList.Add(docEntry);
                                }
                                row++;
                            }
                            if (docEntryList.Count == 0)
                            {
                                errorText = BDOSResources.getTranslate("HighlightTheRowsToPerformAnOperation") + "!"; //"ოპერაციის შესასრულებლად მონიშნეთ სასურველი სტრიქონები!";
                                Program.uiApp.SetStatusBarMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                return;
                            }
                            if (importType == "singlePayment" && docEntryList.Count > 10 && selectedOperation == "import")
                            {
                                errorText = BDOSResources.getTranslate("YouCannotTransferMoreThan10DocumentWhenImportTypeIs") + " " + BDOSResources.getTranslate("SinglePayment") + "!"; //"არ შეიძლება 10 დოკუმენტზე მეტის გადარიცხვა როცა იმპორტის ტიპი არის ინდივიდუალური გადარიცხვა!";
                                Program.uiApp.SetStatusBarMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                return;
                            }

                            if (bankProgram == "TBC")
                            {
                                BDOSAuthenticationFormTBC.createForm(oForm, selectedOperation, docEntryList, batchPayment, batchName, null, out errorText);
                            }
                            else if (bankProgram == "BOG")
                            {
                                BDOSAuthenticationFormBOG.createForm(oForm, selectedOperation, docEntryList, batchPayment, batchName, null, out errorText);
                            }
                            if (string.IsNullOrEmpty(bankProgram))
                            {
                                errorText = BDOSResources.getTranslate("BankProgramIsNotFilled") + "!";
                                Program.uiApp.SetStatusBarMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                return;
                            }
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

                        if (selectedOperation != null)
                        {
                            List<int> docEntryList = new List<int>();
                            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("importMTR").Specific));
                            int row = 1;

                            while (row <= oMatrix.RowCount)
                            {
                                if (oMatrix.GetCellSpecific("CheckBox", row).Checked)
                                {
                                    int docEntry = Convert.ToInt32(oMatrix.GetCellSpecific("DocEntry", row).Value);
                                    docEntryList.Add(docEntry);
                                }
                                row++;
                            }
                            if (docEntryList.Count == 0)
                            {
                                errorText = BDOSResources.getTranslate("HighlightTheRowsToPerformAnOperation") + "!"; //"ოპერაციის შესასრულებლად მონიშნეთ სასურველი სტრიქონები!";
                                Program.uiApp.SetStatusBarMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                return;
                            }

                            int answer = Program.uiApp.MessageBox(BDOSResources.getTranslate("SelectedDocumentsStatusWillBe") + " \"" + BDOSResources.getTranslate(selectedOperation) + "\". " + BDOSResources.getTranslate("WouldYouWantToContinueTheOperation") + "?", 1, BDOSResources.getTranslate("Yes"), BDOSResources.getTranslate("No"), ""); //მონიშნული დოკუმენტების სტატუსი გახდება, გსურთ ოპერაციის გაგრძელება

                            if (answer == 2)
                            {
                                return;
                            }

                            List<string> infoList = setStatusImport(docEntryList, selectedOperation);
                            BDOSInternetBanking.fillImportMTR(oForm, out errorText);
                            for (int i = 0; i < infoList.Count; i++)
                            {
                                Program.uiApp.SetStatusBarMessage(infoList[i], SAPbouiCOM.BoMessageTime.bmt_Short, false);
                            }
                            return;
                        }
                    }
                }
                else
                {
                    if (pVal.ItemUID == "operationB")
                    {
                        oForm.Freeze(true);
                    }
                    else if (pVal.ItemUID == "setStatusB")
                    {
                        oForm.Freeze(true);
                    }
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

        public static List<string> setStatusImport(List<int> docEntryList, string selectedStatus)
        {
            string info = null;
            List<string> infoList = new List<string>();
            int docEntry;

            for (int i = 0; i < docEntryList.Count; i++)
            {
                docEntry = docEntryList[i];

                SAPbobsCOM.Payments oVendorPayments;
                oVendorPayments = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments);
                oVendorPayments.GetByKey(docEntry);
                string status = oVendorPayments.UserFields.Fields.Item("U_status").Value;
                string tempStatus = null;

                if ((status == "finishedWithErrors" || status == "cancelled" || status == "failed") && selectedStatus == "resend")
                {
                    tempStatus = selectedStatus;
                }
                else if ((status == "resend" || status == "readyToLoad") && selectedStatus == "notToUpload")
                {
                    tempStatus = selectedStatus;
                }
                else if (status == "notToUpload" && selectedStatus == "readyToLoad")
                {
                    tempStatus = selectedStatus;
                }

                if (string.IsNullOrEmpty(tempStatus) == false)
                {
                    oVendorPayments.UserFields.Fields.Item("U_status").Value = tempStatus;

                    int returnCode = oVendorPayments.Update();
                    if (returnCode != 0)
                    {
                        int errCode;
                        string errMsg;
                        Program.oCompany.GetLastError(out errCode, out errMsg);
                        info = BDOSResources.getTranslate("CannotSetStatus") + " \"" + BDOSResources.getTranslate(selectedStatus) + "\"! " + BDOSResources.getTranslate("Document") + " : " + docEntry; //ვერ მიენიჭა სტატუსი
                        infoList.Add(info);
                    }
                    else
                    {
                        info = BDOSResources.getTranslate("SuccessfullySetStatus") + " \"" + BDOSResources.getTranslate(selectedStatus) + "\"! " + BDOSResources.getTranslate("Document") + " : " + docEntry; //წარმატებით მიენიჭა სტატუსი
                        infoList.Add(info);
                    }
                }
                else
                {
                    if (selectedStatus == "resend")
                        info = BDOSResources.getTranslate("DocumentSStatusShouldBeTheOneFromTheList") + " : \"" + BDOSResources.getTranslate("finishedWithErrors") + "\", \"" + BDOSResources.getTranslate("cancelled") + "\", \"" + BDOSResources.getTranslate("failed") + "\"! " + BDOSResources.getTranslate("Document") + " : " + docEntry;
                    else if (selectedStatus == "notToUpload")
                        info = BDOSResources.getTranslate("DocumentSStatusShouldBeTheOneFromTheList") + " : \"" + BDOSResources.getTranslate("resend") + "\", \"" + BDOSResources.getTranslate("readyToLoad") + "\"! " + BDOSResources.getTranslate("Document") + " : " + docEntry;
                    else if (selectedStatus == "readyToLoad")
                        info = BDOSResources.getTranslate("DocumentSStatusShouldBeTheOneFromTheList") + " : \"" + BDOSResources.getTranslate("notToUpload") + "\"! " + BDOSResources.getTranslate("Document") + " : " + docEntry;
                    infoList.Add(info);
                }
            }
            return infoList;
        }

        public static void checkUncheckMTRImport(SAPbouiCOM.Form oForm, string checkOperation, out string errorText)
        {
            errorText = null;
            try
            {

                oForm.Freeze(true);
                SAPbouiCOM.CheckBox oCheckBox;
                SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("importMTR").Specific));

                for (int j = 1; j <= oMatrix.RowCount; j++)
                {
                    oCheckBox = oMatrix.Columns.Item("CheckBox").Cells.Item(j).Specific;

                    oCheckBox.Checked = (checkOperation == "checkB");
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

        public static void matrixColumnSetLinkedObjectTypeImport(SAPbouiCOM.Form oForm, SAPbouiCOM.ItemEvent pVal)
        {
            try
            {
                if (pVal.ColUID == "CardCode")
                {
                    if (pVal.BeforeAction)
                    {
                        SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("importMTR").Specific));

                        SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("importMTR");
                        string opType = oDataTable.GetValue("OpType", pVal.Row - 1).ToString();

                        SAPbouiCOM.Column oColumn;

                        if (opType == "paymentToEmployee")
                        {
                            oColumn = oMatrix.Columns.Item(pVal.ColUID);
                            SAPbouiCOM.LinkedButton oLink = oColumn.ExtendedObject;
                            oLink.LinkedObjectType = "171"; //SAPbouiCOM.BoLinkedObject.lf_Employee
                        }
                        else
                        {
                            oColumn = oMatrix.Columns.Item(pVal.ColUID);
                            SAPbouiCOM.LinkedButton oLink = oColumn.ExtendedObject;
                            oLink.LinkedObjectType = "2";
                        }
                    }
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

        public static void setVisibleFormItemsImport(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            SAPbouiCOM.Item oItem = null;

            try
            {
                string importType = oForm.DataSources.UserDataSources.Item("imptTypeCB").ValueEx;

                string account = oForm.DataSources.UserDataSources.Item("accountE").ValueEx;
                string bankProgram = CommonFunctions.getBankProgram(null, account);

                if (importType == "batchPayment" && bankProgram == "TBC")
                {
                    oItem = oForm.Items.Item("batchNameE");
                    oItem.Visible = true;
                }
                else
                {
                    oItem = oForm.Items.Item("batchNameE");
                    oItem.Visible = false;
                }

                oForm.Items.Item("DocTypeS").Visible = CommonFunctions.isHRAddOnConnected();
                oForm.Items.Item("DocTypeCB").Visible = CommonFunctions.isHRAddOnConnected();

                oItem = oForm.Items.Item("rprtCodeS");
                oItem.Visible = false;

                oItem = oForm.Items.Item("rprtCodeCB");
                oItem.Visible = false;

                if (bankProgram == "BOG" && account.Contains("GEL") == false)
                {
                    oItem = oForm.Items.Item("rprtCodeS");
                    oItem.Visible = true;

                    oItem = oForm.Items.Item("rprtCodeCB");
                    oItem.Visible = true;

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

        public static void setVisibleFormItemsMatrixColumns(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            try
            {
                SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("exportMTR").Specific));
                oMatrix.Columns.Item("DocNum").Visible = (oForm.Items.Item("listView").Specific.Value.Trim() == "withDetl");
                oMatrix.Columns.Item("ExPaymntID").Visible = (oForm.Items.Item("listView").Specific.Value.Trim() == "withDetl");
                oMatrix.Columns.Item("ValueDate").Visible = (oForm.Items.Item("listView").Specific.Value.Trim() == "withDetl");
                oMatrix.Columns.Item("AcctNumber").Visible = (oForm.Items.Item("listView").Specific.Value.Trim() == "withDetl");
                oMatrix.Columns.Item("CurrencyEx").Visible = (oForm.Items.Item("listView").Specific.Value.Trim() == "withDetl");
                oMatrix.Columns.Item("PCurrency").Visible = (oForm.Items.Item("listView").Specific.Value.Trim() == "withDetl");
                oMatrix.Columns.Item("DocNumber").Visible = (oForm.Items.Item("listView").Specific.Value.Trim() == "withDetl");
                oMatrix.Columns.Item("OpCode").Visible = (oForm.Items.Item("listView").Specific.Value.Trim() == "withDetl");
                oMatrix.Columns.Item("BPCode").Visible = (oForm.Items.Item("listView").Specific.Value.Trim() == "withDetl");

                oMatrix.Columns.Item("DocEntry").Width = 40;
                oMatrix.Columns.Item("Descrpt").Width = 50;
                oMatrix.Columns.Item("AddDescrpt").Width = 50;
                oMatrix.Columns.Item("CFWId").Width = 20;
                oMatrix.Columns.Item("BCFWId").Width = 20;

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

        public static void fillImportMTR(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("importMTR");
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string startDate = oForm.DataSources.UserDataSources.Item("startDateE").ValueEx;
            string endDate = oForm.DataSources.UserDataSources.Item("endDateE").ValueEx;
            string account = oForm.DataSources.UserDataSources.Item("accountE").ValueEx;
            string allDocsOB = oForm.DataSources.UserDataSources.Item("allDocsOB").ValueEx;
            string uplDocsOB = oForm.DataSources.UserDataSources.Item("uplDocsOB").ValueEx;

            if (string.IsNullOrEmpty(startDate) || string.IsNullOrEmpty(endDate) || string.IsNullOrEmpty(account))
            {
                errorText = BDOSResources.getTranslate("TheFollowingFieldsAreMandatory") + " : \"" + oForm.Items.Item("periodS").Specific.caption + "\", \"" + oForm.Items.Item("accountS").Specific.caption + "\""; //"პერიოდის და ანგარიშის ნომრის მითითება აუცილებელია!";
                Program.uiApp.SetStatusBarMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                return;
            }
            bool allDocuments;
            if (allDocsOB == "1")
            {
                allDocuments = true;
            }
            else if (allDocsOB == "2")
            {
                allDocuments = false;
            }
            else
            {
                errorText = BDOSResources.getTranslate("SelectFillType") + "!"; //მიუთითეთ შევსების ტიპი
                Program.uiApp.SetStatusBarMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                return;
            }
            string bankProgram = CommonFunctions.getBankProgram(null, account);

            if (string.IsNullOrEmpty(bankProgram))
            {
                errorText = BDOSResources.getTranslate("BankProgramIsNotFilled") + "!";
                Program.uiApp.SetStatusBarMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                return;
            }
            
            var docType = oForm.DataSources.UserDataSources.Item("DocTypeCB").ValueEx == "2" ? "paymentToEmployee" : "";

            string query = OutgoingPayment.getQueryForImport(null, account, startDate, endDate, bankProgram, allDocuments, docType);
            string queryOnlyLocalisationAddOn = OutgoingPayment.getQueryForImportOnlyLocalisationAddOn(null, account, startDate, endDate, bankProgram, allDocuments);
            try
            {
                oRecordSet.DoQuery(query);
            }
            catch
            {
                oRecordSet.DoQuery(queryOnlyLocalisationAddOn);
            }
            oDataTable.Rows.Clear();

            try
            {
                int rowIndex = 0;

                while (!oRecordSet.EoF)
                {
                    Dictionary<string, string> dataForTransferType = OutgoingPayment.getDataForTransferType(oRecordSet);
                    string transferType = OutgoingPayment.getTransferType(dataForTransferType, out errorText);
                    Dictionary<string, object> dataForImport = OutgoingPayment.getDataForImport(oRecordSet, dataForTransferType, transferType);
                    oDataTable.Rows.Add();
                    oDataTable.SetValue("LineNum", rowIndex, rowIndex + 1);
                    oDataTable.SetValue("CheckBox", rowIndex, "N");
                    oDataTable.SetValue("DocEntry", rowIndex, oRecordSet.Fields.Item("DocEntry").Value);
                    oDataTable.SetValue("DocNum", rowIndex, oRecordSet.Fields.Item("DocNum").Value);
                    oDataTable.SetValue("DocDate", rowIndex, oRecordSet.Fields.Item("DocDate").Value.ToString("yyyyMMdd") == "18991230" ? "" : oRecordSet.Fields.Item("DocDate").Value.ToString("yyyyMMdd"));
                    oDataTable.SetValue("OpType", rowIndex, oRecordSet.Fields.Item("U_opType").Value);
                    oDataTable.SetValue("TransferType", rowIndex, transferType ?? "");
                    oDataTable.SetValue("PaymentID", rowIndex, oRecordSet.Fields.Item("U_paymentID").Value);
                    oDataTable.SetValue("BatchPaymentID", rowIndex, oRecordSet.Fields.Item("U_bPaymentID").Value);
                    oDataTable.SetValue("Status", rowIndex, oRecordSet.Fields.Item("U_status").Value);
                    oDataTable.SetValue("BatchStatus", rowIndex, oRecordSet.Fields.Item("U_bStatus").Value);
                    oDataTable.SetValue("DebitAccount", rowIndex, dataForImport["DebitAccount"] == null ? "" : dataForImport["DebitAccount"]);
                    oDataTable.SetValue("DebitAccountCurrencyCode", rowIndex, dataForImport["DebitAccountCurrencyCode"] == null ? "" : dataForImport["DebitAccountCurrencyCode"]);
                    oDataTable.SetValue("Amount", rowIndex, Convert.ToDouble(dataForImport["Amount"]));
                    //oDataTable.SetValue("Amount", rowIndex, oRecordSet.Fields.Item("Amount").Value);                   
                    oDataTable.SetValue("Currency", rowIndex, dataForImport["Currency"] == null ? "" : dataForImport["Currency"]);
                    oDataTable.SetValue("CardCode", rowIndex, dataForImport["CardCode"] == null ? "" : dataForImport["CardCode"]);
                    oDataTable.SetValue("BeneficiaryName", rowIndex, dataForImport["BeneficiaryName"] == null ? "" : dataForImport["BeneficiaryName"]);
                    oDataTable.SetValue("BeneficiaryTaxCode", rowIndex, dataForImport["BeneficiaryTaxCode"] == null ? "" : dataForImport["BeneficiaryTaxCode"]);
                    oDataTable.SetValue("CreditAccount", rowIndex, dataForImport["CreditAccount"] == null ? "" : dataForImport["CreditAccount"]);
                    oDataTable.SetValue("CreditAccountCurrencyCode", rowIndex, dataForImport["CreditAccountCurrencyCode"] == null ? "" : dataForImport["CreditAccountCurrencyCode"]);
                    oDataTable.SetValue("Description", rowIndex, dataForImport["Description"] == null ? "" : dataForImport["Description"]);
                    oDataTable.SetValue("AdditionalDescription", rowIndex, dataForImport["AdditionalDescription"] == null ? "" : dataForImport["AdditionalDescription"]);
                    oDataTable.SetValue("TreasuryCode", rowIndex, dataForImport["TreasuryCode"] == null ? "" : dataForImport["TreasuryCode"]);
                    //oDataTable.SetValue("PaymentType", rowIndex, oRecordSet.Fields.Item("PaymentType").Value);
                    oDataTable.SetValue("BeneficiaryBankName", rowIndex, dataForImport["BeneficiaryBankName"] == null ? "" : dataForImport["BeneficiaryBankName"]);
                    oDataTable.SetValue("BeneficiaryBankCode", rowIndex, dataForImport["BeneficiaryBankCode"] == null ? "" : dataForImport["BeneficiaryBankCode"]);
                    oDataTable.SetValue("BeneficiaryAddress", rowIndex, dataForImport["BeneficiaryAddress"] == null ? "" : dataForImport["BeneficiaryAddress"]);
                    oDataTable.SetValue("ChargeDetails", rowIndex, dataForImport["ChargeDetails"] == null ? "" : dataForImport["ChargeDetails"]);
                    oDataTable.SetValue("Comments", rowIndex, oRecordSet.Fields.Item("Comments").Value);
                    oDataTable.SetValue("DocumentStatus", rowIndex, oRecordSet.Fields.Item("Status").Value);

                    oRecordSet.MoveNext();
                    rowIndex++;
                }

                SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("importMTR").Specific));
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
        #endregion 

        #region Get Data from Bank System
        public static void setVisibleFormItemsExport(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            SAPbouiCOM.Item oItem = null;
            oForm.Freeze(true);

            try
            {
                string TBCOB = oForm.DataSources.UserDataSources.Item("TBCOB").ValueEx;

                if (TBCOB == "1") //TBC
                {
                    oItem = oForm.Items.Item("periodS2");
                    oItem.Visible = true;
                    oItem = oForm.Items.Item("startDatE2");
                    oItem.Visible = true;
                    oItem = oForm.Items.Item("endDateE2");
                    oItem.Visible = true;
                    oItem = oForm.Items.Item("accountS2");
                    oItem.Visible = true;
                    oItem = oForm.Items.Item("accountE2");
                    oItem.Visible = true;
                    oForm.ActiveItem = "accountE2";
                    oItem = oForm.Items.Item("currencyCB");
                    oItem.Visible = true;
                    oItem = oForm.Items.Item("transTypeS");
                    oItem.Visible = true;
                    oItem = oForm.Items.Item("transTypCB");
                    oItem.Visible = true;
                }
                else if (TBCOB == "2") //BOG
                {
                    oItem = oForm.Items.Item("periodS2");
                    oItem.Visible = true;
                    oItem = oForm.Items.Item("startDatE2");
                    oItem.Visible = true;
                    oItem = oForm.Items.Item("endDateE2");
                    oItem.Visible = true;
                    oItem = oForm.Items.Item("accountS2");
                    oItem.Visible = true;
                    oItem = oForm.Items.Item("accountE2");
                    oItem.Visible = true;
                    oForm.ActiveItem = "accountE2";
                    oItem = oForm.Items.Item("currencyCB");
                    oItem.Visible = true;
                    oItem = oForm.Items.Item("transTypeS");
                    oItem.Visible = true;
                    oItem = oForm.Items.Item("transTypCB");
                    oItem.Visible = true;
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

        public static void matrixColumnSetLinkedObjectTypeExport(SAPbouiCOM.Form oForm, SAPbouiCOM.ItemEvent pVal)
        {
            try
            {
                if (pVal.ColUID == "DocEntry")
                {
                    if (!pVal.BeforeAction)
                    {
                        SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("exportMTR").Specific));

                        SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("exportMTR");
                        string debitCredit = oDataTable.GetValue("DebitCredit", pVal.Row - 1);

                        SAPbouiCOM.Column oColumn;

                        if (debitCredit == "0")
                        {
                            oColumn = oMatrix.Columns.Item(pVal.ColUID);
                            SAPbouiCOM.LinkedButton oLink = oColumn.ExtendedObject;
                            oLink.LinkedObjectType = "46"; //oVendorPayments
                        }
                        else if (debitCredit == "1")
                        {
                            oColumn = oMatrix.Columns.Item(pVal.ColUID);
                            SAPbouiCOM.LinkedButton oLink = oColumn.ExtendedObject;
                            oLink.LinkedObjectType = "24"; //oIncomingPayments
                        }
                    }
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

        public static void comboSelectExport(SAPbouiCOM.Form oForm, SAPbouiCOM.ItemEvent pVal)
        {
            try
            {
                if (!pVal.BeforeAction)
                {
                    if (pVal.ItemUID == "transTypCB")
                    {
                        SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("exportMTR");
                        oDataTable.Rows.Clear();
                        SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("exportMTR").Specific));
                        oMatrix.Clear();
                        oMatrix.LoadFromDataSource();
                        //oMatrix.AutoResizeColumns();

                        oForm.Freeze(false);

                        string TBCOB = oForm.DataSources.UserDataSources.Item("TBCOB").ValueEx;
                        string bankProgram = TBCOB == "1" ? "TBC" : "BOG";

                        if (bankProgram == "TBC" && oBaseQueryResultIoStc != null && oAccountMovementDetailIoStc != null)
                            fillExportMTR_TBC(oForm, oBaseQueryResultIoStc, oAccountMovementDetailIoStc, true);
                        else if (bankProgram == "BOG" && oStatementDetailStc != null)
                            fillExportMTR_BOG(oForm, oStatementDetailStc, true);
                    }
                }
                else
                {
                    if (pVal.ItemUID == "transTypCB")
                    {
                        oForm.Freeze(true);
                    }
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

        public static void checkUncheckMTRExport(SAPbouiCOM.Form oForm, string checkOperation, out string errorText)
        {
            errorText = null;
            try
            {
                SelectAllImportPressed = true;

                oForm.Freeze(true);

                SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("exportMTR").Specific));
                SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("exportMTR");

                oMatrix.FlushToDataSource();

                for (int j = 1; j <= oMatrix.RowCount; j++)
                {
                    /*oCheckBox = oMatrix.Columns.Item("CheckBox").Cells.Item(j).Specific;*/

                    if (j == oMatrix.RowCount)
                    {
                        SelectAllImportPressed = false;
                    }

                    if (checkOperation == "checkB2")
                    {
                        oDataTable.SetValue("CheckBox", j - 1, "Y");
                    }
                    else
                    {
                        oDataTable.SetValue("CheckBox", j - 1, "N");
                    }

                    oMatrix.LoadFromDataSource();

                    OnCheckImportDocuments(oForm, j, "CheckBox");
                    /*oCheckBox.Checked = (checkOperation == "checkB2");*/
                }

                oMatrix.LoadFromDataSource();
                oForm.Update();
                oMatrix.SelectRow(1, true, false);

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

        public static void chooseFromListExport(SAPbouiCOM.Form oForm, SAPbouiCOM.IChooseFromListEvent oCFLEvento, SAPbouiCOM.ItemEvent pVal)
        {
            try
            {
                oForm.Freeze(true);

                string sCFL_ID = oCFLEvento.ChooseFromListUID;
                SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(sCFL_ID);

                if (pVal.BeforeAction == true)
                {
                    if (sCFL_ID == "HouseBankAccount_CFL2" && pVal.ItemUID == "accountE2")
                    {
                        string TBCOB = oForm.DataSources.UserDataSources.Item("TBCOB").ValueEx;
                        string bankProgram = TBCOB == "1" ? "TBC" : "BOG";

                        SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();
                        SAPbouiCOM.Condition oCon = oCons.Add();
                        oCon.Alias = "U_program";
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = bankProgram;
                        oCFL.SetConditions(oCons);
                    }
                    if (sCFL_ID == "GLAccount_CFL")
                    {
                        if (pVal.ItemUID == "exportMTR" && pVal.ColUID == "GLAcctCode")
                        {
                            SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("exportMTR");
                            string transactionType = oDataTable.GetValue("TransactionType", pVal.Row - 1);
                            string partnerAccountNumber = oDataTable.GetValue("PartnerAccountNumber", pVal.Row - 1);

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
                            if ((transactionType == OperationTypeFromIntBank.TransferFromBP.ToString() || transactionType == OperationTypeFromIntBank.TransferToBP.ToString() || transactionType == OperationTypeFromIntBank.TreasuryTransferPaymentOrderIoBP.ToString()) && string.IsNullOrEmpty(partnerAccountNumber) == false)
                                oCon.CondVal = "Y"; //bp
                            else
                                oCon.CondVal = "N";

                            oCFL.SetConditions(oCons);
                        }
                    }
                    else if (sCFL_ID == "CashFlowLineItem_CFL")
                    {

                    }
                    else if (sCFL_ID == "BlnkAgr_CFL")
                    {
                        SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("exportMTR");
                        string partnerAccountNumber = oDataTable.GetValue("PartnerAccountNumber", pVal.Row - 1);
                        string partnerCurrency = oDataTable.GetValue("PartnerCurrency", pVal.Row - 1);
                        string partnerTaxCode = oDataTable.GetValue("PartnerTaxCode", pVal.Row - 1);
                        string transactionType = oDataTable.GetValue("TransactionType", pVal.Row - 1);
                        string cardType = null;
                        if (transactionType == OperationTypeFromIntBank.TransferFromBP.ToString())
                        {
                            cardType = "C";
                        }
                        else if (transactionType == OperationTypeFromIntBank.TransferToBP.ToString() || transactionType == OperationTypeFromIntBank.TreasuryTransferPaymentOrderIoBP.ToString())
                        {
                            cardType = "S";
                        }

                        SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();
                        SAPbouiCOM.Condition oCon = oCons.Add();
                        oCon.Alias = "BpType";
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = cardType;

                        SAPbobsCOM.Recordset oRecordSet = CommonFunctions.getBPBankInfo(partnerAccountNumber + partnerCurrency, partnerTaxCode, cardType);
                        if (oRecordSet != null)
                        {
                            oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                            oCon = oCons.Add();
                            oCon.Alias = "BpCode";
                            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                            oCon.CondVal = oRecordSet.Fields.Item("CardCode").Value;
                        }

                        oCFL.SetConditions(oCons);
                    }
                }
                else if (!pVal.BeforeAction)
                {
                    SAPbouiCOM.DataTable oDataTable = null;
                    oDataTable = oCFLEvento.SelectedObjects;

                    if (oDataTable != null)
                    {
                        if (sCFL_ID == "HouseBankAccount_CFL2" && pVal.ItemUID == "accountE2")
                        {
                            string account = Convert.ToString(oDataTable.GetValue("Account", 0));
                            string currency;
                            CommonFunctions.accountParse(account, out currency);

                            LanguageUtils.IgnoreErrors<string>(() => oForm.Items.Item("accountE2").Specific.Value = account);
                            LanguageUtils.IgnoreErrors<string>(() => oForm.Items.Item("currencyCB").Specific.Select(currency, SAPbouiCOM.BoSearchKey.psk_ByValue));
                        }
                        if (oCFLEvento.ChooseFromListUID == "Budg_CFL")
                        {
                            if (pVal.ItemUID == "exportMTR" && pVal.ColUID == "BCFWId")
                            {
                                string BCFWId = Convert.ToString(oDataTable.GetValue("Code", 0));
                                string BCFWName = Convert.ToString(oDataTable.GetValue("Name", 0));

                                SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("exportMTR").Specific));
                                LanguageUtils.IgnoreErrors<string>(() => oMatrix.Columns.Item("BCFWId").Cells.Item(pVal.Row).Specific.Value = BCFWId);
                                LanguageUtils.IgnoreErrors<string>(() => oMatrix.Columns.Item("BCFWName").Cells.Item(pVal.Row).Specific.Value = BCFWName);
                            }
                        }
                        else if (sCFL_ID == "GLAccount_CFL")
                        {
                            if (pVal.ItemUID == "exportMTR" && pVal.ColUID == "GLAcctCode")
                            {
                                string acctCode = Convert.ToString(oDataTable.GetValue("AcctCode", 0));

                                SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("exportMTR").Specific));
                                LanguageUtils.IgnoreErrors<string>(() => oMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value = acctCode);
                            }
                        }
                        else if (sCFL_ID == "Project_CFLA")
                        {
                            if (pVal.ItemUID == "exportMTR" && pVal.ColUID == "Project")
                            {
                                string prjCode = Convert.ToString(oDataTable.GetValue("PrjCode", 0));

                                SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("exportMTR").Specific));
                                LanguageUtils.IgnoreErrors<string>(() => oMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value = prjCode);
                            }
                        }
                        else if (sCFL_ID == "BlnkAgr_CFL")
                        {
                            if (pVal.ItemUID == "exportMTR" && pVal.ColUID == "BlnkAgr")
                            {
                                string absID = Convert.ToString(oDataTable.GetValue("AbsID", 0));

                                SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("exportMTR").Specific));
                                LanguageUtils.IgnoreErrors<string>(() => oMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value = absID);
                                if (!string.IsNullOrEmpty(absID) && !BlanketAgreement.UsesCurrencyExchangeRates(Convert.ToInt32(absID)))
                                {
                                    SAPbouiCOM.CheckBox oCheckBox = oMatrix.Columns.Item("UseBlaAgRt").Cells.Item(pVal.Row).Specific;
                                    oCheckBox.Checked = false;
                                }
                                setMTRCellEditableSetting(oForm, pVal.ItemUID, pVal.Row);
                            }
                        }
                        else if (sCFL_ID == "CashFlowLineItem_CFL")
                        {
                            if (pVal.ItemUID == "exportMTR" && pVal.ColUID == "CFWId")
                            {
                                string CFWId = Convert.ToString(oDataTable.GetValue("CFWId", 0));
                                string CFWName = Convert.ToString(oDataTable.GetValue("CFWName", 0));

                                SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("exportMTR").Specific));
                                LanguageUtils.IgnoreErrors<string>(() => oMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value = CFWId);
                                LanguageUtils.IgnoreErrors<string>(() => oMatrix.Columns.Item("CFWName").Cells.Item(pVal.Row).Specific.Value = CFWName);
                            }
                        }
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

        private static void setMTRCellEditableSetting(SAPbouiCOM.Form oForm, string mtrName, int rowIndex = 0)
        {
            try
            {
                oForm.Freeze(true);

                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(mtrName).Specific;
                int rowCount = rowIndex == 0 ? oMatrix.RowCount : rowIndex;
                int i = rowIndex == 0 ? 1 : rowIndex;

                for (; i <= rowCount; i++)
                {
                    string absID = oMatrix.GetCellSpecific("BlnkAgr", i).Value;
                    string docEntry = oMatrix.GetCellSpecific("DocEntry", i).Value;
                    if (!string.IsNullOrEmpty(absID) && BlanketAgreement.UsesCurrencyExchangeRates(Convert.ToInt32(absID)))
                    {
                        oMatrix.CommonSetting.SetCellEditable(i, 31, true);
                    }
                    else
                    {
                        oMatrix.CommonSetting.SetCellEditable(i, 31, false);
                    }

                    oMatrix.CommonSetting.SetCellEditable(i, 33, string.IsNullOrEmpty(docEntry));
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                oForm.Update();
                oForm.Freeze(false);
            }
        }

        public static void getPaymentsDocument(string table, string paymentID, string ePaymentID, string docNumber, 
            string transCode, string operationCode, out string docEntry, out string docNum, out string cFWId, out string cFWName)
        {
            docEntry = "";
            docNum = "";
            cFWId = "";
            cFWName = "";
            string transactionCode;

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                string query = @"SELECT
	                  ""PMNT"".""DocEntry"",
	                  ""PMNT"".""DocNum"",
	                  ""PMNT"".""U_paymentID"",
	                  ""PMNT"".""U_ePaymentID"",
	                  ""PMNT"".""U_docNumber"",
	                  ""PMNT"".""U_transCode"",
                      ""PMNT"".""U_opCode"",
                      ""CFW"".""CFWId"",
                      ""CFW"".""CFWName""
                 FROM """ + table + @""" AS ""PMNT"" 
                 join 
	                (SELECT 
		                ""OCFW"".""CFWId"" as ""CFWId"", 
                        ""OCFW"".""CFWName"" as ""CFWName"",
                        ""OJDT"".""CreatedBy"" as ""CreatedBy""
                    FROM ""OJDT""
                    JOIN ""OCFT""  ON ""OJDT"".""TransId"" = ""OCFT"".""JDTId""
                    JOIN ""OCFW"" ON ""OCFT"".""CFWId"" = ""OCFW"".""CFWId""
                    ) as ""CFW""

                on ""CFW"".""CreatedBy"" = ""PMNT"".""DocEntry""

                 WHERE ""PMNT"".""Canceled"" = 'N' AND ""PMNT"".""U_paymentID"" = '" + paymentID + @"'
                 
                 AND ((""PMNT"".""U_ePaymentID"" = '' OR ""PMNT"".""U_ePaymentID"" IS NULL) 
                       OR (""PMNT"".""U_ePaymentID"" = '" + ePaymentID + @"'))
                 AND ((""PMNT"".""U_docNumber"" = '' OR ""PMNT"".""U_docNumber"" IS NULL) 
                       OR (""PMNT"".""U_docNumber"" = '" + docNumber + @"'))
                 AND ((""PMNT"".""U_transCode"" = '' OR ""PMNT"".""U_transCode"" IS NULL) 
                       OR (""PMNT"".""U_transCode"" = '" + transCode + @"'))
                 AND ((""PMNT"".""U_opCode"" = '' OR ""PMNT"".""U_opCode"" IS NULL) 
                       OR (""PMNT"".""U_opCode"" = '" + operationCode + @"'))";

                if (string.IsNullOrEmpty(paymentID))
                {
                    query = query +
                       @" AND ""PMNT"".""U_ePaymentID"" = '" + ePaymentID + @"'  
                       AND ""PMNT"".""U_docNumber"" = '" + docNumber + @"' 
                       AND ""PMNT"".""U_transCode"" = '" + transCode + @"'
                       AND ""PMNT"".""U_opCode"" = '" + operationCode + @"'";
                }

                oRecordSet.DoQuery(query);

                if (!oRecordSet.EoF)
                {
                    transactionCode = oRecordSet.Fields.Item("U_transCode").Value.ToString();
                    if ((operationCode == "*TF%*" || operationCode == "CCO") && string.IsNullOrEmpty(transactionCode))
                    {
                        docEntry = "";
                        docNum = "";
                        cFWId = "";
                        cFWName = "";
                    }
                    else
                    {
                        docEntry = oRecordSet.Fields.Item("DocEntry").Value.ToString();
                        docNum = oRecordSet.Fields.Item("DocNum").Value.ToString();

                        cFWId = oRecordSet.Fields.Item("CFWId").Value.ToString();
                        cFWName = oRecordSet.Fields.Item("CFWName").Value.ToString();
                    }
                }
            }
            catch { }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                oRecordSet = null;
            }
        }

        public static void getPairPaymentsDocument(string table, string paymentID, string documentNumber, string ePaymentID, string opType, out string docEntry)
        {
            docEntry = "";
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            if (string.IsNullOrEmpty(paymentID) && string.IsNullOrEmpty(documentNumber))
            {
                return;
            }

            try
            {
                string query = @"SELECT
	                  ""PMNT"".""DocEntry""
                 FROM """ + table + @""" AS ""PMNT"" 
                 WHERE ""PMNT"".""Canceled"" = 'N' AND ""PMNT"".""U_paymentID"" = '" + paymentID +
                 @"' AND ((""PMNT"".""U_docNumber"" = '' OR ""PMNT"".""U_docNumber"" IS NULL) 
                       OR (""PMNT"".""U_docNumber"" = '" + documentNumber + @"'))
                 AND ""PMNT"".""U_opType"" = '" + opType +
                 @"' AND (""PMNT"".""U_outDoc"" = '' OR ""PMNT"".""U_outDoc"" IS NULL)";

                if (string.IsNullOrEmpty(paymentID) && !string.IsNullOrEmpty(documentNumber) && documentNumber != "0")
                {
                    query = query +
                       @" AND ""PMNT"".""U_docNumber"" = '" + documentNumber + @"'";
                }

                else if (!string.IsNullOrEmpty(ePaymentID))
                {
                    long ePaymentIDInt = Convert.ToInt64(ePaymentID);
                    ePaymentIDInt = table == "OVPM" ? ePaymentIDInt - 1 : ePaymentIDInt + 1;

                    query = query +
                      @" AND ""PMNT"".""U_ePaymentID"" = '" + ePaymentIDInt + @"'";
                }

                oRecordSet.DoQuery(query);

                if (!oRecordSet.EoF)
                {
                    docEntry = oRecordSet.Fields.Item("DocEntry").Value.ToString();
                }
            }
            catch { }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                oRecordSet = null;
            }
        }

        public static void fillExportMTR_TBC(SAPbouiCOM.Form oForm, BaseQueryResultIo oBaseQueryResultIo, AccountMovementDetailIo[] oAccountMovementDetailIo, bool filterByTransType)
        {
            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("exportMTR").Specific));

            RowsWithDifferentDates = new List<int>();
            RowsWithCorrectedDates = new List<int>();

            SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("exportMTR");
            string currency = oForm.DataSources.UserDataSources.Item("currencyCB").ValueEx;
            currency = CommonFunctions.getCurrencyInternationalCode(currency);
            string account = oForm.DataSources.UserDataSources.Item("accountE2").ValueEx;
            string transTypeForFilter = oForm.DataSources.UserDataSources.Item("transTypCB").ValueEx;

            if (oAccountMovementDetailIo.Length > 0 && filterByTransType == false)
            {
                oAccountMovementDetailIoStc = new AccountMovementDetailIo[oAccountMovementDetailIo.Length];
                oAccountMovementDetailIo.CopyTo(oAccountMovementDetailIoStc, 0);
                oBaseQueryResultIoStc = new BaseQueryResultIo();
                oBaseQueryResultIoStc = oBaseQueryResultIo;
            }

            TableExportMTRForDetail.Rows.Clear();
            TableExportMTRForDetail.AcceptChanges();

            oDataTable.Rows.Clear();

            try
            {
                int totalCount = oBaseQueryResultIo.totalCount;
                int size = oAccountMovementDetailIo.Length;
                string[] currencyArray;
                string[] numberArray;
                SAPbobsCOM.Recordset oRecordSet;
                string transactionType;
                string rate;
                string destinationCurrency;
                string sourceCurrency;
                int debitCredit;
                string exchangeRate;
                string treasuryCode;
                string operationCode;
                string partnerAccountNumber;
                string partnerTaxCode;
                string partnerCurrency;
                string GLAccountCodeBP;
                string projectCod;
                string blnkAgr;
                int row = 0;
                string docEntry = "";
                string docNum = "";
                string cFWId = "";
                string cFWName = "";
                string paymentID;
                string ePaymentID;
                string docNumber;
                string transCode;
                string EntryComment;
                string payrollKeyword = CommonFunctions.getOADM("U_BDOSIBKW").ToString();
                DateTime? docDate;

                for (int rowIndex = 0; rowIndex < size; rowIndex++)
                {
                    CommonFunctions.nullsToEmptyString(oAccountMovementDetailIo[rowIndex]);

                    docDate = null;
                    EntryComment = oAccountMovementDetailIo[rowIndex].description;
                    transactionType = oAccountMovementDetailIo[rowIndex].transactionType;
                    rate = "";
                    destinationCurrency = oAccountMovementDetailIo[rowIndex].amount.currency;
                    sourceCurrency = oAccountMovementDetailIo[rowIndex].amount.currency;
                    debitCredit = oAccountMovementDetailIo[rowIndex].debitCredit;
                    exchangeRate = oAccountMovementDetailIo[rowIndex].exchangeRate;
                    treasuryCode = oAccountMovementDetailIo[rowIndex].treasuryCode;
                    operationCode = oAccountMovementDetailIo[rowIndex].operationCode;
                    partnerAccountNumber = CommonFunctions.accountParse(oAccountMovementDetailIo[rowIndex].partnerAccountNumber);
                    partnerTaxCode = oAccountMovementDetailIo[rowIndex].partnerTaxCode;
                    partnerCurrency = "";

                    if (transactionType == "5") //კონვერტაცია
                    {
                        numberArray = CommonFunctions.getNumberArrayFromText(exchangeRate);
                        currencyArray = Regex.Split(exchangeRate, @"[^a-z]+", RegexOptions.IgnoreCase).Where(c => c.Trim() != "").ToArray();

                        if (numberArray.Length > 0 && currencyArray.Length > 0)
                        {
                            rate = numberArray[1];

                            if (debitCredit == 0) //გასვლა
                            {
                                destinationCurrency = oAccountMovementDetailIo[rowIndex].amount.currency; //მიმღები
                                sourceCurrency = destinationCurrency == currencyArray[0] ? currencyArray[1] : currencyArray[0]; //გამგზავნი 
                            }
                            else if (debitCredit == 1) //შემოსვლა
                            {
                                sourceCurrency = oAccountMovementDetailIo[rowIndex].amount.currency; //გამგზავნი 
                                destinationCurrency = sourceCurrency == currencyArray[0] ? currencyArray[1] : currencyArray[0]; //მიმღები                               
                            }
                        }
                    }
                    //else
                    //{
                    if (debitCredit == 1) //შემოსვლა
                        partnerCurrency = destinationCurrency;
                    else if (debitCredit == 0) //გასვლა
                        partnerCurrency = sourceCurrency;
                    //}
                    OperationTypeFromIntBank oOperationTypeFromIntBank = getOperationTypeIntBankTBC(transactionType, operationCode, partnerAccountNumber, partnerCurrency, partnerTaxCode, debitCredit, out GLAccountCodeBP, out projectCod, out blnkAgr);

                    if (oOperationTypeFromIntBank == OperationTypeFromIntBank.TreasuryTransfer)
                    {
                        partnerCurrency = "";
                    }

                    string payrollKeywordTemp = payrollKeyword;
                    while (payrollKeywordTemp.IndexOf(";") > 0 || payrollKeywordTemp != "")
                    {
                        int iSep = payrollKeywordTemp.IndexOf(";");
                        string Sep = payrollKeywordTemp;

                        if (iSep > 0)
                        {
                            Sep = payrollKeywordTemp.Substring(0, iSep);
                            payrollKeywordTemp = payrollKeywordTemp.Substring(iSep + 1);
                        }
                        else
                        {
                            payrollKeywordTemp = "";
                        }

                        if (EntryComment.IndexOf(Sep) > 0)
                        {
                            oOperationTypeFromIntBank = OperationTypeFromIntBank.Salary;
                        }
                    }

                    paymentID = oAccountMovementDetailIo[rowIndex].paymentId;
                    //paymentID = "";
                    ePaymentID = oAccountMovementDetailIo[rowIndex].externalPaymentId;
                    docNumber = oAccountMovementDetailIo[rowIndex].documentNumber;
                    transCode = oAccountMovementDetailIo[rowIndex].transactionType;

                    if (debitCredit == 0) //გასვლა
                        getPaymentsDocument("OVPM", paymentID, ePaymentID, docNumber, transCode, operationCode, out docEntry, out docNum, out cFWId, out cFWName);
                    else
                        getPaymentsDocument("ORCT", paymentID, ePaymentID, docNumber, transCode, operationCode, out docEntry, out docNum, out cFWId, out cFWName);

                    bool deleteFromUnsynchronized = false;
                    //ნაპოვნი დოკუმენტის განახლება --->
                    if (string.IsNullOrEmpty(docEntry) == false)
                    {
                        SAPbobsCOM.Payments oPayments = null;
                        if (debitCredit == 0)
                            oPayments = (SAPbobsCOM.Payments)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments);
                        else
                            oPayments = (SAPbobsCOM.Payments)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments);

                        if (oPayments.GetByKey(Convert.ToInt32(docEntry)))
                        {
                            if (string.IsNullOrEmpty(oPayments.UserFields.Fields.Item("U_docNumber").Value))
                            {
                                oPayments.UserFields.Fields.Item("U_docNumber").Value = docNumber;
                                oPayments.UserFields.Fields.Item("U_transCode").Value = transCode;
                                oPayments.UserFields.Fields.Item("U_ePaymentID").Value = ePaymentID;
                                oPayments.UserFields.Fields.Item("U_opCode").Value = operationCode;

                                int returnCode = oPayments.Update();
                            }

                            docDate = oPayments.DocDate;

                        }

                        Marshal.FinalReleaseComObject(oPayments);
                        oPayments = null;
                    }
                    //ნაპოვნი დოკუმენტის განახლება <---
                    //filterByTransType && 
                    if (transTypeForFilter == OperationTypeFromIntBank.WithoutSalary.ToString())
                    {
                        if (oOperationTypeFromIntBank.ToString() == OperationTypeFromIntBank.Salary.ToString() || oOperationTypeFromIntBank.ToString() == OperationTypeFromIntBank.None.ToString())
                        {
                            if (deleteFromUnsynchronized)
                            {
                                RowsWithDifferentDates.Remove(RowsWithDifferentDates.Last());
                            }

                            continue;
                        }
                    }
                    else if (string.IsNullOrEmpty(transTypeForFilter) == false && transTypeForFilter != OperationTypeFromIntBank.None.ToString() && transTypeForFilter != oOperationTypeFromIntBank.ToString())
                    {
                        if (deleteFromUnsynchronized)
                        {
                            RowsWithDifferentDates.Remove(RowsWithDifferentDates.Last());
                        }
                        continue;
                    }

                    oDataTable.Rows.Add();
                    oDataTable.SetValue("LineNum", row, row + 1);
                    oDataTable.SetValue("CheckBox", row, "N");
                    oDataTable.SetValue("DocEntry", row, docEntry);
                    oDataTable.SetValue("DocNum", row, docNum);
                    oDataTable.SetValue("PaymentID", row, paymentID);
                    oDataTable.SetValue("ExternalPaymentID", row, ePaymentID);
                    oDataTable.SetValue("ValueDate", row, oAccountMovementDetailIo[rowIndex].valueDate);
                    oDataTable.SetValue("DocumentDate", row, oAccountMovementDetailIo[rowIndex].valueDate); //oAccountMovementDetailIo[rowIndex].documentDate
                    oDataTable.SetValue("DebitCredit", row, debitCredit.ToString());
                    oDataTable.SetValue("AccountNumber", row, oAccountMovementDetailIo[rowIndex].accountNumber);
                    oDataTable.SetValue("Currency", row, currency);
                    oDataTable.SetValue("Description", row, oAccountMovementDetailIo[rowIndex].description);
                    oDataTable.SetValue("AdditionalDescription", row, oAccountMovementDetailIo[rowIndex].additionalDescription);
                    oDataTable.SetValue("Amount", row, Convert.ToDouble(oAccountMovementDetailIo[rowIndex].amount.amount));
                    if (oOperationTypeFromIntBank == OperationTypeFromIntBank.CurrencyExchange) //კონვერტაცია
                    {
                        oDataTable.SetValue("CurrencyExchange", row, sourceCurrency);
                        oDataTable.SetValue("Rate", row, rate);
                    }
                    oDataTable.SetValue("PartnerAccountNumber", row, partnerAccountNumber);
                    oDataTable.SetValue("PartnerCurrency", row, partnerCurrency);
                    oDataTable.SetValue("PartnerName", row, oAccountMovementDetailIo[rowIndex].partnerName);
                    oDataTable.SetValue("PartnerTaxCode", row, partnerTaxCode);
                    oDataTable.SetValue("PartnerBankCode", row, oAccountMovementDetailIo[rowIndex].partnerBankCode);
                    oDataTable.SetValue("PartnerBank", row, oAccountMovementDetailIo[rowIndex].partnerBank);
                    oDataTable.SetValue("ChargeDetail", row, oAccountMovementDetailIo[rowIndex].chargeDetail);
                    oDataTable.SetValue("TreasuryCode", row, treasuryCode);
                    oDataTable.SetValue("DocumentNumber", row, docNumber);
                    oDataTable.SetValue("OperationCode", row, operationCode);

                    oDataTable.SetValue("TransactionType", row, oOperationTypeFromIntBank.ToString());
                    oDataTable.SetValue("TransactionCode", row, transCode);

                    oRecordSet = BDOSInternetBankingIntegrationServicesRules.getRules(oOperationTypeFromIntBank, treasuryCode);

                    if (oOperationTypeFromIntBank == OperationTypeFromIntBank.TransferFromBP || oOperationTypeFromIntBank == OperationTypeFromIntBank.TransferToBP || oOperationTypeFromIntBank == OperationTypeFromIntBank.TreasuryTransferPaymentOrderIoBP)
                    {
                        oDataTable.SetValue("GLAccountCode", row, GLAccountCodeBP == null ? "" : GLAccountCodeBP);
                        oDataTable.SetValue("Project", row, projectCod == null ? "" : projectCod);
                        oDataTable.SetValue("BlnkAgr", row, blnkAgr == null ? "" : blnkAgr);
                    }
                    else if (oRecordSet != null)
                    {
                        oDataTable.SetValue("GLAccountCode", row, oRecordSet.Fields.Item("U_AcctCode").Value.ToString() == null ? "" : oRecordSet.Fields.Item("U_AcctCode").Value.ToString());
                    }

                    if (!string.IsNullOrEmpty(cFWId))
                    {
                        oDataTable.SetValue("CashFlowLineItemID", row, cFWId);
                        oDataTable.SetValue("CashFlowLineItemName", row, cFWName);
                    }

                    else if (oRecordSet != null)
                    {
                        oDataTable.SetValue("CashFlowLineItemID", row, oRecordSet.Fields.Item("U_CFWId").Value.ToString() == null ? "" : oRecordSet.Fields.Item("U_CFWId").Value.ToString());
                        oDataTable.SetValue("CashFlowLineItemName", row, oRecordSet.Fields.Item("U_CFWName").Value.ToString() == null ? "" : oRecordSet.Fields.Item("U_CFWName").Value.ToString());
                    }
                    
                    if (CommonFunctions.IsDevelopment())
                    {
                        if (oRecordSet != null)
                        {
                            oDataTable.SetValue("BudgetCashFlowID", row, oRecordSet.Fields.Item("U_BCFWId").Value.ToString() == null ? "" : oRecordSet.Fields.Item("U_BCFWId").Value.ToString());
                            oDataTable.SetValue("BudgetCashFlowName", row, oRecordSet.Fields.Item("U_BCFWName").Value.ToString() == null ? "" : oRecordSet.Fields.Item("U_BCFWName").Value.ToString());
                        }
                    }

                    if (string.IsNullOrEmpty(docEntry) && (oOperationTypeFromIntBank == OperationTypeFromIntBank.TransferFromBP))
                    {
                        bool automaticPaymentInternetBanking = (oForm.DataSources.UserDataSources.Item("autoPay").ValueEx == "Y");

                        if (automaticPaymentInternetBanking == false)
                        {
                            oDataTable.SetValue("PaymentOnAccount", row, Convert.ToDouble(oAccountMovementDetailIo[rowIndex].amount.amount));
                        }
                        else
                        {
                            oDataTable.SetValue("AddDownPaymentAmount", row, Convert.ToDouble(oAccountMovementDetailIo[rowIndex].amount.amount));

                        }

                        oDataTable.SetValue("InDetail", row, "SPI_INFO"); //BPMN_ICON_ACTIVITY_ERROR
                        oDataTable.SetValue("DocRateIN", row, 0);
                    }

                    if (docDate != null && docDate != oDataTable.GetValue("DocumentDate", row))
                    {
                        RowsWithDifferentDates.Add(row + 1);
                        deleteFromUnsynchronized = true;
                    }
                    else
                    {
                        RowsWithCorrectedDates.Add(row + 1);

                        if (RowsWithDifferentDates.Contains(row + 1))
                        {
                            RowsWithDifferentDates.Remove(row + 1);
                        }
                    }

                    row++;
                }

                oForm.Freeze(true);
                oMatrix.Clear();
                oMatrix.LoadFromDataSource();
                //oMatrix.AutoResizeColumns();
                oForm.Update();

                foreach (int rowNum in RowsWithDifferentDates)
                {
                    oMatrix.CommonSetting.SetRowBackColor(rowNum, FormsB1.getLongIntRGB(255, 0, 0));
                }

                if (RowsWithCorrectedDates != null)
                {
                    foreach (int rowNum in RowsWithCorrectedDates)
                    {
                        oMatrix.CommonSetting.SetRowBackColor(rowNum, -1);
                    }
                }

                setMTRCellEditableSetting(oForm, "exportMTR");
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

        public static void fillExportMTR_BOG(SAPbouiCOM.Form oForm, List<StatementDetail> oStatementDetail, bool filterByTransType)
        {
            RowsWithDifferentDates = new List<int>();
            RowsWithCorrectedDates = new List<int>();


            SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("exportMTR");
            string currency = oForm.DataSources.UserDataSources.Item("currencyCB").ValueEx;
            currency = CommonFunctions.getCurrencyInternationalCode(currency);
            string account = oForm.DataSources.UserDataSources.Item("accountE2").ValueEx;
            string transTypeForFilter = oForm.DataSources.UserDataSources.Item("transTypCB").ValueEx;

            if (oStatementDetail.Count > 0 && filterByTransType == false)
            {
                oStatementDetailStc = new List<StatementDetail>();
                oStatementDetailStc = oStatementDetail.ToList();
            }

            TableExportMTRForDetail.Rows.Clear();
            TableExportMTRForDetail.AcceptChanges();

            oDataTable.Rows.Clear();

            try
            {
                int totalCount = oStatementDetail.Count;
                int size = oStatementDetail.Count;
                SAPbobsCOM.Recordset oRecordSet;
                string transactionType;
                decimal? amount;
                decimal? rate;
                DateTime? valueDate;
                DateTime? documentDate;
                decimal amountDc;
                decimal rateDc;
                DateTime valueDateDt;
                DateTime documentDateDt;
                string destinationCurrency;
                string sourceCurrency;
                int debitCredit;
                string treasuryCode;
                string partnerAccountNumber;
                string partnerTaxCode;
                string partnerCurrency;
                string GLAccountCodeBP;
                string projectCod;
                string blnkAgr;
                int row = 0;
                string EntryComment;

                string correspondentAccountNumber;
                string senderAccountNumber;
                string beneficiaryAccountNumber;
                string docEntry = "";
                string docNum = "";
                string cFWId = "";
                string cFWName = "";
                string paymentID;
                string ePaymentID;
                string docNumber;
                string transCode;
                string payrollKeyword = CommonFunctions.getOADM("U_BDOSIBKW").ToString();
                DateTime? docDate;

                for (int rowIndex = 0; rowIndex < size; rowIndex++)
                {
                    CommonFunctions.nullsToEmptyString(oStatementDetail[rowIndex]);
                    CommonFunctions.nullsToEmptyString(oStatementDetail[rowIndex].SenderDetails);
                    CommonFunctions.nullsToEmptyString(oStatementDetail[rowIndex].BeneficiaryDetails);

                    docDate = null;
                    EntryComment = oStatementDetail[rowIndex].EntryComment;
                    transactionType = oStatementDetail[rowIndex].DocumentProductGroup;
                    correspondentAccountNumber = oStatementDetail[rowIndex].DocumentCorrespondentAccountNumber;
                    senderAccountNumber = oStatementDetail[rowIndex].SenderDetails.AccountNumber;
                    beneficiaryAccountNumber = oStatementDetail[rowIndex].BeneficiaryDetails.AccountNumber;
                    destinationCurrency = oStatementDetail[rowIndex].DocumentDestinationCurrency;
                    sourceCurrency = oStatementDetail[rowIndex].DocumentSourceCurrency;
                    treasuryCode = oStatementDetail[rowIndex].DocumentTreasuryCode;

                    debitCredit = -1;

                    if (transactionType == "CCO")
                    {
                        if (correspondentAccountNumber == senderAccountNumber)
                            debitCredit = 1; //შემოსვლა
                        else
                            debitCredit = 0; //გასვლა

                        sourceCurrency = oStatementDetail[rowIndex].DocumentDestinationCurrency;
                        destinationCurrency = oStatementDetail[rowIndex].DocumentSourceCurrency;

                    }
                    else if (transactionType == "COM")
                    {
                        sourceCurrency = currency;
                        destinationCurrency = currency;

                        debitCredit = 0;
                    }
                    else
                    {
                        if (senderAccountNumber == account || account.IndexOf(senderAccountNumber) >= 0)
                            debitCredit = 0; //გასვლა
                        else if (beneficiaryAccountNumber == account || account.IndexOf(beneficiaryAccountNumber) >= 0)
                            debitCredit = 1; //შემოსვლა

                        else
                        {
                            int s = 0;//wtf
                            s++;
                            continue;
                        }
                    }

                    senderAccountNumber = CommonFunctions.accountParse(senderAccountNumber);
                    beneficiaryAccountNumber = CommonFunctions.accountParse(beneficiaryAccountNumber);

                    if(transactionType == "CCO") //swap
                    {
                        var tempAcc = senderAccountNumber;
                        senderAccountNumber = beneficiaryAccountNumber;
                        beneficiaryAccountNumber = tempAcc;
                    }

                    amount = oStatementDetail[rowIndex].EntryAmountDebit;
                    amountDc = 0;
                    if (debitCredit == 1)//(amount == null)
                    {
                        amount = oStatementDetail[rowIndex].EntryAmountCredit;
                    }
                    //if (debitCredit == 1)
                    amountDc = Convert.ToDecimal(amount);

                    rate = oStatementDetail[rowIndex].DocumentRate;
                    rateDc = 0;
                    if (rate != null)
                        rateDc = Convert.ToDecimal(rate);

                    valueDate = oStatementDetail[rowIndex].EntryDate;
                    //valueDateDt = DateTime.Today; //დროებით
                    valueDateDt = new DateTime();
                    if (valueDate != null)
                        valueDateDt = Convert.ToDateTime(valueDate);

                    documentDate = oStatementDetail[rowIndex].DocumentValueDate;
                    //documentDateDt = DateTime.Today; //დროებით
                    documentDateDt = new DateTime();
                    if (documentDate != null)
                        documentDateDt = Convert.ToDateTime(documentDate);

                    if (debitCredit == 0) //გასვლა
                    {
                        partnerAccountNumber = beneficiaryAccountNumber;
                        partnerCurrency = destinationCurrency;
                        partnerTaxCode = oStatementDetail[rowIndex].BeneficiaryDetails.Inn;
                    }
                    else if (debitCredit == 1)//შემოსვლა
                    {
                        partnerAccountNumber = senderAccountNumber;
                        partnerCurrency = sourceCurrency;
                        partnerTaxCode = oStatementDetail[rowIndex].SenderDetails.Inn;
                    }
                    else
                    {
                        partnerAccountNumber = "";
                        partnerCurrency = "";
                        partnerTaxCode = "";
                    }

                    OperationTypeFromIntBank oOperationTypeFromIntBank = getOperationTypeIntBankBOG(transactionType, partnerAccountNumber, partnerCurrency, partnerTaxCode, debitCredit, treasuryCode, out GLAccountCodeBP, out projectCod, out blnkAgr);

                    if (oOperationTypeFromIntBank == OperationTypeFromIntBank.TreasuryTransfer)
                    {
                        partnerCurrency = "";
                    }
                    string payrollKeywordTemp = payrollKeyword;
                    while (payrollKeywordTemp.IndexOf(";") > 0 || payrollKeywordTemp != "")
                    {
                        int iSep = payrollKeywordTemp.IndexOf(";");
                        string Sep = payrollKeywordTemp;

                        if (iSep > 0)
                        {
                            Sep = payrollKeywordTemp.Substring(0, iSep);
                            payrollKeywordTemp = payrollKeywordTemp.Substring(iSep + 1);
                        }
                        else
                        {
                            payrollKeywordTemp = "";
                        }

                        if (EntryComment.IndexOf(Sep) > 0)
                        {
                            oOperationTypeFromIntBank = OperationTypeFromIntBank.Salary;
                        }
                    }

                    paymentID = (oStatementDetail[rowIndex].EntryDocumentNumber == "0") ? oStatementDetail[rowIndex].EntryId : oStatementDetail[rowIndex].EntryDocumentNumber; //EntryDocumentNumber
                    ePaymentID = oStatementDetail[rowIndex].EntryId;
                    docNumber = oStatementDetail[rowIndex].EntryDocumentNumber;
                    transCode = oStatementDetail[rowIndex].DocumentProductGroup;

                    if (debitCredit == 0) //გასვლა
                        getPaymentsDocument("OVPM", paymentID, ePaymentID, docNumber, transCode, transCode, out docEntry, out docNum, out cFWId, out cFWName);
                    else
                        getPaymentsDocument("ORCT", paymentID, ePaymentID, docNumber, transCode, transCode, out docEntry, out docNum, out cFWId, out cFWName);

                    //ნაპოვნი დოკუმენტის განახლება --->
                    bool deleteFromUnsynchronized = false;

                    if (string.IsNullOrEmpty(docEntry) == false)
                    {
                        SAPbobsCOM.Payments oPayments = null;
                        if (debitCredit == 0)
                            oPayments = (SAPbobsCOM.Payments)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments);
                        else
                            oPayments = (SAPbobsCOM.Payments)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments);

                        if (oPayments.GetByKey(Convert.ToInt32(docEntry)))
                        {
                            if (string.IsNullOrEmpty(oPayments.UserFields.Fields.Item("U_docNumber").Value))
                            {
                                oPayments.UserFields.Fields.Item("U_docNumber").Value = docNumber;
                                oPayments.UserFields.Fields.Item("U_transCode").Value = transCode;
                                oPayments.UserFields.Fields.Item("U_ePaymentID").Value = ePaymentID;
                                oPayments.UserFields.Fields.Item("U_opCode").Value = transCode;

                                int returnCode = oPayments.Update();
                            }

                            docDate = oPayments.DocDate;

                        }

                        Marshal.FinalReleaseComObject(oPayments);
                        oPayments = null;
                    }
                    //ნაპოვნი დოკუმენტის განახლება <---
                    //filterByTransType && 
                    if (transTypeForFilter == OperationTypeFromIntBank.WithoutSalary.ToString())
                    {
                        if (oOperationTypeFromIntBank.ToString() == OperationTypeFromIntBank.Salary.ToString() || oOperationTypeFromIntBank.ToString() == OperationTypeFromIntBank.None.ToString())
                        {
                            if (deleteFromUnsynchronized)
                            {
                                RowsWithDifferentDates.Remove(RowsWithDifferentDates.Last());
                            }
                            continue;
                        }
                    }
                    else if (string.IsNullOrEmpty(transTypeForFilter) == false && transTypeForFilter != OperationTypeFromIntBank.None.ToString() && transTypeForFilter != oOperationTypeFromIntBank.ToString())
                    {
                        if (deleteFromUnsynchronized)
                        {
                            RowsWithDifferentDates.Remove(RowsWithDifferentDates.Last());
                        }
                        continue;
                    }
                    string cardType = debitCredit == 0 ? "S" : "C";

                    SAPbobsCOM.Recordset oRecordSet2 = CommonFunctions.getBPBankInfo(partnerAccountNumber + partnerCurrency, partnerTaxCode, cardType);
                    string cardCode = "";
                    if (oRecordSet2 != null)
                    {
                        cardCode = oRecordSet2.Fields.Item("CardCode").Value;
                    }
                    oDataTable.Rows.Add();
                    oDataTable.SetValue("LineNum", row, row + 1);
                    oDataTable.SetValue("CheckBox", row, "N");
                    oDataTable.SetValue("DocEntry", row, docEntry);
                    oDataTable.SetValue("DocNum", row, docNum);
                    oDataTable.SetValue("PaymentID", row, paymentID);
                    oDataTable.SetValue("ExternalPaymentID", row, ePaymentID);
                    oDataTable.SetValue("ValueDate", row, valueDateDt);
                    oDataTable.SetValue("DocumentDate", row, documentDateDt);
                    oDataTable.SetValue("DebitCredit", row, debitCredit.ToString());
                    if (debitCredit == 0) //გასვლა
                    {
                        oDataTable.SetValue("AccountNumber", row, senderAccountNumber);
                        oDataTable.SetValue("Currency", row, (sourceCurrency == "RUR") ? "RUB" : sourceCurrency); //currency
                    }
                    else if (debitCredit == 1)//შემოსვლა
                    {
                        oDataTable.SetValue("AccountNumber", row, beneficiaryAccountNumber);
                        oDataTable.SetValue("Currency", row, (destinationCurrency == "RUR") ? "RUB" : destinationCurrency); //currency
                    }
                    oDataTable.SetValue("BPCode", row, cardCode);
                    oDataTable.SetValue("Description", row, oStatementDetail[rowIndex].DocumentNomination);
                    oDataTable.SetValue("AdditionalDescription", row, EntryComment);
                    oDataTable.SetValue("Amount", row, Convert.ToDouble(amountDc, NumberFormatInfo.InvariantInfo));
                    if (oOperationTypeFromIntBank == OperationTypeFromIntBank.CurrencyExchange) //კონვერტაცია
                    {
                        oDataTable.SetValue("DocumentDate", row, valueDateDt);

                        if (rateDc != 0)
                        {
                            oDataTable.SetValue("CurrencyExchange", row, (destinationCurrency == "RUR") ? "RUB" : destinationCurrency);
                            oDataTable.SetValue("Rate", row, rateDc.ToString(NumberFormatInfo.InvariantInfo));
                        }
                    }
                    if (debitCredit == 0) //გასვლა
                    {
                        oDataTable.SetValue("PartnerAccountNumber", row, partnerAccountNumber);
                        oDataTable.SetValue("PartnerCurrency", row, (partnerCurrency == "RUR") ? "RUB" : partnerCurrency);
                        oDataTable.SetValue("PartnerName", row, CommonFunctions.RemoveSymbols(oStatementDetail[rowIndex].BeneficiaryDetails.Name));
                        oDataTable.SetValue("PartnerTaxCode", row, partnerTaxCode);
                        oDataTable.SetValue("PartnerBankCode", row, oStatementDetail[rowIndex].BeneficiaryDetails.BankCode);
                        oDataTable.SetValue("PartnerBank", row, oStatementDetail[rowIndex].BeneficiaryDetails.BankName);
                    }
                    else if (debitCredit == 1)//შემოსვლა
                    {
                        oDataTable.SetValue("PartnerAccountNumber", row, partnerAccountNumber);
                        oDataTable.SetValue("PartnerCurrency", row, (partnerCurrency == "RUR") ? "RUB" : partnerCurrency);
                        oDataTable.SetValue("PartnerName", row, CommonFunctions.RemoveSymbols(oStatementDetail[rowIndex].SenderDetails.Name));
                        oDataTable.SetValue("PartnerTaxCode", row, partnerTaxCode);
                        oDataTable.SetValue("PartnerBankCode", row, oStatementDetail[rowIndex].SenderDetails.BankCode);
                        oDataTable.SetValue("PartnerBank", row, oStatementDetail[rowIndex].SenderDetails.BankName);
                    }
                    oDataTable.SetValue("ChargeDetail", row, "");
                    oDataTable.SetValue("TreasuryCode", row, treasuryCode);
                    oDataTable.SetValue("DocumentNumber", row, docNumber);
                    oDataTable.SetValue("OperationCode", row, transCode);

                    oDataTable.SetValue("TransactionType", row, oOperationTypeFromIntBank.ToString());
                    oDataTable.SetValue("TransactionCode", row, transCode);

                    oRecordSet = BDOSInternetBankingIntegrationServicesRules.getRules(oOperationTypeFromIntBank, treasuryCode);

                    if (oOperationTypeFromIntBank == OperationTypeFromIntBank.TransferFromBP || oOperationTypeFromIntBank == OperationTypeFromIntBank.TransferToBP || oOperationTypeFromIntBank == OperationTypeFromIntBank.TreasuryTransferPaymentOrderIoBP)
                    {
                        oDataTable.SetValue("GLAccountCode", row, GLAccountCodeBP == null ? "" : GLAccountCodeBP);
                        oDataTable.SetValue("Project", row, projectCod == null ? "" : projectCod);
                        oDataTable.SetValue("BlnkAgr", row, blnkAgr == null ? "" : blnkAgr);
                    }
                    else if (oRecordSet != null)
                    {
                        oDataTable.SetValue("GLAccountCode", row, oRecordSet.Fields.Item("U_AcctCode").Value.ToString() == null ? "" : oRecordSet.Fields.Item("U_AcctCode").Value.ToString());
                    }

                    if (!string.IsNullOrEmpty(cFWId))
                    {
                        oDataTable.SetValue("CashFlowLineItemID", row, cFWId);
                        oDataTable.SetValue("CashFlowLineItemName", row, cFWName);
                    }

                    else if (oRecordSet != null)
                    {
                        oDataTable.SetValue("CashFlowLineItemID", row, oRecordSet.Fields.Item("U_CFWId").Value.ToString() == null ? "" : oRecordSet.Fields.Item("U_CFWId").Value.ToString());
                        oDataTable.SetValue("CashFlowLineItemName", row, oRecordSet.Fields.Item("U_CFWName").Value.ToString() == null ? "" : oRecordSet.Fields.Item("U_CFWName").Value.ToString());
                    }

                    if (CommonFunctions.IsDevelopment())
                    {
                        if (oRecordSet != null)
                        {
                            oDataTable.SetValue("BudgetCashFlowID", row, oRecordSet.Fields.Item("U_BCFWId").Value.ToString() == null ? "" : oRecordSet.Fields.Item("U_BCFWId").Value.ToString());
                            oDataTable.SetValue("BudgetCashFlowName", row, oRecordSet.Fields.Item("U_BCFWName").Value.ToString() == null ? "" : oRecordSet.Fields.Item("U_BCFWName").Value.ToString());
                        }
                    }

                    if (string.IsNullOrEmpty(docEntry) && (oOperationTypeFromIntBank == OperationTypeFromIntBank.TransferFromBP))
                    {

                        //
                        bool automaticPaymentInternetBanking = (oForm.DataSources.UserDataSources.Item("autoPay").ValueEx == "Y");

                        if (automaticPaymentInternetBanking == false)
                        {
                            oDataTable.SetValue("PaymentOnAccount", row, Convert.ToDouble(amountDc, NumberFormatInfo.InvariantInfo));
                        }
                        else
                        {
                            oDataTable.SetValue("AddDownPaymentAmount", row, Convert.ToDouble(amountDc, NumberFormatInfo.InvariantInfo));
                        }
                        //

                        oDataTable.SetValue("InDetail", row, "SPI_INFO");
                        oDataTable.SetValue("DocRateIN", row, 0);
                    }

                    if (docDate != null && docDate != oDataTable.GetValue("DocumentDate", row))
                    {
                        RowsWithDifferentDates.Add(row + 1);
                        deleteFromUnsynchronized = true;
                    }
                    else
                    {
                        RowsWithCorrectedDates.Add(row + 1);

                        if (RowsWithDifferentDates.Contains(row + 1))
                        {
                            RowsWithDifferentDates.Remove(row + 1);
                        }
                    }

                    row++;
                }

                SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("exportMTR").Specific));
                oForm.Freeze(true);
                oMatrix.Clear();
                oMatrix.LoadFromDataSource();
                //oMatrix.AutoResizeColumns();
                oForm.Update();

                foreach (int rowNum in RowsWithDifferentDates)
                {
                    oMatrix.CommonSetting.SetRowBackColor(rowNum, FormsB1.getLongIntRGB(255, 0, 0));
                }

                if (RowsWithCorrectedDates != null)
                {
                    foreach (int rowNum in RowsWithCorrectedDates)
                    {
                        oMatrix.CommonSetting.SetRowBackColor(rowNum, -1);
                    }
                }

                setMTRCellEditableSetting(oForm, "exportMTR");
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

        private static OperationTypeFromIntBank getOperationTypeIntBankTBC(string transactionType, string operationCode, string partnerAccountNumber, string partnerCurrency, string partnerTaxCode, int debitCredit, out string GLAccountCodeBP, out string projectCod, out string blnkAgr)
        {
            OperationTypeFromIntBank oOperationType;
            SAPbobsCOM.Recordset oRecordSet;
            string cardType = debitCredit == 0 ? "S" : "C";
            GLAccountCodeBP = null;
            projectCod = null;
            blnkAgr = null;

            if (transactionType == "1") //Transfer between own accounts
                oOperationType = OperationTypeFromIntBank.TransferToOwnAccount;
            else if (transactionType == "5") //Currency exchange
                oOperationType = OperationTypeFromIntBank.CurrencyExchange;
            else if (CommonFunctions.isAccountInHouseBankAccount(partnerAccountNumber + partnerCurrency) == true && (transactionType != "20" || transactionType == "30"))
            {
                oOperationType = OperationTypeFromIntBank.TransferToOwnAccount;
            }
            else if (transactionType == "20" && string.IsNullOrEmpty(partnerAccountNumber) == false && string.IsNullOrEmpty(partnerCurrency) == false) //Income
            {
                oOperationType = OperationTypeFromIntBank.TransferFromBP;
                oRecordSet = CommonFunctions.getBPBankInfo(partnerAccountNumber + partnerCurrency, partnerTaxCode, cardType);
                if (oRecordSet != null)
                {
                    GLAccountCodeBP = oRecordSet.Fields.Item("DebPayAcct").Value;
                    projectCod = oRecordSet.Fields.Item("ProjectCod").Value;
                    blnkAgr = oRecordSet.Fields.Item("BlnkAgr").Value;
                }
                else if (debitCredit == 0)
                {
                    oRecordSet = CommonFunctions.getEmployeeInfo(partnerTaxCode);

                    if (oRecordSet != null)
                    {
                        oOperationType = OperationTypeFromIntBank.None;
                    }
                }
            }
            else if (transactionType == "20") //Income
                oOperationType = OperationTypeFromIntBank.OtherIncomes;
            else if ((transactionType == "30" || transactionType == "32") && string.IsNullOrEmpty(partnerAccountNumber) == false && string.IsNullOrEmpty(partnerCurrency) == false) //Transfer out and cash withdrawal or Treasury transfers
            {
                oOperationType = OperationTypeFromIntBank.TransferToBP;
                oRecordSet = CommonFunctions.getBPBankInfo(partnerAccountNumber + partnerCurrency, partnerTaxCode, cardType);
                if (oRecordSet != null)
                {
                    GLAccountCodeBP = oRecordSet.Fields.Item("DebPayAcct").Value;
                    projectCod = oRecordSet.Fields.Item("ProjectCod").Value;
                    blnkAgr = oRecordSet.Fields.Item("BlnkAgr").Value;
                    if (oRecordSet.Fields.Item("U_treasury").Value == "Y")
                    {
                        oOperationType = OperationTypeFromIntBank.TreasuryTransferPaymentOrderIoBP;
                    }
                    else if (transactionType == "32")
                        oOperationType = OperationTypeFromIntBank.TreasuryTransfer;
                }
                else if (debitCredit == 0 && transactionType == "30")
                {
                    oRecordSet = CommonFunctions.getEmployeeInfo(partnerTaxCode);

                    if (oRecordSet != null)
                    {
                        oOperationType = OperationTypeFromIntBank.None;
                    }
                }
                else if (transactionType == "32")
                    oOperationType = OperationTypeFromIntBank.TreasuryTransfer;
            }
            else if (transactionType == "30") //Transfer out and cash withdrawal
                oOperationType = OperationTypeFromIntBank.OtherExpenses;
            else if (transactionType == "31") //Bill, mobile phone, fine payments
                oOperationType = OperationTypeFromIntBank.OtherExpenses;
            else if (transactionType == "32") //Treasury transfers
                oOperationType = OperationTypeFromIntBank.TreasuryTransfer;
            else if (transactionType == "33" && operationCode == "*TF%*")
                oOperationType = OperationTypeFromIntBank.BankCharge;
            else if (debitCredit == 0)
                oOperationType = OperationTypeFromIntBank.OtherExpenses;
            else if (debitCredit == 1)
                oOperationType = OperationTypeFromIntBank.OtherIncomes;
            else
                oOperationType = OperationTypeFromIntBank.None;

            return oOperationType;
        }

        private static OperationTypeFromIntBank getOperationTypeIntBankBOG(string transactionType, string partnerAccountNumber, string partnerCurrency, string partnerTaxCode, int debitCredit, string treasuryCode, out string GLAccountCodeBP, out string projectCod, out string blnkAgr)
        {
            try
            {
                OperationTypeFromIntBank oOperationType;
                SAPbobsCOM.Recordset oRecordSet;
                string cardType = debitCredit == 0 ? "S" : "C";
                GLAccountCodeBP = null;
                projectCod = null;
                blnkAgr = null;

                if (transactionType == "COM" || transactionType == "FEE")
                    oOperationType = OperationTypeFromIntBank.BankCharge;
                else if (CommonFunctions.isAccountInHouseBankAccount(partnerAccountNumber + partnerCurrency) == true && transactionType != "CCO")
                    oOperationType = OperationTypeFromIntBank.TransferToOwnAccount;
                else if (string.IsNullOrEmpty(treasuryCode) == false)
                {
                    oOperationType = OperationTypeFromIntBank.TreasuryTransfer;
                    if (debitCredit == 0)
                    {
                        oRecordSet = CommonFunctions.getBPBankInfo(treasuryCode + "GEL", partnerTaxCode, cardType);

                        if (oRecordSet != null && oRecordSet.Fields.Item("U_treasury").Value == "Y")
                        {
                            GLAccountCodeBP = oRecordSet.Fields.Item("DebPayAcct").Value;
                            projectCod = oRecordSet.Fields.Item("ProjectCod").Value;
                            blnkAgr = oRecordSet.Fields.Item("BlnkAgr").Value;
                            oOperationType = OperationTypeFromIntBank.TreasuryTransferPaymentOrderIoBP;
                        }
                    }
                }
                else if (transactionType == "CCO")
                    oOperationType = OperationTypeFromIntBank.CurrencyExchange;
                else if (transactionType == "LFG")
                {
                    if (debitCredit == 0)
                        oOperationType = OperationTypeFromIntBank.OtherExpenses;
                    else if (debitCredit == 1)
                        oOperationType = OperationTypeFromIntBank.OtherIncomes;
                    else
                        oOperationType = OperationTypeFromIntBank.None;
                }
                else if (transactionType == "PMC" || transactionType == "LND")
                {
                    if (debitCredit == 0)
                        oOperationType = OperationTypeFromIntBank.OtherExpenses;
                    else if (debitCredit == 1)
                        oOperationType = OperationTypeFromIntBank.OtherIncomes;
                    else
                        oOperationType = OperationTypeFromIntBank.None;
                }
                else if (string.IsNullOrEmpty(partnerAccountNumber) == false && string.IsNullOrEmpty(partnerCurrency) == false)
                {
                    if (debitCredit == 0)
                        oOperationType = OperationTypeFromIntBank.TransferToBP;
                    else
                        oOperationType = OperationTypeFromIntBank.TransferFromBP;

                    oRecordSet = CommonFunctions.getBPBankInfo(partnerAccountNumber + partnerCurrency, partnerTaxCode, cardType);

                    if (oRecordSet != null)
                    {
                        GLAccountCodeBP = oRecordSet.Fields.Item("DebPayAcct").Value;
                        projectCod = oRecordSet.Fields.Item("ProjectCod").Value;
                        blnkAgr = oRecordSet.Fields.Item("BlnkAgr").Value;
                        if (oRecordSet.Fields.Item("U_treasury").Value == "Y" && oOperationType == OperationTypeFromIntBank.TransferToBP)
                        {
                            oOperationType = OperationTypeFromIntBank.TreasuryTransferPaymentOrderIoBP;
                        }
                    }
                    else if (debitCredit == 0)
                    {
                        oRecordSet = CommonFunctions.getEmployeeInfo(partnerTaxCode);

                        if (oRecordSet != null)
                        {
                            oOperationType = OperationTypeFromIntBank.None;
                        }
                    }
                }
                else if (debitCredit == 0)
                    oOperationType = OperationTypeFromIntBank.OtherExpenses;
                else if (debitCredit == 1)
                    oOperationType = OperationTypeFromIntBank.OtherIncomes;
                else
                    oOperationType = OperationTypeFromIntBank.None;

                return oOperationType;
            }
            catch (Exception ex)
            {
                System.Diagnostics.StackTrace st = new System.Diagnostics.StackTrace(ex, true);
                //Get the first stack frame
                System.Diagnostics.StackFrame frame = st.GetFrame(0);

                //Get the file name
                string fileName = frame.GetFileName();

                //Get the method name
                string methodName = frame.GetMethod().Name;

                //Get the line number from the stack frame
                int line = frame.GetFileLineNumber();

                //Get the column number
                int col = frame.GetFileColumnNumber();

                throw new Exception(ex.Message);
            }
        }

        public static void getData(SAPbouiCOM.Form oForm)
        {
            string errorText = null;

            string selectedOperation = "getData";
            string TBCOB = oForm.DataSources.UserDataSources.Item("TBCOB").ValueEx;

            string bankProgram = TBCOB == "1" ? "TBC" : "BOG";
            List<int> docEntryList = new List<int>();

            SAPbobsCOM.SBObob oSBOBob = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);
            string account = oForm.DataSources.UserDataSources.Item("accountE2").ValueEx;
            string currency = oForm.DataSources.UserDataSources.Item("currencyCB").ValueEx;

            SAPbobsCOM.IRecordset oIRecordset = oSBOBob.Format_StringToDate(oForm.DataSources.UserDataSources.Item("startDatE2").ValueEx);
            DateTime startDate = oIRecordset.Fields.Item("Date").Value;
            oIRecordset = oSBOBob.Format_StringToDate(oForm.DataSources.UserDataSources.Item("endDateE2").ValueEx);
            DateTime endDate = oIRecordset.Fields.Item("Date").Value;

            SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("exportMTR");
            oDataTable.Rows.Clear();
            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("exportMTR").Specific));
            oForm.Freeze(true);
            oMatrix.Clear();
            oMatrix.LoadFromDataSource();
            //oMatrix.AutoResizeColumns();
            oForm.Update();
            oForm.Freeze(false);

            if (bankProgram == "TBC")
            {
                AccountMovementFilterIo oAccountMovementFilterIo = new AccountMovementFilterIo();
                if (string.IsNullOrEmpty(account) == false && string.IsNullOrEmpty(currency) == false)
                {
                    string currencyTmp;
                    currency = CommonFunctions.getCurrencyInternationalCode(currency);
                    oAccountMovementFilterIo.accountNumber = CommonFunctions.accountParse(account, out currencyTmp);
                    oAccountMovementFilterIo.accountCurrencyCode = currency;
                    if (currency != currencyTmp)
                    {
                        errorText = BDOSResources.getTranslate("CurrencyAndTheAccountSCurrencyIsDifferent") + "!";
                        Program.uiApp.SetStatusBarMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                        return;
                    }
                }
                else
                {
                    errorText = BDOSResources.getTranslate("TheFollowingFieldsAreMandatory") + " : \"" + oForm.Items.Item("accountS2").Specific.caption + "\", \"" + BDOSResources.getTranslate("Currency") + "\""; //აუცილებელია შემდეგი ველების შევსება
                    Program.uiApp.SetStatusBarMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    return;
                }
                if (startDate != new DateTime())
                {
                    oAccountMovementFilterIo.periodFrom = startDate;
                    oAccountMovementFilterIo.periodFromSpecified = true;
                }
                if (endDate != new DateTime())
                {
                    oAccountMovementFilterIo.periodTo = endDate;
                    oAccountMovementFilterIo.periodToSpecified = true;
                }
                BDOSAuthenticationFormTBC.createForm(oForm, selectedOperation, docEntryList, false, null, oAccountMovementFilterIo, out errorText);
            }
            else if (bankProgram == "BOG")
            {
                StatementFilter oStatementFilter = new StatementFilter();

                if (string.IsNullOrEmpty(account) == false && string.IsNullOrEmpty(currency) == false)
                {
                    string currencyTmp;
                    currency = CommonFunctions.getCurrencyInternationalCode(currency);
                    currency = (currency == "RUB") ? "RUR" : currency;

                    oStatementFilter.AccountNumber = CommonFunctions.accountParse(account, out currencyTmp);
                    oStatementFilter.Currency = currency;
                    currencyTmp = (currencyTmp == "RUB") ? "RUR" : currencyTmp;
                    if (currency != currencyTmp)
                    {
                        errorText = BDOSResources.getTranslate("CurrencyAndTheAccountSCurrencyIsDifferent") + "!";
                        Program.uiApp.SetStatusBarMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                        return;
                    }
                }
                else
                {
                    errorText = BDOSResources.getTranslate("TheFollowingFieldsAreMandatory") + " : \"" + oForm.Items.Item("accountS2").Specific.caption + "\", \"" + BDOSResources.getTranslate("Currency") + "\""; //აუცილებელია შემდეგი ველების შევსება
                    Program.uiApp.SetStatusBarMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    return;
                }

                if (startDate != new DateTime())
                {
                    oStatementFilter.PeriodFrom = startDate;
                }
                if (endDate != new DateTime())
                {
                    oStatementFilter.PeriodTo = endDate;
                }
                BDOSAuthenticationFormBOG.createForm(oForm, selectedOperation, docEntryList, false, null, oStatementFilter, out errorText);
            }
        }

        public static List<string> createDocuments(SAPbouiCOM.Form oForm)
        {
            string info = null;
            List<string> infoList = new List<string>();
            string errorText;
            string docEntryStr;
            int docEntry;
            int docNum;
            string checkBox;
            int debitCredit;
            string transactionType;
            SAPbobsCOM.SBObob oSBOBob = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);

            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("exportMTR").Specific));
            oMatrix.FlushToDataSource();

            bool automaticPaymentInternetBanking = (oForm.DataSources.UserDataSources.Item("autoPay").ValueEx == "Y");

            SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("exportMTR");

            for (int i = 0; i < oDataTable.Rows.Count; i++)
            {
                checkBox = oDataTable.GetValue("CheckBox", i);
                docEntryStr = oDataTable.GetValue("DocEntry", i);
                debitCredit = Convert.ToInt32(oDataTable.GetValue("DebitCredit", i));
                transactionType = oDataTable.GetValue("TransactionType", i);

                if (checkBox == "Y" && string.IsNullOrEmpty(docEntryStr))
                {
                    if (transactionType == OperationTypeFromIntBank.None.ToString() || string.IsNullOrEmpty(transactionType))
                    {
                        continue;
                    }
                    else if (transactionType == OperationTypeFromIntBank.BankCharge.ToString())
                    {
                        info = OutgoingPayment.createDocumentOtherExpensesType(oDataTable, i, out docEntry, out docNum, out errorText);
                        infoList.Add(errorText == null ? info : errorText);
                    }
                    else if (transactionType == OperationTypeFromIntBank.CurrencyExchange.ToString())
                    {
                        if (debitCredit == 0) //გასვლა
                        {
                            info = OutgoingPayment.createDocumentCurrencyExchangeType(oDataTable, i, out docEntry, out docNum, out errorText);
                            infoList.Add(errorText == null ? info : errorText);
                        }
                        else if (debitCredit == 1) //შემოსვლა
                        {
                            info = IncomingPayment.createDocumentCurrencyExchangeType(oDataTable, i, out docEntry, out docNum, out errorText);
                            infoList.Add(errorText == null ? info : errorText);
                        }
                    }
                    else if (transactionType == OperationTypeFromIntBank.OtherExpenses.ToString())
                    {
                        info = OutgoingPayment.createDocumentOtherExpensesType(oDataTable, i, out docEntry, out docNum, out errorText);
                        infoList.Add(errorText == null ? info : errorText);
                    }
                    else if (transactionType == OperationTypeFromIntBank.OtherIncomes.ToString())
                    {
                        info = IncomingPayment.createDocumentOtherIncomesType(oDataTable, i, out docEntry, out docNum, out errorText);
                        infoList.Add(errorText == null ? info : errorText);
                    }
                    else if (transactionType == OperationTypeFromIntBank.TransferFromBP.ToString())
                    {
                        /*if (automaticPaymentInternetBanking)
                        {
                            info = ARDownPaymentRequest.createDocumentTransferFromBPType( oDataTable, oForm, i, out docEntry, out docNum, out errorText);
                        }*/

                        info = IncomingPayment.createDocumentTransferFromBPType(oDataTable, oForm, i, out docEntry, out docNum, out errorText);

                        infoList.Add(String.IsNullOrEmpty(errorText) ? info : errorText);
                    }
                    else if (transactionType == OperationTypeFromIntBank.TransferToBP.ToString() || transactionType == OperationTypeFromIntBank.TreasuryTransferPaymentOrderIoBP.ToString())
                    {
                        info = OutgoingPayment.createDocumentTransferToBPType(oDataTable, i, out docEntry, out docNum, out errorText, transactionType);
                        infoList.Add(errorText == null ? info : errorText);
                    }
                    else if (transactionType == OperationTypeFromIntBank.TransferToOwnAccount.ToString())
                    {
                        if (debitCredit == 0) //გასვლა
                        {
                            info = OutgoingPayment.createDocumentTransferToOwnAccountType(oDataTable, i, out docEntry, out docNum, out errorText);
                            infoList.Add(errorText == null ? info : errorText);
                        }
                        else if (debitCredit == 1) //შემოსვლა
                        {
                            info = IncomingPayment.createDocumentTransferToOwnAccountType(oDataTable, i, out docEntry, out docNum, out errorText);
                            infoList.Add(errorText == null ? info : errorText);
                        }
                    }
                    else if (transactionType == OperationTypeFromIntBank.TreasuryTransfer.ToString())
                    {
                        info = OutgoingPayment.createDocumentTreasuryTransferType(oDataTable, i, out docEntry, out docNum, out errorText);
                        infoList.Add(errorText == null ? info : errorText);
                    }
                }
            }

            oForm.Freeze(true);
            oMatrix.Clear();
            oMatrix.LoadFromDataSource();
            //oMatrix.AutoResizeColumns();
            oForm.Update();
            oForm.Freeze(false);

            return infoList;
        }

        public static void updateExportMTRRow(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("exportMTR");

            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("exportMTR").Specific));
            oMatrix.FlushToDataSource();

            oDataTable.SetValue("DownPaymentAmount", CurrentRowExportMTRForDetail - 1, Convert.ToDouble(BDOSInternetBankingDocuments.downPaymentAmount));
            oDataTable.SetValue("InvoicesAmount", CurrentRowExportMTRForDetail - 1, Convert.ToDouble(BDOSInternetBankingDocuments.invoicesAmount));
            oDataTable.SetValue("PaymentOnAccount", CurrentRowExportMTRForDetail - 1, Convert.ToDouble(BDOSInternetBankingDocuments.paymentOnAccount));
            oDataTable.SetValue("DocRateIN", CurrentRowExportMTRForDetail - 1, Convert.ToDouble(BDOSInternetBankingDocuments.docRateIN));
            oDataTable.SetValue("AddDownPaymentAmount", CurrentRowExportMTRForDetail - 1, Convert.ToDouble(BDOSInternetBankingDocuments.addDownPaymentAmount));

            oForm.Freeze(true);
            oMatrix.LoadFromDataSource();
            oForm.Update();
            oForm.Freeze(false);
        }
        #endregion

        public static DataTable create_TableExportMTRForDetail()
        {
            TableExportMTRForDetail = new DataTable();
            TableExportMTRForDetail.Columns.Add("LineNumExportMTR", typeof(Int32));
            TableExportMTRForDetail.Columns.Add("LineNum", typeof(Int32));
            TableExportMTRForDetail.Columns.Add("CheckBox", typeof(string));
            TableExportMTRForDetail.Columns.Add("DocEntry", typeof(Int32));
            TableExportMTRForDetail.Columns.Add("DocNum", typeof(Int32));
            TableExportMTRForDetail.Columns.Add("InstallmentID", typeof(Int32));
            TableExportMTRForDetail.Columns.Add("LineID", typeof(Int32));
            TableExportMTRForDetail.Columns.Add("DocType", typeof(string));
            TableExportMTRForDetail.Columns.Add("DocDate", typeof(DateTime));
            TableExportMTRForDetail.Columns.Add("DueDate", typeof(DateTime));
            TableExportMTRForDetail.Columns.Add("Arrears", typeof(string));
            TableExportMTRForDetail.Columns.Add("OverdueDays", typeof(Int32));
            TableExportMTRForDetail.Columns.Add("Comments", typeof(string));
            TableExportMTRForDetail.Columns.Add("Total", typeof(decimal));
            TableExportMTRForDetail.Columns.Add("BalanceDue", typeof(decimal));
            TableExportMTRForDetail.Columns.Add("TotalPayment", typeof(decimal));
            TableExportMTRForDetail.Columns.Add("Currency", typeof(string));
            TableExportMTRForDetail.Columns.Add("TotalPaymentLocal", typeof(decimal));

            return TableExportMTRForDetail;
        }

        public static void createForm(out string errorText)
        {
            errorText = null;

            TableExportMTRForDetail = create_TableExportMTRForDetail();

            int formHeight = Program.uiApp.Desktop.Width;
            int formWidth = Program.uiApp.Desktop.Width;

            //ფორმის აუცილებელი თვისებები
            Dictionary<string, object> formProperties = new Dictionary<string, object>();
            formProperties.Add("UniqueID", "BDOSInternetBankingForm");
            formProperties.Add("BorderStyle", SAPbouiCOM.BoFormBorderStyle.fbs_Sizable);
            formProperties.Add("Title", BDOSResources.getTranslate("InternetBanking"));
            //formProperties.Add("Left", (Program.uiApp.Desktop.Width - formWidth) / 2);
            formProperties.Add("ClientWidth", formWidth);
            //formProperties.Add("Top", (Program.uiApp.Desktop.Height - formHeight) / 3);
            formProperties.Add("ClientHeight", formHeight);

            SAPbouiCOM.Form oForm;
            bool newForm;
            bool formExist = FormsB1.createForm(formProperties, out oForm, out newForm, out errorText);

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

                    int left_s = 6;
                    int left_e = 120;
                    int height = 15;
                    int top = 25;
                    top = top + 15;
                    int width_s = 121 - 15;
                    int width_e = 148;

                    //-----------------------------Import Data into Bank System----------------------------->

                    formItems = new Dictionary<string, object>();
                    itemName = "periodS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("Period"));
                    formItems.Add("LinkTo", "startDateE");
                    formItems.Add("FromPane", 1);
                    formItems.Add("ToPane", 1);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    DateTime startDate = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);
                    string startDateTxt = startDate.ToString("yyyyMMdd");

                    DateTime endDate = DateTime.Today;
                    string endDateTxt = endDate.ToString("yyyyMMdd");

                    formItems = new Dictionary<string, object>();
                    itemName = "startDateE";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_DATE);
                    formItems.Add("Length", 1);
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Left", left_e);
                    formItems.Add("Width", width_e / 2);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("ValueEx", startDateTxt);
                    formItems.Add("FromPane", 1);
                    formItems.Add("ToPane", 1);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "endDateE";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_DATE);
                    formItems.Add("Length", 1);
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Left", left_e + width_e / 2);
                    formItems.Add("Width", width_e / 2);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("ValueEx", endDateTxt);
                    formItems.Add("FromPane", 1);
                    formItems.Add("ToPane", 1);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }
                    //-----------------------------------------------

                    formItems = new Dictionary<string, object>();
                    itemName = "rprtCodeS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_e + width_e + 5);
                    formItems.Add("Width", width_s / 1.8);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", "Report Code");
                    formItems.Add("LinkTo", "rprtCodeCB");
                    formItems.Add("FromPane", 1);
                    formItems.Add("ToPane", 1);
                    formItems.Add("Visible", false);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    Dictionary<string, string> listValidValuesRprtCode = new Dictionary<string, string>();
                    listValidValuesRprtCode.Add("GDS", "GDS");
                    listValidValuesRprtCode.Add("ACM", "AGENCY COMMISSION");
                    listValidValuesRprtCode.Add("DCM", "TRADE COMMISION");
                    listValidValuesRprtCode.Add("AKA", "BONUS");


                    formItems = new Dictionary<string, object>();
                    itemName = "rprtCodeCB"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("Length", 30);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                    formItems.Add("Left", left_e + width_e + 120);
                    formItems.Add("Width", width_e);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
                    formItems.Add("ValidValues", listValidValuesRprtCode);
                    formItems.Add("FromPane", 1);
                    formItems.Add("ToPane", 1);
                    formItems.Add("Visible", false);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }





                    //-----------------------------------------------
                    top = top + height + 1;

                    bool multiSelection = false;
                    string objectType = "231"; // HouseBankAccounts object
                    string uniqueID_lf_HouseBankAccountCFL = "HouseBankAccount_CFL";
                    FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_lf_HouseBankAccountCFL);

                    formItems = new Dictionary<string, object>();
                    itemName = "accountS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("Account"));
                    formItems.Add("LinkTo", "accountE");
                    formItems.Add("FromPane", 1);
                    formItems.Add("ToPane", 1);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "accountE"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("Length", 30);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
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
                    formItems.Add("FromPane", 1);
                    formItems.Add("ToPane", 1);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    top = top + height + 1;

                    formItems = new Dictionary<string, object>();
                    itemName = "allDocsOB"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("Length", 1);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_OPTION_BUTTON);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("AllDocuments"));
                    formItems.Add("ValOn", "Y");
                    formItems.Add("ValOff", "N");
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("Selected", true);
                    formItems.Add("FromPane", 1);
                    formItems.Add("ToPane", 1);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "uplDocsOB"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("Length", 1);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_OPTION_BUTTON);
                    formItems.Add("Left", left_s + width_s);
                    formItems.Add("Width", width_e);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("OnlyUploadDocuments"));
                    formItems.Add("GroupWith", "allDocsOB");
                    formItems.Add("ValOn", "Y");
                    formItems.Add("ValOff", "N");
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("FromPane", 1);
                    formItems.Add("ToPane", 1);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    top = top + height + 10;

                    formItems = new Dictionary<string, object>();
                    itemName = "imptDataS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", width_s * 2);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("DataForExport"));
                    formItems.Add("TextStyle", 4);
                    formItems.Add("FromPane", 1);
                    formItems.Add("ToPane", 1);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    top = top + height + 5;

                    formItems = new Dictionary<string, object>();
                    itemName = "imptTypeS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s + 5);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("ExportTypeTBC"));
                    formItems.Add("LinkTo", "imptTypeCB");
                    formItems.Add("FromPane", 1);
                    formItems.Add("ToPane", 1);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    Dictionary<string, string> listValidValuesDict = new Dictionary<string, string>();
                    listValidValuesDict.Add("singlePayment", BDOSResources.getTranslate("SinglePayment")); //ინდივიდუალური გადარიცხვა
                    listValidValuesDict.Add("batchPayment", BDOSResources.getTranslate("BatchPayment")); //პაკეტური გადარიცხვა

                    formItems = new Dictionary<string, object>();
                    itemName = "imptTypeCB"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("Length", 30);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                    formItems.Add("Left", left_e);
                    formItems.Add("Width", width_e);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Value", "singlePayment");
                    formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("ValidValues", listValidValuesDict);
                    formItems.Add("FromPane", 1);
                    formItems.Add("ToPane", 1);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }


                    formItems = new Dictionary<string, object>();
                    itemName = "DocTypeS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s + 5);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top + height + 1);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("DocType"));
                    formItems.Add("Visible", false);
                    formItems.Add("LinkTo", "DocTypeCB");
                    formItems.Add("FromPane", 1);
                    formItems.Add("ToPane", 1);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    listValidValuesDict = new Dictionary<string, string>
                    {
                        {"1", BDOSResources.getTranslate("All")},
                        {"2", BDOSResources.getTranslate("Salary")}
                    };

                    formItems = new Dictionary<string, object>();
                    itemName = "DocTypeCB"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("Length", 30);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                    formItems.Add("Left", left_e);
                    formItems.Add("Width", width_e);
                    formItems.Add("Top", top + height + 1);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
                    formItems.Add("ValidValues", listValidValuesDict);
                    formItems.Add("Visible", false);
                    formItems.Add("FromPane", 1);
                    formItems.Add("ToPane", 1);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }


                    formItems = new Dictionary<string, object>();
                    itemName = "batchNameE"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("Length", 30);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("Left", left_e + width_e + 1);
                    formItems.Add("Width", width_e);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("Visible", false);
                    formItems.Add("Description", BDOSResources.getTranslate("BatchName"));
                    formItems.Add("Value", BDOSResources.getTranslate("BatchPayment"));
                    formItems.Add("FromPane", 1);
                    formItems.Add("ToPane", 1);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    top = top + 3 * height + 1;

                    itemName = "checkB";
                    formItems = new Dictionary<string, object>();
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", 19);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Image", "HANA_CHECKBOX_CH");
                    formItems.Add("FromPane", 1);
                    formItems.Add("ToPane", 1);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left_s = left_s + 20;

                    itemName = "unCheckB";
                    formItems = new Dictionary<string, object>();
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", 19);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Image", "HANA_CHECKBOX_UH");
                    formItems.Add("FromPane", 1);
                    formItems.Add("ToPane", 1);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left_s = left_s + 20;

                    itemName = "fillB";
                    formItems = new Dictionary<string, object>();
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", 65);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("Fill"));
                    formItems.Add("FromPane", 1);
                    formItems.Add("ToPane", 1);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    listValidValuesDict = new Dictionary<string, string>();
                    listValidValuesDict.Add("updateStatus", BDOSResources.getTranslate("UpdateStatus"));
                    listValidValuesDict.Add("import", BDOSResources.getTranslate("Export"));

                    formItems = new Dictionary<string, object>();
                    itemName = "operationB";
                    formItems.Add("Caption", BDOSResources.getTranslate("Operations"));
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON_COMBO);
                    formItems.Add("Left", left_s + 65 + 2);
                    formItems.Add("Width", width_s - 20);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("ValidValues", listValidValuesDict);
                    formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
                    formItems.Add("AffectsFormMode", false);
                    formItems.Add("FromPane", 1);
                    formItems.Add("ToPane", 1);

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
                    formItems.Add("Left", left_s + 65 + 2 + width_s - 20 + 2);
                    formItems.Add("Width", width_s - 20);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("ValidValues", listValidValuesDict);
                    formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
                    formItems.Add("AffectsFormMode", false);
                    formItems.Add("FromPane", 1);
                    formItems.Add("ToPane", 1);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    top = top + height + 1;
                    left_s = 6;

                    formItems = new Dictionary<string, object>();
                    itemName = "importMTR"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", oForm.ClientWidth);
                    formItems.Add("Top", top);
                    formItems.Add("Height", (oForm.ClientHeight - 25 - top));
                    formItems.Add("UID", itemName);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("AffectsFormMode", false);
                    formItems.Add("FromPane", 1);
                    formItems.Add("ToPane", 1);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    SAPbouiCOM.Matrix oMatrixImport = ((SAPbouiCOM.Matrix)(oForm.Items.Item("importMTR").Specific));
                    oMatrixImport.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
                    SAPbouiCOM.Columns oColumns = oMatrixImport.Columns;
                    SAPbouiCOM.Column oColumn;

                    SAPbouiCOM.DataTable oDataTable;
                    oDataTable = oForm.DataSources.DataTables.Add("importMTR");

                    oDataTable.Columns.Add("CheckBox", SAPbouiCOM.BoFieldsType.ft_Text, 1); // 0 - ინდექსი 
                    oDataTable.Columns.Add("LineNum", SAPbouiCOM.BoFieldsType.ft_Integer, 50); // 0 - ინდექსი 
                    oDataTable.Columns.Add("DocEntry", SAPbouiCOM.BoFieldsType.ft_Integer, 50); //1 //ენთრი
                    oDataTable.Columns.Add("DocNum", SAPbouiCOM.BoFieldsType.ft_Integer, 50); //2 //ნომერი
                    oDataTable.Columns.Add("DocDate", SAPbouiCOM.BoFieldsType.ft_Date, 50); //3 //თარიღი
                    oDataTable.Columns.Add("OpType", SAPbouiCOM.BoFieldsType.ft_Text, 50); //4 //ოპერაციის ტიპი 
                    oDataTable.Columns.Add("TransferType", SAPbouiCOM.BoFieldsType.ft_Text, 50); //5 //ტრანსფერის სახე (ინტ.ბანკის ტიპები)
                    oDataTable.Columns.Add("PaymentID", SAPbouiCOM.BoFieldsType.ft_Text, 50); //6 //ტრანზაქციის ID
                    oDataTable.Columns.Add("BatchPaymentID", SAPbouiCOM.BoFieldsType.ft_Text, 50); //7 //პაკეტური ტრანზაქციის ID
                    oDataTable.Columns.Add("Status", SAPbouiCOM.BoFieldsType.ft_Text, 50); //8 //ინტ. ბანკში ტრანზაქციის სტატუსი
                    oDataTable.Columns.Add("BatchStatus", SAPbouiCOM.BoFieldsType.ft_Text, 50); //9 //ინტ. ბანკში პაკეტური ტრანზაქციის სტატუსი
                    oDataTable.Columns.Add("DebitAccount", SAPbouiCOM.BoFieldsType.ft_Text, 50); //10 //გამგზავნი ანგარიშის ნომერი
                    oDataTable.Columns.Add("DebitAccountCurrencyCode", SAPbouiCOM.BoFieldsType.ft_Text, 50); //11 //გამგზავნი ანგარიშის ვალუტა
                    oDataTable.Columns.Add("Amount", SAPbouiCOM.BoFieldsType.ft_Sum); //12 //თანხა
                    oDataTable.Columns.Add("Currency", SAPbouiCOM.BoFieldsType.ft_Text, 50); //13 //დოკუმენტის ვალუტა
                    oDataTable.Columns.Add("CardCode", SAPbouiCOM.BoFieldsType.ft_Text, 50); //14 //მიმღები (კოდი)
                    oDataTable.Columns.Add("BeneficiaryName", SAPbouiCOM.BoFieldsType.ft_Text, 50); //15 //მიმღები (სახელი)
                    oDataTable.Columns.Add("BeneficiaryTaxCode", SAPbouiCOM.BoFieldsType.ft_Text, 50); //16 //მიმღები (გსნ)
                    oDataTable.Columns.Add("CreditAccount", SAPbouiCOM.BoFieldsType.ft_Text, 50); //17 //მიმღები ანგარიშის ნომერი
                    oDataTable.Columns.Add("CreditAccountCurrencyCode", SAPbouiCOM.BoFieldsType.ft_Text, 50); //18 //მიმღები ანგარიშის ვალუტა
                    oDataTable.Columns.Add("Description", SAPbouiCOM.BoFieldsType.ft_Text, 50); //19 //დანიშნულება
                    oDataTable.Columns.Add("AdditionalDescription", SAPbouiCOM.BoFieldsType.ft_Text, 50); //20 //დამატებითი დანიშნულება
                    oDataTable.Columns.Add("TreasuryCode", SAPbouiCOM.BoFieldsType.ft_Text, 50); //21 //სახაზინო კოდი
                    //oDataTable.Columns.Add("PaymentType", SAPbouiCOM.BoFieldsType.ft_Text, 50); //22 //ტრანზაქციის მეთოდი (ინდივიდუალური, პაკეტური)
                    oDataTable.Columns.Add("BeneficiaryBankName", SAPbouiCOM.BoFieldsType.ft_Text, 50); //23 //მიმღები ბანკის სახელი
                    oDataTable.Columns.Add("BeneficiaryBankCode", SAPbouiCOM.BoFieldsType.ft_Text, 50); //24 //მიმღები ბანკის კოდი
                    oDataTable.Columns.Add("BeneficiaryAddress", SAPbouiCOM.BoFieldsType.ft_Text, 50); //25 //მისამართი
                    oDataTable.Columns.Add("ChargeDetails", SAPbouiCOM.BoFieldsType.ft_Text, 50); //26 //ხარჯი (SHA, OUR)
                    oDataTable.Columns.Add("Comments", SAPbouiCOM.BoFieldsType.ft_Text, 50); //27 //კომენტარი
                    oDataTable.Columns.Add("DocumentStatus", SAPbouiCOM.BoFieldsType.ft_Text, 50); //28 //სტატუსი 

                    SAPbouiCOM.LinkedButton oLink;

                    string UID = "importMTR";

                    SAPbobsCOM.UserTablesMD oUserTablesMD = null;
                    bool boolIdent = false;
                    oUserTablesMD = ((SAPbobsCOM.UserTablesMD)(Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)));
                    boolIdent = oUserTablesMD.GetByKey("OVPM");
                    Marshal.ReleaseComObject(oUserTablesMD);

                    SAPbobsCOM.Payments oVendorPayments = null;
                    oVendorPayments = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments);

                    foreach (SAPbouiCOM.DataColumn column in oDataTable.Columns)
                    {
                        string columnName = column.Name;

                        if (columnName == "LineNum")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = "#";
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind(UID, columnName);
                        }
                        else if (columnName == "CheckBox")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
                            oColumn.TitleObject.Caption = "";
                            oColumn.Editable = true;
                            oColumn.ValOff = "N";
                            oColumn.ValOn = "Y";
                            oColumn.DataBind.Bind(UID, columnName);
                        }
                        else if (columnName == "DocEntry")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate(columnName);
                            oColumn.Editable = false;
                            oLink = oColumn.ExtendedObject;
                            oLink.LinkedObjectType = "46";
                            oColumn.DataBind.Bind(UID, columnName);
                        }
                        else if (columnName == "OpType")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate(columnName);
                            oColumn.Visible = false;
                            oColumn.DataBind.Bind(UID, columnName);
                        }
                        else if (columnName == "TransferType")
                        {
                            oColumn = oColumns.Add("TrnsfrType", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate(columnName);
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind(UID, columnName);
                            oColumn.DisplayDesc = true;
                            oColumn.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;

                            string value = "TransferToOwnAccountPaymentOrderIo";
                            oColumn.ValidValues.Add(value, BDOSResources.getTranslate(value));
                            value = "TreasuryTransferPaymentOrderIo";
                            oColumn.ValidValues.Add(value, BDOSResources.getTranslate(value));
                            value = "TreasuryTransferPaymentOrderIoBP";
                            oColumn.ValidValues.Add(value, BDOSResources.getTranslate(value));
                            value = "TransferWithinBankPaymentOrderIo";
                            oColumn.ValidValues.Add(value, BDOSResources.getTranslate(value));
                            value = "TransferToOtherBankNationalCurrencyPaymentOrderIo";
                            oColumn.ValidValues.Add(value, BDOSResources.getTranslate(value));
                            value = "TransferToOtherBankForeignCurrencyPaymentOrderIo";
                            oColumn.ValidValues.Add(value, BDOSResources.getTranslate(value));
                            value = "TransferToNationalCurrencyPaymentOrderIo";
                            oColumn.ValidValues.Add(value, BDOSResources.getTranslate(value));
                            value = "TransferToForeignCurrencyPaymentOrderIo";
                            oColumn.ValidValues.Add(value, BDOSResources.getTranslate(value));
                            value = "CurrencyExchangePaymentOrderIo";
                            oColumn.ValidValues.Add(value, BDOSResources.getTranslate(value));
                        }
                        else if (columnName == "Status")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate(columnName);
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind(UID, columnName);
                            oColumn.DisplayDesc = true;
                            oColumn.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;

                            SAPbobsCOM.ValidValues validValues = oVendorPayments.UserFields.Fields.Item("U_status").ValidValues;
                            for (int i = 0; i < validValues.Count; i++)
                            {
                                string value = validValues.Item(i).Value;
                                oColumn.ValidValues.Add(value, BDOSResources.getTranslate(value));
                            }
                        }
                        else if (columnName == "BatchStatus")
                        {
                            oColumn = oColumns.Add("bStatus", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate(columnName);
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind(UID, columnName);
                            oColumn.DisplayDesc = true;
                            oColumn.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;

                            SAPbobsCOM.ValidValues validValues = oVendorPayments.UserFields.Fields.Item("U_bStatus").ValidValues;
                            for (int i = 0; i < validValues.Count; i++)
                            {
                                string value = validValues.Item(i).Value;
                                oColumn.ValidValues.Add(value, BDOSResources.getTranslate(value));
                            }
                        }
                        else if (columnName == "CardCode")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("BeneficiaryCode");
                            oColumn.Editable = false;
                            oLink = oColumn.ExtendedObject;
                            oLink.LinkedObjectType = "2";
                            oColumn.DataBind.Bind(UID, columnName);
                        }
                        else if (columnName == "ChargeDetails")
                        {
                            oColumn = oColumns.Add("chrgDtls", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate(columnName);
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind(UID, columnName);
                            oColumn.DisplayDesc = true;
                            oColumn.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;

                            SAPbobsCOM.ValidValues validValues = oVendorPayments.UserFields.Fields.Item("U_chrgDtls").ValidValues;
                            for (int i = 0; i < validValues.Count; i++)
                            {
                                string value = validValues.Item(i).Value;
                                oColumn.ValidValues.Add(value, BDOSResources.getTranslate(value));
                            }
                        }
                        else if (columnName == "DocumentStatus")
                        {
                            oColumn = oColumns.Add("DocStatus", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate(columnName);
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind(UID, columnName);
                            oColumn.DisplayDesc = true;
                            oColumn.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;

                            SAPbobsCOM.ValidValues validValues = oVendorPayments.UserFields.Fields.Item("Status").ValidValues;
                            for (int i = 0; i < validValues.Count; i++)
                            {
                                string value = validValues.Item(i).Value;
                                oColumn.ValidValues.Add(value, validValues.Item(i).Description);
                            }
                        }
                        else
                        {
                            if (columnName == "BatchPaymentID")
                            {
                                oColumn = oColumns.Add("BPaymentID", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            }
                            else if (columnName == "DebitAccount")
                            {
                                oColumn = oColumns.Add("DebitAcct", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            }
                            else if (columnName == "DebitAccountCurrencyCode")
                            {
                                oColumn = oColumns.Add("DedtActCur", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            }
                            else if (columnName == "BeneficiaryName")
                            {
                                oColumn = oColumns.Add("BenfName", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            }
                            else if (columnName == "BeneficiaryTaxCode")
                            {
                                oColumn = oColumns.Add("BenfTin", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            }
                            else if (columnName == "CreditAccount")
                            {
                                oColumn = oColumns.Add("CreditAcct", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            }
                            else if (columnName == "CreditAccountCurrencyCode")
                            {
                                oColumn = oColumns.Add("CrdtActCur", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            }
                            else if (columnName == "Description")
                            {
                                oColumn = oColumns.Add("Descrpt", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            }
                            else if (columnName == "AdditionalDescription")
                            {
                                oColumn = oColumns.Add("AddDescrpt", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            }
                            else if (columnName == "TreasuryCode")
                            {
                                oColumn = oColumns.Add("TresrCode", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            }
                            else if (columnName == "BeneficiaryBankName")
                            {
                                oColumn = oColumns.Add("BenfBankN", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            }
                            else if (columnName == "BeneficiaryBankCode")
                            {
                                oColumn = oColumns.Add("BenfBankC", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            }
                            else if (columnName == "BeneficiaryAddress")
                            {
                                oColumn = oColumns.Add("BenfAddrs", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            }
                            else
                            {
                                oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            }
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate(columnName);
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind(UID, columnName);
                        }
                    }

                    //<-----------------------------Import Data into Bank System-----------------------------

                    left_s = 6;
                    left_e = 120;
                    height = 15;
                    top = 25;
                    top = top + 15;
                    width_s = 121 - 15;
                    width_e = 148;

                    //----------------------------->Get Data from Bank System-----------------------------

                    int pane = 2;

                    formItems = new Dictionary<string, object>();
                    itemName = "TBCOB"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("Length", 1);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_OPTION_BUTTON);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", width_e);
                    formItems.Add("Top", top - 10);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", "TBC(Web - Service)");
                    formItems.Add("ValOn", "Y");
                    formItems.Add("ValOff", "N");
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("Selected", true);
                    formItems.Add("FromPane", pane);
                    formItems.Add("ToPane", pane);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "BOGOB"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("Length", 1);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_OPTION_BUTTON);
                    formItems.Add("Left", left_s + width_e);
                    formItems.Add("Width", width_e);
                    formItems.Add("Top", top - 10);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", "BOG (Web-Service)");
                    formItems.Add("GroupWith", "TBCOB");
                    formItems.Add("ValOn", "Y");
                    formItems.Add("ValOff", "N");
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("FromPane", pane);
                    formItems.Add("ToPane", pane);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    top = 40;
                    top = top + 15;
                    width_s = 140 - 15;

                    left_e = left_e + left_s + 5;

                    formItems = new Dictionary<string, object>();
                    itemName = "periodS2"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("Period"));
                    formItems.Add("LinkTo", "startDatE2");
                    formItems.Add("Visible", false);
                    formItems.Add("FromPane", pane);
                    formItems.Add("ToPane", pane);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    startDate = DateTime.Today;
                    startDateTxt = startDate.ToString("yyyyMMdd");

                    endDate = DateTime.Today;
                    endDateTxt = endDate.ToString("yyyyMMdd");

                    formItems = new Dictionary<string, object>();
                    itemName = "startDatE2";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_DATE);
                    formItems.Add("Length", 1);
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Left", left_e);
                    formItems.Add("Width", width_e / 2);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("ValueEx", startDateTxt);
                    formItems.Add("Visible", false);
                    formItems.Add("FromPane", pane);
                    formItems.Add("ToPane", pane);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "endDateE2";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_DATE);
                    formItems.Add("Length", 1);
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Left", left_e + width_e / 2);
                    formItems.Add("Width", width_e / 2);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("ValueEx", endDateTxt);
                    formItems.Add("Visible", false);
                    formItems.Add("FromPane", pane);
                    formItems.Add("ToPane", pane);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    top = 40;
                    top = top + 15;
                    width_s = 140 - 15;

                    top = top + height + 1;

                    multiSelection = false;
                    objectType = "231"; // HouseBankAccounts object
                    string uniqueID_lf_HouseBankAccountCFL2 = "HouseBankAccount_CFL2";
                    FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_lf_HouseBankAccountCFL2);

                    formItems = new Dictionary<string, object>();
                    itemName = "accountS2"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("Account"));
                    formItems.Add("LinkTo", "accountE2");
                    formItems.Add("Visible", false);
                    formItems.Add("FromPane", pane);
                    formItems.Add("ToPane", pane);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "accountE2"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("Length", 30);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("Left", left_e);
                    formItems.Add("Width", width_e);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("ChooseFromListUID", uniqueID_lf_HouseBankAccountCFL2);
                    formItems.Add("ChooseFromListAlias", "Account");
                    formItems.Add("Visible", false);
                    formItems.Add("FromPane", pane);
                    formItems.Add("ToPane", pane);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    listValidValuesDict = CommonFunctions.getCurrencyListForValidValues();

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
                    formItems.Add("Left", left_e + width_e + 5);
                    formItems.Add("Width", width_s / 2);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("ValidValues", listValidValuesDict);
                    formItems.Add("Description", BDOSResources.getTranslate("Currency"));
                    formItems.Add("Visible", false);
                    formItems.Add("FromPane", pane);
                    formItems.Add("ToPane", pane);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    top = top + height + 1;

                    formItems = new Dictionary<string, object>();
                    itemName = "transTypeS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("TransactionType"));
                    formItems.Add("LinkTo", "transTypCB");
                    formItems.Add("Visible", false);
                    formItems.Add("FromPane", pane);
                    formItems.Add("ToPane", pane);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    listValidValuesDict = new Dictionary<string, string>();
                    listValidValuesDict.Add(OperationTypeFromIntBank.None.ToString(), BDOSResources.getTranslate(OperationTypeFromIntBank.None.ToString()));
                    listValidValuesDict.Add(OperationTypeFromIntBank.BankCharge.ToString(), BDOSResources.getTranslate(OperationTypeFromIntBank.BankCharge.ToString()));
                    listValidValuesDict.Add(OperationTypeFromIntBank.CurrencyExchange.ToString(), BDOSResources.getTranslate(OperationTypeFromIntBank.CurrencyExchange.ToString()));
                    listValidValuesDict.Add(OperationTypeFromIntBank.OtherExpenses.ToString(), BDOSResources.getTranslate(OperationTypeFromIntBank.OtherExpenses.ToString()));
                    listValidValuesDict.Add(OperationTypeFromIntBank.OtherIncomes.ToString(), BDOSResources.getTranslate(OperationTypeFromIntBank.OtherIncomes.ToString()));
                    listValidValuesDict.Add(OperationTypeFromIntBank.TransferFromBP.ToString(), BDOSResources.getTranslate(OperationTypeFromIntBank.TransferFromBP.ToString()));
                    listValidValuesDict.Add(OperationTypeFromIntBank.TransferToBP.ToString(), BDOSResources.getTranslate(OperationTypeFromIntBank.TransferToBP.ToString()));
                    listValidValuesDict.Add(OperationTypeFromIntBank.TransferToOwnAccount.ToString(), BDOSResources.getTranslate(OperationTypeFromIntBank.TransferToOwnAccount.ToString()));
                    listValidValuesDict.Add(OperationTypeFromIntBank.TreasuryTransfer.ToString(), BDOSResources.getTranslate(OperationTypeFromIntBank.TreasuryTransfer.ToString()));
                    listValidValuesDict.Add(OperationTypeFromIntBank.WithoutSalary.ToString(), BDOSResources.getTranslate(OperationTypeFromIntBank.WithoutSalary.ToString()));
                    listValidValuesDict.Add(OperationTypeFromIntBank.TreasuryTransferPaymentOrderIoBP.ToString(), BDOSResources.getTranslate(OperationTypeFromIntBank.TreasuryTransferPaymentOrderIoBP.ToString()));

                    formItems = new Dictionary<string, object>();
                    itemName = "transTypCB"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("Length", 40);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                    formItems.Add("Left", left_e);
                    formItems.Add("Width", width_e);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
                    formItems.Add("ValidValues", listValidValuesDict);
                    formItems.Add("Description", BDOSResources.getTranslate("TransactionType"));
                    formItems.Add("Visible", false);
                    formItems.Add("FromPane", pane);
                    formItems.Add("ToPane", pane);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }
                    oForm.Items.Item("transTypCB").Specific.Select("WithoutSalary", SAPbouiCOM.BoSearchKey.psk_ByValue);

                    left_s = 6;

                    //Budget Cash Flow - Chartulia Alami da Gashvebulia Construction
                    /*if (CommonFunctions.IsDevelopment())
                    {
                        top = top + 2 * height + 1;      

                        formItems = new Dictionary<string, object>();
                        itemName = "BDOSDefCfS"; //10 characters
                        formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                        formItems.Add("Left", left_s);
                        formItems.Add("Width", width_s);
                        formItems.Add("Top", top);
                        formItems.Add("Height", height);
                        formItems.Add("FromPane", pane);
                        formItems.Add("ToPane", pane);
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
                        formItems.Add("DataSource", "UserDataSources");
                        formItems.Add("ValueEx", CommonFunctions.getOADM("U_BDOSDefCf"));
                        formItems.Add("TableName", "");
                        formItems.Add("Length", 200);
                        formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                        formItems.Add("Alias", "BDOSDefCfE");
                        formItems.Add("Bound", true);
                        formItems.Add("FromPane", pane);
                        formItems.Add("ToPane", pane);
                        formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                        formItems.Add("Left", left_e);
                        formItems.Add("Width", width_e / 3);
                        formItems.Add("Top", top);
                        formItems.Add("Height", height);
                        formItems.Add("UID", itemName);
                        formItems.Add("DisplayDesc", true);
                        formItems.Add("ChooseFromListUID", uniqueID_lf_Budg_CFL);
                        formItems.Add("ChooseFromListAlias", "Code");

                        FormsB1.createFormItem(oForm, formItems, out errorText);
                        if (errorText != null)
                        {
                            return;
                        }

                        string bCode = oForm.DataSources.UserDataSources.Item("BDOSDefCfE").ValueEx;
                        string bName = UDO.GetUDOFieldValueByParam("UDO_F_BDOSBUCFW_D", "Code", bCode, "Name");

                        formItems = new Dictionary<string, object>();
                        itemName = "BDOSDefCfN"; //10 characters
                        formItems.Add("isDataSource", true);
                        formItems.Add("DataSource", "UserDataSources");
                        formItems.Add("TableName", "");
                        formItems.Add("Length", 200);
                        formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                        formItems.Add("Alias", "BDOSDefCfN");
                        formItems.Add("ValueEx", bName);
                        formItems.Add("Bound", true);
                        formItems.Add("FromPane", pane);
                        formItems.Add("ToPane", pane);
                        formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                        formItems.Add("Left", left_e + width_e / 3 + 5);
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
                    }*/

                    top = top + 2 * height + 1;

                    left_s = 6;
                    left_e = 120;

                    itemName = "checkB2";
                    formItems = new Dictionary<string, object>();
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", 19);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Image", "HANA_CHECKBOX_CH");
                    formItems.Add("FromPane", pane);
                    formItems.Add("ToPane", pane);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left_s = left_s + 20;

                    itemName = "unCheckB2";
                    formItems = new Dictionary<string, object>();
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", 19);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Image", "HANA_CHECKBOX_UH");
                    formItems.Add("FromPane", pane);
                    formItems.Add("ToPane", pane);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left_s = left_s + 20;

                    itemName = "getDataB";
                    formItems = new Dictionary<string, object>();
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", 65);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("GetData"));
                    formItems.Add("FromPane", pane);
                    formItems.Add("ToPane", pane);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    itemName = "createDocB";
                    formItems = new Dictionary<string, object>();
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left_s + 65 + 2);
                    formItems.Add("Width", 65 * 2);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("CreateDocuments"));
                    formItems.Add("FromPane", pane);
                    formItems.Add("ToPane", pane);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left_s = left_s + 65 * 2;
                    itemName = "syncDate";
                    formItems = new Dictionary<string, object>();
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left_s + 65 + 4);
                    formItems.Add("Width", 65 * 2);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("SynchronizeDateAndCurrency"));
                    formItems.Add("FromPane", pane);
                    formItems.Add("ToPane", pane);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    listValidValuesDict = new Dictionary<string, string>();
                    listValidValuesDict.Add("withDetl", BDOSResources.getTranslate("withDetl"));
                    listValidValuesDict.Add("withoutDetl", BDOSResources.getTranslate("withoutDetl"));

                    formItems = new Dictionary<string, object>();
                    itemName = "listView"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("Length", 30);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                    formItems.Add("Left", left_s + 3 * 65 + 5);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("ValidValues", listValidValuesDict);
                    formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
                    formItems.Add("Visible", false);
                    formItems.Add("FromPane", pane);
                    formItems.Add("ToPane", pane);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }
                    oForm.Items.Item("listView").Specific.Select("withoutDetl", SAPbouiCOM.BoSearchKey.psk_ByValue);

                    //Automatic payment 
                    left_s = left_s + 5 + width_s;

                    formItems = new Dictionary<string, object>();
                    itemName = "autoPay"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("Left", left_s + 3 * 65 + 5);
                    formItems.Add("Length", 1);
                    formItems.Add("Width", width_e * 1.26);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("Description", BDOSResources.getTranslate("AutoPaymentInternetBank"));
                    formItems.Add("Caption", BDOSResources.getTranslate("AutoPaymentInternetBank"));
                    formItems.Add("ValOff", "N");
                    formItems.Add("ValOn", "Y");
                    formItems.Add("FromPane", pane);
                    formItems.Add("ToPane", pane);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    //Automatic payment 

                    top = top + height + 1;
                    left_s = 6;

                    formItems = new Dictionary<string, object>();
                    itemName = "exportMTR"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", oForm.ClientWidth);
                    formItems.Add("Top", top);
                    formItems.Add("Height", (oForm.ClientHeight - 25 - top));
                    formItems.Add("UID", itemName);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("AffectsFormMode", false);
                    formItems.Add("FromPane", pane);
                    formItems.Add("ToPane", pane);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    SAPbouiCOM.Matrix oMatrixExport = ((SAPbouiCOM.Matrix)(oForm.Items.Item("exportMTR").Specific));
                    oMatrixExport.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
                    oColumns = oMatrixExport.Columns;
                    oDataTable = oForm.DataSources.DataTables.Add("exportMTR");

                    oDataTable.Columns.Add("CheckBox", SAPbouiCOM.BoFieldsType.ft_Text, 1);
                    oDataTable.Columns.Add("LineNum", SAPbouiCOM.BoFieldsType.ft_Integer, 100);
                    oDataTable.Columns.Add("DocEntry", SAPbouiCOM.BoFieldsType.ft_Text, 100);
                    oDataTable.Columns.Add("DocNum", SAPbouiCOM.BoFieldsType.ft_Text, 100);
                    oDataTable.Columns.Add("PaymentID", SAPbouiCOM.BoFieldsType.ft_Text, 100);
                    oDataTable.Columns.Add("ExternalPaymentID", SAPbouiCOM.BoFieldsType.ft_Text, 100);
                    oDataTable.Columns.Add("ValueDate", SAPbouiCOM.BoFieldsType.ft_Date);
                    oDataTable.Columns.Add("DocumentDate", SAPbouiCOM.BoFieldsType.ft_Date);
                    oDataTable.Columns.Add("DebitCredit", SAPbouiCOM.BoFieldsType.ft_Text, 100);
                    oDataTable.Columns.Add("AccountNumber", SAPbouiCOM.BoFieldsType.ft_Text, 100);
                    oDataTable.Columns.Add("Currency", SAPbouiCOM.BoFieldsType.ft_Text, 100);
                    oDataTable.Columns.Add("Description", SAPbouiCOM.BoFieldsType.ft_Text, 100);
                    oDataTable.Columns.Add("AdditionalDescription", SAPbouiCOM.BoFieldsType.ft_Text, 100);
                    oDataTable.Columns.Add("Amount", SAPbouiCOM.BoFieldsType.ft_Sum);
                    oDataTable.Columns.Add("CurrencyExchange", SAPbouiCOM.BoFieldsType.ft_Text, 100);
                    oDataTable.Columns.Add("Rate", SAPbouiCOM.BoFieldsType.ft_Text, 100);
                    oDataTable.Columns.Add("PartnerAccountNumber", SAPbouiCOM.BoFieldsType.ft_Text, 100);
                    oDataTable.Columns.Add("PartnerCurrency", SAPbouiCOM.BoFieldsType.ft_Text, 100);
                    oDataTable.Columns.Add("BPCode", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 15);
                    oDataTable.Columns.Add("PartnerName", SAPbouiCOM.BoFieldsType.ft_Text, 100);
                    oDataTable.Columns.Add("PartnerTaxCode", SAPbouiCOM.BoFieldsType.ft_Text, 100);
                    oDataTable.Columns.Add("PartnerBankCode", SAPbouiCOM.BoFieldsType.ft_Text, 100);
                    oDataTable.Columns.Add("PartnerBank", SAPbouiCOM.BoFieldsType.ft_Text, 100);
                    oDataTable.Columns.Add("ChargeDetail", SAPbouiCOM.BoFieldsType.ft_Text, 100);
                    oDataTable.Columns.Add("TreasuryCode", SAPbouiCOM.BoFieldsType.ft_Text, 100);
                    oDataTable.Columns.Add("DocumentNumber", SAPbouiCOM.BoFieldsType.ft_Text, 100);
                    oDataTable.Columns.Add("OperationCode", SAPbouiCOM.BoFieldsType.ft_Text, 100);
                    oDataTable.Columns.Add("TransactionType", SAPbouiCOM.BoFieldsType.ft_Text, 100);
                    oDataTable.Columns.Add("TransactionCode", SAPbouiCOM.BoFieldsType.ft_Text, 100);
                    oDataTable.Columns.Add("GLAccountCode", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 15);
                    oDataTable.Columns.Add("Project", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20);
                    oDataTable.Columns.Add("UseBlaAgRt", SAPbouiCOM.BoFieldsType.ft_Text, 1);
                    oDataTable.Columns.Add("BlnkAgr", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20);
                    oDataTable.Columns.Add("CashFlowLineItemID", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 11);
                    oDataTable.Columns.Add("CashFlowLineItemName", SAPbouiCOM.BoFieldsType.ft_Text, 100);

                    if (CommonFunctions.IsDevelopment())
                    {
                        oDataTable.Columns.Add("BudgetCashFlowID", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 11);
                        oDataTable.Columns.Add("BudgetCashFlowName", SAPbouiCOM.BoFieldsType.ft_Text, 100);
                    }

                    oDataTable.Columns.Add("DownPaymentAmount", SAPbouiCOM.BoFieldsType.ft_Sum); //ავანსის თანხა
                    oDataTable.Columns.Add("InvoicesAmount", SAPbouiCOM.BoFieldsType.ft_Sum); //ინვოისის თანხა
                    oDataTable.Columns.Add("PaymentOnAccount", SAPbouiCOM.BoFieldsType.ft_Sum); //ბპ. ანგარიშზე   
                    oDataTable.Columns.Add("DocRateIN", SAPbouiCOM.BoFieldsType.ft_Rate); //ბპ. ვალუტის კურსი 
                    oDataTable.Columns.Add("InDetail", SAPbouiCOM.BoFieldsType.ft_Text, 20);

                    //Automatic Payment
                    oDataTable.Columns.Add("AddDownPaymentAmount", SAPbouiCOM.BoFieldsType.ft_Sum); //Additional Down Payment

                    oDataTable.Columns.Add("Test", SAPbouiCOM.BoFieldsType.ft_Text, 100);

                    UID = "exportMTR";

                    multiSelection = false;
                    objectType = "1"; //oChartOfAccounts
                    string uniqueID_lf_GLAccountCFL = "GLAccount_CFL";
                    FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_lf_GLAccountCFL);

                    multiSelection = false;
                    objectType = "242"; //CashFlowLineItem
                    string uniqueID_lf_CashFlowLineItemCFL = "CashFlowLineItem_CFL";
                    FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_lf_CashFlowLineItemCFL);

                    string uniqueID_lf_Budg_CFL = "Budg_CFL";
                    if (CommonFunctions.IsDevelopment())
                    {
                        multiSelection = false;
                        objectType = "UDO_F_BDOSBUCFW_D";
                        FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_lf_Budg_CFL);
                    }

                    multiSelection = false;
                    objectType = "63"; //Project
                    string uniqueID_lf_Project = "Project_CFLA";
                    FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_lf_Project);

                    multiSelection = false;
                    objectType = "1250000025"; //Blanket Agreement
                    string uniqueID_BlnkAgrCFL = "BlnkAgr_CFL";
                    FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_BlnkAgrCFL);
                    int o = 0;


                    multiSelection = false;
                    objectType = "2";
                    string uniqueID_lf_BP_CFL = "BusinessPartner_CFL";
                    FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_lf_BP_CFL);

                    foreach (SAPbouiCOM.DataColumn column in oDataTable.Columns)
                    {
                        o++;
                        string columnName = column.Name;

                        if (columnName == "LineNum")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = "#";
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind(UID, columnName);
                        }
                        else if (columnName == "CheckBox")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
                            oColumn.TitleObject.Caption = "";
                            oColumn.Editable = true;
                            oColumn.ValOff = "N";
                            oColumn.ValOn = "Y";
                            oColumn.DataBind.Bind(UID, columnName);
                        }
                        else if (columnName == "DocEntry")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate(columnName);
                            oColumn.Editable = false;
                            oLink = oColumn.ExtendedObject;
                            oLink.LinkedObjectType = "46"; //"24"
                            oColumn.DataBind.Bind(UID, columnName);
                            oColumn.Width = 50;

                        }
                        else if (columnName == "PaymentID")
                        {
                            oColumn = oColumns.Add("PaymentID", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate(columnName);
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind(UID, columnName);
                        }
                        else if (columnName == "TransactionType")
                        {
                            oColumn = oColumns.Add("TransType", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate(columnName);
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind(UID, columnName);
                            oColumn.DisplayDesc = true;
                            oColumn.TitleObject.Sortable = true;
                            oColumn.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
                            //SAPbobsCOM.SBObob oSBOBob = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);

                            oColumn.ValidValues.Add(OperationTypeFromIntBank.None.ToString(), BDOSResources.getTranslate(OperationTypeFromIntBank.None.ToString()));
                            oColumn.ValidValues.Add(OperationTypeFromIntBank.BankCharge.ToString(), BDOSResources.getTranslate(OperationTypeFromIntBank.BankCharge.ToString()));
                            oColumn.ValidValues.Add(OperationTypeFromIntBank.CurrencyExchange.ToString(), BDOSResources.getTranslate(OperationTypeFromIntBank.CurrencyExchange.ToString()));
                            oColumn.ValidValues.Add(OperationTypeFromIntBank.OtherExpenses.ToString(), BDOSResources.getTranslate(OperationTypeFromIntBank.OtherExpenses.ToString()));
                            oColumn.ValidValues.Add(OperationTypeFromIntBank.OtherIncomes.ToString(), BDOSResources.getTranslate(OperationTypeFromIntBank.OtherIncomes.ToString()));
                            oColumn.ValidValues.Add(OperationTypeFromIntBank.TransferFromBP.ToString(), BDOSResources.getTranslate(OperationTypeFromIntBank.TransferFromBP.ToString()));
                            oColumn.ValidValues.Add(OperationTypeFromIntBank.TransferToBP.ToString(), BDOSResources.getTranslate(OperationTypeFromIntBank.TransferToBP.ToString()));
                            oColumn.ValidValues.Add(OperationTypeFromIntBank.TransferToOwnAccount.ToString(), BDOSResources.getTranslate(OperationTypeFromIntBank.TransferToOwnAccount.ToString()));
                            oColumn.ValidValues.Add(OperationTypeFromIntBank.TreasuryTransfer.ToString(), BDOSResources.getTranslate(OperationTypeFromIntBank.TreasuryTransfer.ToString()));
                            oColumn.ValidValues.Add(OperationTypeFromIntBank.TreasuryTransferPaymentOrderIoBP.ToString(), BDOSResources.getTranslate(OperationTypeFromIntBank.TreasuryTransferPaymentOrderIoBP.ToString()));
                        }
                        else if (columnName == "TransactionCode")
                        {
                            oColumn = oColumns.Add("TransCode", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate(columnName);
                            oColumn.Visible = false;
                            oColumn.DataBind.Bind(UID, columnName);
                        }
                        else if (columnName == "DebitCredit")
                        {
                            oColumn = oColumns.Add("DbitCrdit", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate(columnName);
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind(UID, columnName);
                            oColumn.DisplayDesc = true;
                            oColumn.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;

                            oColumn.ValidValues.Add("0", BDOSResources.getTranslate("Output")); //გასვლა 
                            oColumn.ValidValues.Add("1", BDOSResources.getTranslate("Input")); //შემოსვლა
                            oColumn.ValidValues.Add("-1", BDOSResources.getTranslate("NONE")); //შემოსვლა
                        }
                        else if (columnName == "GLAccountCode")
                        {
                            oColumn = oColumns.Add("GLAcctCode", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate(columnName);
                            oColumn.Editable = true;
                            oColumn.DataBind.Bind(UID, columnName);
                            oColumn.ChooseFromListUID = uniqueID_lf_GLAccountCFL;
                            oColumn.ChooseFromListAlias = "AcctCode";
                            oLink = oColumn.ExtendedObject;
                            oLink.LinkedObjectType = "1";
                        }
                        else if (columnName == "Project")
                        {
                            oColumn = oColumns.Add("Project", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate(columnName);
                            oColumn.Editable = true;
                            oColumn.DataBind.Bind(UID, columnName);
                            oColumn.ChooseFromListUID = uniqueID_lf_Project;
                            oColumn.ChooseFromListAlias = "PrjCode";
                            oLink = oColumn.ExtendedObject;
                            oLink.LinkedObjectType = "63";
                        }
                        else if (columnName == "UseBlaAgRt")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("UseBlAgrRt");
                            oColumn.Editable = false;
                            oColumn.ValOff = "N";
                            oColumn.ValOn = "Y";
                            oColumn.DataBind.Bind(UID, columnName);

                        }
                        else if (columnName == "BlnkAgr")
                        {
                            oColumn = oColumns.Add("BlnkAgr", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate(columnName);
                            oColumn.Editable = true;
                            oColumn.DataBind.Bind(UID, columnName);
                            oColumn.ChooseFromListUID = uniqueID_BlnkAgrCFL;
                            oColumn.ChooseFromListAlias = "AbsID";
                            oLink = oColumn.ExtendedObject;
                            oLink.LinkedObjectType = "1250000025";
                        }
                        else if (columnName == "CashFlowLineItemID")
                        {
                            oColumn = oColumns.Add("CFWId", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate(columnName);
                            oColumn.Editable = true;
                            oColumn.DataBind.Bind(UID, columnName);
                            oColumn.ChooseFromListUID = uniqueID_lf_CashFlowLineItemCFL;
                            oColumn.ChooseFromListAlias = "CFWId";
                        }
                        else if (columnName == "CashFlowLineItemName")
                        {
                            oColumn = oColumns.Add("CFWName", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate(columnName);
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind(UID, columnName);
                        }
                        else if (columnName == "BudgetCashFlowID")
                        {
                            oColumn = oColumns.Add("BCFWId", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("BudgetCashFlowCodeOutgoingWizard");
                            oColumn.Editable = true;
                            oColumn.DataBind.Bind(UID, columnName);
                            oColumn.ChooseFromListUID = uniqueID_lf_Budg_CFL;
                            oColumn.ChooseFromListAlias = "Code";
                        }
                        else if (columnName == "BudgetCashFlowName")
                        {
                            oColumn = oColumns.Add("BCFWName", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("BudgetCashFlowCodeOutgoingWizard");
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind(UID, columnName);
                        }
                        else if (columnName == "DocRateIN")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("BPCurrencyRate");
                            oColumn.Editable = false;
                            oColumn.Visible = false;
                            oColumn.DataBind.Bind(UID, columnName);
                        }
                        else if (columnName == "BPCode")
                        {
                            oColumn = oColumns.Add("BPCode", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("BPCardCode");
                            oColumn.Editable = false;
                            oColumn.Visible = false;
                            oColumn.DataBind.Bind(UID, columnName);
                            oLink = oColumn.ExtendedObject;
                            oLink.LinkedObjectType = "2";
                            oColumn.ChooseFromListUID = uniqueID_lf_BP_CFL;
                            oColumn.ChooseFromListAlias = "CardCode";
                        }
                        else
                        {
                            if (columnName == "ExternalPaymentID")
                            {
                                oColumn = oColumns.Add("ExPaymntID", SAPbouiCOM.BoFormItemTypes.it_EDIT);

                            }
                            else if (columnName == "Description")
                            {
                                oColumn = oColumns.Add("Descrpt", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                                oColumn.Width = 50;
                            }
                            else if (columnName == "CurrencyExchange")
                            {
                                oColumn = oColumns.Add("CurrencyEx", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            }
                            else if (columnName == "AccountNumber")
                            {
                                oColumn = oColumns.Add("AcctNumber", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            }
                            else if (columnName == "AccountName")
                            {
                                oColumn = oColumns.Add("AcctName", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            }
                            else if (columnName == "AdditionalInformation")
                            {
                                oColumn = oColumns.Add("AdditnInfo", SAPbouiCOM.BoFormItemTypes.it_EDIT);

                            }
                            else if (columnName == "DocumentDate")
                            {
                                oColumn = oColumns.Add("DocDate", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            }
                            else if (columnName == "DocumentNumber")
                            {
                                oColumn = oColumns.Add("DocNumber", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            }
                            else if (columnName == "PartnerAccountNumber")
                            {
                                oColumn = oColumns.Add("PAcctNmber", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            }
                            else if (columnName == "PartnerCurrency")
                            {
                                oColumn = oColumns.Add("PCurrency", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            }
                            else if (columnName == "PartnerName")
                            {
                                oColumn = oColumns.Add("PName", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            }
                            else if (columnName == "PartnerTaxCode")
                            {
                                oColumn = oColumns.Add("PTaxCode", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            }
                            else if (columnName == "PartnerBankCode")
                            {
                                oColumn = oColumns.Add("PBankCode", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            }
                            else if (columnName == "PartnerBank")
                            {
                                oColumn = oColumns.Add("PBank", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            }
                            else if (columnName == "IntermediaryBankCode")
                            {
                                oColumn = oColumns.Add("IntBnkCode", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            }
                            else if (columnName == "IntermediaryBank")
                            {
                                oColumn = oColumns.Add("IntBank", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            }
                            else if (columnName == "ChargeDetail")
                            {
                                oColumn = oColumns.Add("ChrgDtls", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            }
                            else if (columnName == "TaxpayerCode")
                            {
                                oColumn = oColumns.Add("TxpyerCode", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            }
                            else if (columnName == "TaxpayerName")
                            {
                                oColumn = oColumns.Add("TxpyerName", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            }
                            else if (columnName == "TreasuryCode")
                            {
                                oColumn = oColumns.Add("TresrCode", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            }
                            else if (columnName == "OperationCode")
                            {
                                oColumn = oColumns.Add("OpCode", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            }
                            else if (columnName == "AdditionalDescription")
                            {
                                oColumn = oColumns.Add("AddDescrpt", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                                oColumn.Width = 50;
                            }
                            else if (columnName == "InvoicesAmount")
                            {
                                oColumn = oColumns.Add("InvAmt", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            }
                            else if (columnName == "DownPaymentAmount")
                            {
                                oColumn = oColumns.Add("DPayment", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            }
                            else if (columnName == "PaymentOnAccount")
                            {
                                oColumn = oColumns.Add("OnAccount", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            }
                            else if (columnName == "AddDownPaymentAmount")
                            {
                                oColumn = oColumns.Add("AddDPAmt", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                                //oColumn.Width = 10;
                                oColumn.Visible = false;
                            }
                            else if (columnName == "InDetail")
                            {
                                oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_PICTURE);
                            }
                            else
                            {
                                oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            }

                            oColumn.TitleObject.Caption = BDOSResources.getTranslate(columnName);
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind(UID, columnName);
                        }
                    }
                    //<-----------------------------Get Data from Bank System-----------------------------

                    Marshal.FinalReleaseComObject(oVendorPayments);

                    oMatrixImport.Clear();
                    oMatrixImport.LoadFromDataSource();
                    oMatrixImport.AutoResizeColumns();

                    oMatrixExport.Clear();
                    oMatrixExport.LoadFromDataSource();
                    oMatrixExport.AutoResizeColumns();

                    setVisibleFormItemsMatrixColumns(oForm, out errorText);

                }

                createFolder(oForm, out errorText);

                oForm.Visible = true;
                oForm.Select();

                oForm.Freeze(true);
                oForm.Items.Item("FolderImp").Click();
                oForm.Freeze(false);
            }
            GC.Collect();
        }

        public static bool SynchronizePayments(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("exportMTR").Specific;
            SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("exportMTR");
            SAPbobsCOM.Payments oPayments;

            SAPbobsCOM.Payments oNewPayments;

            int errCode;
            string errMsg;

            int docEntry;
            string debitCredit;

            SAPbobsCOM.SBObob oSBOBob = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);
            string localCurrency = CommonFunctions.getLocalCurrency();

            bool selected;

            double correctedRate = 0;

            //RowsWithCorrectedDates = new List<int>();
            if (RowsWithDifferentDates == null)
            {
                return false;
            }

            foreach (int rowNum in RowsWithDifferentDates)
            {
                selected = oDataTable.GetValue("CheckBox", rowNum - 1) == "Y" ? true : false;

                if (!selected)
                    continue;

                docEntry = Convert.ToInt32(oDataTable.GetValue("DocEntry", rowNum - 1));
                debitCredit = oDataTable.GetValue("DebitCredit", rowNum - 1);

                string transTypeForFilter = oDataTable.GetValue("TransactionType", rowNum - 1);

                if (transTypeForFilter != OperationTypeFromIntBank.TransferToBP.ToString() && transTypeForFilter != OperationTypeFromIntBank.TreasuryTransferPaymentOrderIoBP.ToString())
                {
                    continue;
                }

                if (debitCredit == "0")
                {
                    oNewPayments = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments);
                    oPayments = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments);  //oVendorPayments
                }
                else
                {
                    continue;
                }

                oPayments.GetByKey(docEntry);

                if (!(docEntry > 0))
                {
                    continue;
                }

                //თუ იბეგრება არ დასინქრონიზდეს
                double WTaxAmount = oPayments.WTAmount;
                string prBase = oPayments.UserFields.Fields.Item("U_prBase").Value;
                if (WTaxAmount > 0 || prBase != "")
                {
                    Program.uiApp.SetStatusBarMessage(BDOSResources.getTranslate("TableRow") + " #" + rowNum + ". " + BDOSResources.getTranslate("SynchNotPossibleDocumentLiableProfitTax"));
                    continue;
                }


                bool mustContinue = false;

                for (int InvRow = 0; InvRow < oPayments.Invoices.Count; InvRow++)
                {
                    oPayments.Invoices.SetCurrentLine(InvRow);
                    int DocEntry = oPayments.Invoices.DocEntry;
                    SAPbobsCOM.BoRcptInvTypes InvoiceType = oPayments.Invoices.InvoiceType;

                    if (InvoiceType == SAPbobsCOM.BoRcptInvTypes.it_PurchaseInvoice ||
                        InvoiceType == SAPbobsCOM.BoRcptInvTypes.it_PurchaseDownPayment)
                    {
                        SAPbobsCOM.Documents oPurchInvoice;
                        if (InvoiceType == SAPbobsCOM.BoRcptInvTypes.it_PurchaseInvoice)
                        {
                            oPurchInvoice = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices);
                        }
                        else
                        {
                            oPurchInvoice = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDownPayments);
                        }
                        oPurchInvoice.GetByKey(DocEntry);

                        WTaxAmount = oPurchInvoice.WTAmount;
                        prBase = oPurchInvoice.UserFields.Fields.Item("U_prBase").Value;
                        if (WTaxAmount > 0 || prBase != "")
                        {
                            Program.uiApp.SetStatusBarMessage(BDOSResources.getTranslate("TableRow") + " #" + rowNum + ". " + BDOSResources.getTranslate("SynchNotPossibleDocumentLiableProfitTax"));
                            mustContinue = true;
                            break;
                        }
                    }
                }

                if (mustContinue)
                {
                    continue;
                }

                decimal invoicesSumApplied = 0;
                bool notEnoughFC;

                notEnoughFC = false;

                string cardCode = oPayments.CardCode;
                SAPbobsCOM.BusinessPartners businessPartner = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);
                businessPartner.GetByKey(cardCode);

                DateTime docDate = oDataTable.GetValue("DocumentDate", rowNum - 1);

                var dtInvoicesSorted = OutgoingPayment.GetPaymentInvoices(oPayments.DocEntry, OutgoingPayment.PaymentType.Payment, docDate);

                //ALL CURRENCY
                string bpCurrency = businessPartner.Currency;

                string DiffCurr = "";
                string DocCurr = oPayments.DocCurrency;

                if (bpCurrency == "##" || String.IsNullOrEmpty(bpCurrency))
                {
                    var invoicesDT = dtInvoicesSorted;
                    var currencies = invoicesDT.AsEnumerable().Select(x => x["DocCur"]);
                    string firstcurrency = (string)currencies.FirstOrDefault();
                    int otherCurrenciesCount = invoicesDT.AsEnumerable().Where(x => (string)x["DocCur"] != firstcurrency).Count();
                    int firstCurrencyCount = invoicesDT.AsEnumerable().Where(x => (string)x["DocCur"] == firstcurrency).Count();

                    if (otherCurrenciesCount > 0)
                    {
                        Program.uiApp.SetStatusBarMessage(BDOSResources.getTranslate("InvoicesDifferentCurrenciesError") + " - " + rowNum, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                        continue;
                    }
                    else
                    {
                        DiffCurr = (DocCurr != firstcurrency) ? "Y" : "N";
                        DocCurr = firstcurrency;
                    }
                }

                if (CommonFunctions.IsDevelopment())
                {
                    string bCode = oPayments.UserFields.Fields.Item("U_BDOSBdgCf").Value;

                    if (String.IsNullOrEmpty(bCode) == false)
                    {
                        oNewPayments.UserFields.Fields.Item("U_BDOSBdgCf").Value = bCode;
                        oNewPayments.UserFields.Fields.Item("U_BDOSBdgCfN").Value = UDO.GetUDOFieldValueByParam("UDO_F_BDOSBUCFW_D", "Code", bCode, "Name");
                    }
                }

                oNewPayments.CardCode = oPayments.CardCode;
                oNewPayments.DocDate = oDataTable.GetValue("DocumentDate", rowNum - 1);
                oNewPayments.TaxDate = oDataTable.GetValue("DocumentDate", rowNum - 1);
                oNewPayments.VatDate = oPayments.VatDate;

                oNewPayments.TransferAccount = oPayments.TransferAccount;
                oNewPayments.TransferDate = oDataTable.GetValue("DocumentDate", rowNum - 1);
                //oNewPayments.TransferSum = oPayments.TransferSum;

                oNewPayments.ControlAccount = oPayments.ControlAccount;

                oNewPayments.DocCurrency = oPayments.DocCurrency;

                oNewPayments.WTCode = oPayments.WTCode;
                oNewPayments.IsPayToBank = oPayments.IsPayToBank;
                oNewPayments.PayToBankCountry = oPayments.PayToBankCountry;
                oNewPayments.PayToBankCode = oPayments.PayToBankCode;
                oNewPayments.PayToBankAccountNo = oPayments.PayToBankAccountNo;

                oNewPayments.PayToBankBranch = oPayments.PayToBankBranch;
                oNewPayments.ProjectCode = oPayments.ProjectCode;
                oNewPayments.BlanketAgreement = oPayments.BlanketAgreement;

                oNewPayments.LocalCurrency = oPayments.LocalCurrency;

                oNewPayments.Remarks = BDOSResources.getTranslate("CreatedAutomaticallyInPlaceOfPaymentDocument") + " " + docEntry;

                //oNewPayments.UserFields.Fields.Item("U_status").Value = oPayments.UserFields.Fields.Item("U_status").Value;
                oNewPayments.UserFields.Fields.Item("U_status").Value = "downloadedFromTheBank";
                oNewPayments.UserFields.Fields.Item("U_paymentID").Value = oPayments.UserFields.Fields.Item("U_paymentID").Value;
                oNewPayments.UserFields.Fields.Item("U_chrgDtls").Value = oPayments.UserFields.Fields.Item("U_chrgDtls").Value;
                oNewPayments.UserFields.Fields.Item("U_descrpt").Value = oPayments.UserFields.Fields.Item("U_descrpt").Value;
                oNewPayments.UserFields.Fields.Item("U_addDescrpt").Value = oPayments.UserFields.Fields.Item("U_addDescrpt").Value;

                oNewPayments.UserFields.Fields.Item("U_docNumber").Value = oPayments.UserFields.Fields.Item("U_docNumber").Value;
                oNewPayments.UserFields.Fields.Item("U_transCode").Value = oPayments.UserFields.Fields.Item("U_transCode").Value;
                oNewPayments.UserFields.Fields.Item("U_ePaymentID").Value = oPayments.UserFields.Fields.Item("U_ePaymentID").Value;
                oNewPayments.UserFields.Fields.Item("U_opCode").Value = oPayments.UserFields.Fields.Item("U_opCode").Value;

                bool cashFlowRelevant = CommonFunctions.isAccountCashFlowRelevant(oNewPayments.TransferAccount);

                if (cashFlowRelevant)
                {
                    for (int j = 0; j < oPayments.PrimaryFormItems.Count; j++)
                    {
                        oPayments.PrimaryFormItems.SetCurrentLine(j);
                        oNewPayments.PrimaryFormItems.CashFlowLineItemID = oPayments.PrimaryFormItems.CashFlowLineItemID;
                        oNewPayments.PrimaryFormItems.AmountFC = oPayments.PrimaryFormItems.AmountFC;

                        if (oPayments.DocCurrency == localCurrency)
                            oNewPayments.PrimaryFormItems.AmountLC = oPayments.PrimaryFormItems.AmountLC;

                        oNewPayments.PrimaryFormItems.PaymentMeans = oPayments.PrimaryFormItems.PaymentMeans;
                        oNewPayments.PrimaryFormItems.Add();
                    }
                }

                decimal TransferSumInCurrency = 0;
                TransferSumInCurrency = Convert.ToDecimal(oPayments.TransferSum);

                bool changeRate = (DiffCurr == "N") || ((!String.IsNullOrEmpty(oPayments.DocCurrency)) && (localCurrency != oPayments.DocCurrency));

                if (changeRate)
                {
                    oNewPayments.DocRate = oSBOBob.GetCurrencyRate(oPayments.DocCurrency, oNewPayments.DocDate).Fields.Item("CurrencyRate").Value;
                    correctedRate = oNewPayments.DocRate;

                    if (oNewPayments.LocalCurrency == SAPbobsCOM.BoYesNoEnum.tYES)
                    {
                        TransferSumInCurrency = TransferSumInCurrency / Convert.ToDecimal(oNewPayments.DocRate);
                    }
                    else
                    {
                        TransferSumInCurrency = Convert.ToDecimal(oDataTable.GetValue("Amount", rowNum - 1));
                    }
                }

                var dtTableoPaymentsInvoices = new DataTable();
                dtTableoPaymentsInvoices.Columns.Add("InstlmntID", typeof(int));
                dtTableoPaymentsInvoices.Columns.Add("InvoiceDocEntry", typeof(int));
                dtTableoPaymentsInvoices.Columns.Add("InvType", typeof(int));
                dtTableoPaymentsInvoices.Columns.Add("LineNum", typeof(int));

                for (int j = 0; j < oPayments.Invoices.Count; j++)
                {
                    oPayments.Invoices.SetCurrentLine(j);

                    var rowTemp = dtTableoPaymentsInvoices.NewRow();
                    rowTemp["InstlmntID"] = oPayments.Invoices.InstallmentId;
                    rowTemp["InvoiceDocEntry"] = oPayments.Invoices.DocEntry;
                    rowTemp["InvType"] = oPayments.Invoices.InvoiceType;
                    rowTemp["LineNum"] = j;

                    dtTableoPaymentsInvoices.Rows.Add(rowTemp);

                }

                if (DiffCurr == "Y")
                {
                    decimal DocRate = Convert.ToDecimal(oSBOBob.GetCurrencyRate(DocCurr, oNewPayments.DocDate).Fields.Item("CurrencyRate").Value);
                    TransferSumInCurrency = CommonFunctions.roundAmountByGeneralSettings(TransferSumInCurrency / DocRate, "Sum");
                }

                //foreach (KeyValuePair<int,DateTime> invoiceEntry in invoicesSorted)
                foreach (DataRow invoiceEntry in dtInvoicesSorted.Rows)
                {
                    if (TransferSumInCurrency <= 0)
                        break;

                    //oPayments.Invoices.SetCurrentLine(j);
                    var dtFoundRow = dtTableoPaymentsInvoices.AsEnumerable().Where(
                        x => (int)x["InvoiceDocEntry"] == (int)invoiceEntry["InvoiceDocEntry"] &&
                                                            (int)x["InvType"] == (int)invoiceEntry["InvType"] &&
                                                            (int)x["InstlmntID"] == (int)invoiceEntry["InstlmntID"]

                                                            ).FirstOrDefault();

                    oPayments.Invoices.SetCurrentLine((int)dtFoundRow["LineNum"]);
                    oNewPayments.Invoices.DocEntry = oPayments.Invoices.DocEntry;
                    oNewPayments.Invoices.InstallmentId = oPayments.Invoices.InstallmentId;

                    oNewPayments.Invoices.InvoiceType = oPayments.Invoices.InvoiceType;
                    oNewPayments.Invoices.DocLine = oPayments.Invoices.DocLine;

                    if ((DiffCurr == "Y") || ((!String.IsNullOrEmpty(oPayments.DocCurrency)) && (localCurrency != oPayments.DocCurrency)))
                    {

                        double minAmt;
                        if (Math.Abs(oPayments.Invoices.AppliedFC) > Convert.ToDouble(TransferSumInCurrency))
                        {
                            minAmt = Convert.ToDouble(TransferSumInCurrency);
                        }
                        else
                        {
                            minAmt = oPayments.Invoices.AppliedFC;
                        }

                        oNewPayments.Invoices.AppliedFC = Math.Abs(Convert.ToDouble(CommonFunctions.roundAmountByGeneralSettings(Convert.ToDecimal(minAmt), "Sum")));

                        if (Convert.ToDouble(TransferSumInCurrency) < oPayments.Invoices.AppliedFC)
                        {
                            notEnoughFC = true;
                        }

                        invoicesSumApplied = invoicesSumApplied + Math.Abs(Convert.ToDecimal(oNewPayments.Invoices.AppliedFC));

                        TransferSumInCurrency = TransferSumInCurrency - Math.Abs(Convert.ToDecimal(minAmt));
                    }
                    else
                    {
                        double minAmt;
                        if (Math.Abs(oPayments.Invoices.SumApplied) > Convert.ToDouble(TransferSumInCurrency))
                        {
                            minAmt = Convert.ToDouble(TransferSumInCurrency);
                        }
                        else
                        {
                            minAmt = oPayments.Invoices.SumApplied;
                        }

                        oNewPayments.Invoices.SumApplied = Convert.ToDouble(CommonFunctions.roundAmountByGeneralSettings(Convert.ToDecimal(minAmt), "Sum"));

                        TransferSumInCurrency = TransferSumInCurrency - Math.Abs(Convert.ToDecimal(minAmt));
                    }

                    oNewPayments.Invoices.Add();
                }

                if ((DiffCurr == "Y") || ((!String.IsNullOrEmpty(oPayments.DocCurrency)) && (localCurrency != oPayments.DocCurrency)))
                {
                    if (TransferSumInCurrency > 0)
                    {
                        for (int j = 0; j < oNewPayments.Invoices.Count; j++)
                        {
                            oNewPayments.Invoices.SetCurrentLine(j);

                            if (oNewPayments.Invoices.DocEntry == 0)
                            {
                                continue;
                            }

                            //double invoiceBalance = APInvoice.GetInvoiceBalanceFC( oNewPayments.Invoices.DocEntry);
                            var dtFoundRow = dtInvoicesSorted.AsEnumerable().Where(x => (int)x["InvoiceDocEntry"] == oNewPayments.Invoices.DocEntry &&
                                                                     (int)x["InvType"] == (int)oNewPayments.Invoices.InvoiceType &&
                                                                     (int)x["InstlmntID"] == (int)oNewPayments.Invoices.InstallmentId

                                                                     ).FirstOrDefault();

                            double invoiceBalance = (double)dtFoundRow["OpenAmountFC"];
                            if (invoiceBalance <= 0)
                                continue;

                            if (invoiceBalance > 0)
                            {
                                TransferSumInCurrency = TransferSumInCurrency + Convert.ToDecimal(oNewPayments.Invoices.AppliedFC);

                                double minAmt;
                                if ((Math.Abs(oNewPayments.Invoices.AppliedFC) + invoiceBalance) > Convert.ToDouble(TransferSumInCurrency))
                                {
                                    minAmt = Convert.ToDouble(TransferSumInCurrency);
                                }
                                else
                                {
                                    minAmt = oNewPayments.Invoices.AppliedFC + invoiceBalance;
                                }

                                invoicesSumApplied = invoicesSumApplied - Math.Abs(Convert.ToDecimal(oNewPayments.Invoices.AppliedFC));

                                oNewPayments.Invoices.AppliedFC = Math.Abs(Convert.ToDouble(CommonFunctions.roundAmountByGeneralSettings(Convert.ToDecimal(minAmt), "Sum")));

                                invoicesSumApplied = invoicesSumApplied + Math.Abs(Convert.ToDecimal(oNewPayments.Invoices.AppliedFC));
                                //gamoklebamde sxva mnishvnelobaa da sachiroa gamoklebac da mimatebac

                                TransferSumInCurrency = TransferSumInCurrency - Math.Abs(Convert.ToDecimal(minAmt));

                            }

                        }
                    }
                }

                CommonFunctions.StartTransaction();

                //DiffCurr == "N" ესე იგი გადახდა ლარში არ არის.
                if (notEnoughFC && (bpCurrency != "##"))
                {
                    if (invoicesSumApplied != 0)
                    {
                        correctedRate = Convert.ToDouble(Convert.ToDecimal(oPayments.TransferSum) / invoicesSumApplied);
                    }

                    if (((!String.IsNullOrEmpty(oPayments.DocCurrency)) && (localCurrency != oPayments.DocCurrency)))
                    {
                        if (correctedRate != oNewPayments.DocRate)
                        {
                            oNewPayments.DocRate = correctedRate;
                        }
                    }
                }


                if (oNewPayments.LocalCurrency == SAPbobsCOM.BoYesNoEnum.tYES)
                {
                    oNewPayments.TransferSum = oPayments.TransferSum;
                }
                else
                {
                    oNewPayments.TransferSum = oDataTable.GetValue("Amount", rowNum - 1);
                }

                if (oPayments.Cancel() == 0)
                {
                    if (oNewPayments.Add() == 0)
                    {
                        if (Program.oCompany.InTransaction)
                        {
                            CommonFunctions.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                            //RowsWithCorrectedDates.Add(rowNum);
                        }
                        else
                        {
                            CommonFunctions.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                            Program.oCompany.GetLastError(out errCode, out errMsg);
                            Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("ErrorDescription") + " : " + errMsg + "! " + BDOSResources.getTranslate("Code") + " : " + errCode + "! ");

                        }
                    }
                    else
                    {
                        CommonFunctions.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                        Program.oCompany.GetLastError(out errCode, out errMsg);
                        Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("ErrorDescription") + " : " + errMsg + "! " + BDOSResources.getTranslate("Code") + " : " + errCode + "! ");

                    }
                }

            }

            return true;
        }

        public static void addMenus()
        {

            SAPbouiCOM.MenuItem menuItem;
            SAPbouiCOM.MenuItem fatherMenuItem;
            SAPbouiCOM.MenuCreationParams oCreationPackage;

            try
            {
                fatherMenuItem = Program.uiApp.Menus.Item("43537");
                // Add a pop-up menu item
                oCreationPackage = (SAPbouiCOM.MenuCreationParams)Program.uiApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                oCreationPackage.Checked = false;
                oCreationPackage.Enabled = true;
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "BDOSInternetBankingForm";
                oCreationPackage.String = BDOSResources.getTranslate("InternetBanking");
                oCreationPackage.Position = fatherMenuItem.SubMenus.Count - 1;

                menuItem = fatherMenuItem.SubMenus.AddEx(oCreationPackage);
            }
            catch
            {

            }
        }

        public static void uiApp_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_CLOSE && pVal.BeforeAction == true)
            {
                int answer = Program.uiApp.MessageBox(BDOSResources.getTranslate("CloseFormIntBank") + "?", 1, BDOSResources.getTranslate("Yes"), BDOSResources.getTranslate("No"), "");

                if (answer != 1)
                {
                    BubbleEvent = false;
                }
            }

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                try
                {
                    SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
                    {
                        if ((pVal.ItemUID == "autoPay") && !pVal.BeforeAction)
                        {
                            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("exportMTR").Specific;
                            int headRow = oMatrix.GetNextSelectedRow();

                            oForm.Freeze(true);

                            SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("exportMTR");
                            bool automaticPaymentInternetBanking = (oForm.DataSources.UserDataSources.Item("autoPay").ValueEx == "Y");

                            for (int i = 0; i < oDataTable.Rows.Count; i++)
                            {
                                string transactionType = oDataTable.GetValue("TransactionType", i);

                                bool selected = oDataTable.GetValue("CheckBox", i) == "Y" ? true : false;

                                if (selected == true && transactionType == OperationTypeFromIntBank.TransferFromBP.ToString())
                                {
                                    int rowNum = i + 1;
                                    string expression = "LineNumExportMTR = '" + rowNum + "'";
                                    DataRow[] foundRows;
                                    foundRows = TableExportMTRForDetail.Select(expression);

                                    if (foundRows.Length == 0)
                                    {
                                        decimal amount = Convert.ToDecimal(oDataTable.GetValue("Amount", i));

                                        if (automaticPaymentInternetBanking)
                                        {
                                            oDataTable.SetValue("AddDownPaymentAmount", i, Convert.ToDouble(amount));
                                            oDataTable.SetValue("PaymentOnAccount", i, 0);
                                        }
                                        else
                                        {
                                            oDataTable.SetValue("AddDownPaymentAmount", i, 0);
                                            oDataTable.SetValue("PaymentOnAccount", i, Convert.ToDouble(amount));
                                        }
                                    }
                                }
                            }
                            oMatrix.LoadFromDataSource();
                            oForm.Update();

                            if (headRow >= 1)
                            {
                                oMatrix.SelectRow(headRow, true, false);
                            }
                            oForm.Freeze(false);
                        }

                        else if (!pVal.BeforeAction && (pVal.ItemUID == "exportMTR" || pVal.ItemUID == "importMTR"))
                        {
                            int row = pVal.Row;

                            oForm.Freeze(true);

                            SAPbouiCOM.Matrix oMatrixGoods = (SAPbouiCOM.Matrix)oForm.Items.Item(pVal.ItemUID).Specific;

                            if (oMatrixGoods.RowCount > 0)
                            {
                                oForm.Freeze(false);

                                if (row <= oMatrixGoods.RowCount && row >= 1)
                                {
                                    if (pVal.ColUID != "CheckBox")
                                    {
                                        oMatrixGoods.SelectRow(row, true, false);
                                    }
                                }
                                oForm.Freeze(true);
                            }

                            oForm.Freeze(false);
                        }

                        else if (pVal.ItemUID == "FolderImp")
                        {
                            if (pVal.BeforeAction)
                            {
                                oForm.PaneLevel = 1;
                                setVisibleFormItemsImport(oForm, out errorText);
                            }
                        }

                        else if (pVal.ItemUID == "FolderExp")
                        {
                            if (pVal.Before_Action)
                            {
                                oForm.PaneLevel = 2;
                            }
                            else
                            {
                                Dictionary<string, string> CompanyInfo = CommonFunctions.getCompanyInfo();

                                if (oForm.Items.Item("TBCOB").Specific.Selected == false && oForm.Items.Item("BOGOB").Specific.Selected == false)
                                {
                                    if (CompanyInfo["DflBnkCode"] == "TBCBGE22")
                                        oForm.Items.Item("TBCOB").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                    else if (CompanyInfo["DflBnkCode"] == "BAGAGE22")
                                        oForm.Items.Item("BOGOB").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                    else
                                        oForm.Items.Item("TBCOB").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                }
                            }
                        }

                        //else if (pVal.ItemUID == "importMTR")
                        //{
                        //    if (pVal.Before_Action)
                        //    {
                        //        oForm.PaneLevel = 2;
                        //    }
                        //    else
                        //    {
                        //        Dictionary<string, string> CompanyInfo = CommonFunctions.getCompanyInfo();

                        //        if (oForm.Items.Item("TBCOB").Specific.Selected == false && oForm.Items.Item("BOGOB").Specific.Selected == false)
                        //        {
                        //            if (CompanyInfo["DflBnkCode"] == "TBCBGE22")
                        //                oForm.Items.Item("TBCOB").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        //            else if (CompanyInfo["DflBnkCode"] == "BAGAGE22")
                        //                oForm.Items.Item("BOGOB").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        //            else
                        //                oForm.Items.Item("TBCOB").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        //        }
                        //    }
                        //}

                        else if (pVal.ItemUID == "TBCOB" || pVal.ItemUID == "BOGOB")
                        {
                            if (!pVal.BeforeAction)
                            {
                                setVisibleFormItemsExport(oForm, out errorText);
                                oForm.DataSources.UserDataSources.Item("accountE2").ValueEx = "";
                            }
                        }
                    }

                    else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                    {
                        SAPbouiCOM.IChooseFromListEvent oCFLEvento = (SAPbouiCOM.IChooseFromListEvent)pVal;
                        if (pVal.ItemUID == "accountE")
                        {
                            chooseFromListImport(oForm, oCFLEvento, pVal, out errorText);
                        }
                        else if (pVal.ItemUID == "accountE2" || pVal.ItemUID == "exportMTR" || pVal.ItemUID == "BDOSDefCfE")
                        {
                            chooseFromListExport(oForm, oCFLEvento, pVal);
                        }
                    }

                    else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK)
                    {
                        if (!pVal.BeforeAction)
                        {
                            if (pVal.ItemUID == "fillB")
                            {
                                fillImportMTR(oForm, out errorText);
                            }
                            else if (pVal.ItemUID == "checkB" || pVal.ItemUID == "unCheckB")
                            {
                                checkUncheckMTRImport(oForm, pVal.ItemUID, out errorText);
                            }
                            else if (pVal.ItemUID == "getDataB")
                            {
                                getData(oForm);
                            }
                            else if (pVal.ItemUID == "checkB2" || pVal.ItemUID == "unCheckB2")
                            {
                                checkUncheckMTRExport(oForm, pVal.ItemUID, out errorText);
                            }
                            else if (pVal.ItemUID == "exportMTR" && pVal.ColUID == "CheckBox")
                            {
                                OnCheckImportDocuments(oForm, pVal.Row, pVal.ColUID);
                            }
                            else if (pVal.ItemUID == "createDocB")
                            {
                                int answer = Program.uiApp.MessageBox(BDOSResources.getTranslate("CreatePaymentDocuments") + "?", 1, BDOSResources.getTranslate("Yes"), BDOSResources.getTranslate("No"), "");

                                if (answer == 1)
                                {
                                    List<string> infoList = createDocuments(oForm);
                                    for (int i = 0; i < infoList.Count; i++)
                                    {
                                        if (string.IsNullOrEmpty(infoList[i]) == false)
                                        {
                                            Program.uiApp.SetStatusBarMessage(infoList[i], SAPbouiCOM.BoMessageTime.bmt_Short, false);
                                        }
                                    }
                                }
                                return;
                            }
                            else if (pVal.ItemUID == "exportMTR" && pVal.ColUID == "InDetail")
                            {
                                SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("exportMTR");

                                if (string.IsNullOrEmpty(oDataTable.GetValue("InDetail", pVal.Row - 1).ToString()))
                                {
                                    BubbleEvent = false;
                                    return;
                                }

                                bool selected = oDataTable.GetValue("CheckBox", pVal.Row - 1) == "Y" ? true : false;
                                if (selected == false)
                                {
                                    Program.uiApp.SetStatusBarMessage(BDOSResources.getTranslate("HighlightTheRowsToPerformAnOperation"), SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                    BubbleEvent = false;
                                    return;
                                }

                                CurrentRowExportMTRForDetail = pVal.Row;
                                DateTime docDate = oDataTable.GetValue("DocumentDate", pVal.Row - 1);

                                decimal amount = Convert.ToDecimal(oDataTable.GetValue("Amount", pVal.Row - 1));
                                string currency = Convert.ToString(oDataTable.GetValue("Currency", pVal.Row - 1));
                                decimal downPaymentAmount = Convert.ToDecimal(oDataTable.GetValue("DownPaymentAmount", pVal.Row - 1));
                                decimal invoicesAmount = Convert.ToDecimal(oDataTable.GetValue("InvoicesAmount", pVal.Row - 1));
                                decimal paymentOnAccount = Convert.ToDecimal(oDataTable.GetValue("PaymentOnAccount", pVal.Row - 1));

                                decimal addDPAmount = Convert.ToDecimal(oDataTable.GetValue("AddDownPaymentAmount", pVal.Row - 1));

                                decimal docRateIN = Convert.ToDecimal(oDataTable.GetValue("DocRateIN", pVal.Row - 1));

                                string transactionType = oDataTable.GetValue("TransactionType", pVal.Row - 1);
                                string cardType = null;
                                if (transactionType == OperationTypeFromIntBank.TransferFromBP.ToString())
                                {
                                    cardType = "C";
                                }
                                else if (transactionType == OperationTypeFromIntBank.TransferToBP.ToString() || transactionType == OperationTypeFromIntBank.TreasuryTransferPaymentOrderIoBP.ToString())
                                {
                                    cardType = "S";
                                }
                                string partnerAccountNumber = oDataTable.GetValue("PartnerAccountNumber", pVal.Row - 1);
                                string partnerCurrency = oDataTable.GetValue("PartnerCurrency", pVal.Row - 1);
                                string partnerTaxCode = oDataTable.GetValue("PartnerTaxCode", pVal.Row - 1);
                                string blnkAgr = oDataTable.GetValue("BlnkAgr", pVal.Row - 1);

                                SAPbobsCOM.Recordset oRecordSet = CommonFunctions.getBPBankInfo(partnerAccountNumber + partnerCurrency, partnerTaxCode, cardType);
                                if (oRecordSet == null)
                                {
                                    errorText = BDOSResources.getTranslate("CouldNotFindBusinessPartner") + "! " + BDOSResources.getTranslate("Account") + " \"" + partnerAccountNumber + partnerCurrency + "\"";
                                    if (string.IsNullOrEmpty(partnerTaxCode) == false)
                                    {
                                        errorText = errorText + ", " + BDOSResources.getTranslate("Tin") + " \"" + partnerTaxCode + "\"! ";
                                    }
                                    else errorText = errorText + "! ";

                                    Program.uiApp.SetStatusBarMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                    return;
                                }

                                string cardCode = oRecordSet.Fields.Item("CardCode").Value;
                                string cardName = oRecordSet.Fields.Item("CardName").Value;
                                string BPCurrency = oRecordSet.Fields.Item("Currency").Value;

                                SAPbouiCOM.Form oFormInternetBankingDocuments;

                                BDOSInternetBankingDocuments.automaticPaymentInternetBanking = (oForm.DataSources.UserDataSources.Item("autoPay").ValueEx == "Y");
                                BDOSInternetBankingDocuments.createForm(oForm, docDate, cardCode, cardName, BPCurrency, amount, currency, downPaymentAmount, invoicesAmount, paymentOnAccount, addDPAmount, docRateIN, out oFormInternetBankingDocuments, out errorText);
                                BDOSInternetBankingDocuments.fillInvoicesMTR(oFormInternetBankingDocuments, blnkAgr, out errorText);
                            }
                            else if (pVal.ItemUID == "syncDate")
                            {
                                if (SynchronizePayments(oForm))
                                {
                                    getData(oForm);
                                }
                            }
                        }
                    }

                    else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_COMBO_SELECT)
                    {
                        oForm.Freeze(true);

                        if (pVal.ItemUID == "listView" && !pVal.BeforeAction)
                        {
                            setVisibleFormItemsMatrixColumns(oForm, out errorText);
                        }
                        else if (pVal.ItemUID == "imptTypeCB" && !pVal.BeforeAction)
                        {
                            setVisibleFormItemsImport(oForm, out errorText);
                        }
                        comboSelectImport(oForm, pVal);
                        comboSelectExport(oForm, pVal);
                        oForm.Update();
                        oForm.Freeze(false);
                    }

                    else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED)
                    {
                        if (pVal.ItemUID == "importMTR")
                        {
                            matrixColumnSetLinkedObjectTypeImport(oForm, pVal);
                        }
                        else if (pVal.ItemUID == "exportMTR")
                        {
                            matrixColumnSetLinkedObjectTypeExport(oForm, pVal);
                        }
                    }

                    else if (pVal.ItemUID == "exportMTR" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_GOT_FOCUS)
                    {
                        if (!pVal.BeforeAction)
                        {
                            if (pVal.ColUID == "BlnkAgr")
                            {
                                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(pVal.ItemUID).Specific;
                                blnkAgrOld = oMatrix.GetCellSpecific(pVal.ColUID, pVal.Row).Value;
                            }
                        }
                    }

                    else if (pVal.ItemUID == "exportMTR" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_LOST_FOCUS)
                    {
                        if (!pVal.BeforeAction)
                        {
                            if (pVal.ColUID == "BlnkAgr")
                            {
                                oForm.Freeze(true);
                                try
                                {
                                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(pVal.ItemUID).Specific;
                                    string blnkAgr = oMatrix.GetCellSpecific(pVal.ColUID, pVal.Row).Value;
                                    if (blnkAgr != blnkAgrOld && !string.IsNullOrEmpty(blnkAgrOld) && string.IsNullOrEmpty(blnkAgr))
                                    {
                                        int rowIndex = pVal.Row;

                                        SAPbouiCOM.CheckBox oCheckBox = oMatrix.Columns.Item("UseBlaAgRt").Cells.Item(rowIndex).Specific;
                                        oCheckBox.Checked = false;

                                        setMTRCellEditableSetting(oForm, pVal.ItemUID, rowIndex);
                                        blnkAgrOld = null;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    blnkAgrOld = null;
                                    throw new Exception(ex.Message);
                                }
                                finally
                                {
                                    oForm.Freeze(false);
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Program.uiApp.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
            }
        }

        public static void OnCheckImportDocuments(SAPbouiCOM.Form oForm, int Row, string ColUID)
        {
            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("exportMTR").Specific));

            if (!SelectAllImportPressed)
            {
                oForm.Freeze(true);
                oMatrix.FlushToDataSource();
                oForm.Update();
                oForm.Freeze(false);
            }
            else
            {
                //oMatrix.LoadFromDataSource();
            }

            SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("exportMTR");
            string transactionType = oDataTable.GetValue("TransactionType", Row - 1);
            bool selected = oDataTable.GetValue(ColUID, Row - 1) == "Y" ? true : false;

            bool automaticPaymentInternetBanking = (oForm.DataSources.UserDataSources.Item("autoPay").ValueEx == "Y");

            if (selected == true && transactionType == OperationTypeFromIntBank.TransferFromBP.ToString())
            {
                decimal amount = Convert.ToDecimal(oDataTable.GetValue("Amount", Row - 1));

                if (automaticPaymentInternetBanking)
                {
                    oDataTable.SetValue("AddDownPaymentAmount", Row - 1, Convert.ToDouble(amount));
                    oDataTable.SetValue("PaymentOnAccount", Row - 1, 0);
                }
                else
                {
                    oDataTable.SetValue("AddDownPaymentAmount", Row - 1, 0);
                    oDataTable.SetValue("PaymentOnAccount", Row - 1, Convert.ToDouble(amount));
                }

                if (!SelectAllImportPressed)
                {
                    oForm.Freeze(true);
                    oMatrix.LoadFromDataSource();
                    oForm.Update();

                    oMatrix.SelectRow(Row, true, false);

                    oForm.Freeze(false);
                }
            }

            if (selected == false && transactionType == OperationTypeFromIntBank.TransferFromBP.ToString())
            {
                string expression = "LineNumExportMTR = '" + Row + "'";
                DataRow[] foundRows;
                foundRows = TableExportMTRForDetail.Select(expression);
                for (int i = 0; i < foundRows.Count(); i++)
                {
                    foundRows[i].Delete();
                }
                TableExportMTRForDetail.AcceptChanges();

                decimal amount = Convert.ToDecimal(oDataTable.GetValue("Amount", Row - 1));
                oDataTable.SetValue("DownPaymentAmount", Row - 1, 0);
                oDataTable.SetValue("InvoicesAmount", Row - 1, 0);

                if (automaticPaymentInternetBanking)
                {
                    oDataTable.SetValue("AddDownPaymentAmount", Row - 1, Convert.ToDouble(amount));
                    oDataTable.SetValue("PaymentOnAccount", Row - 1, 0);
                }
                else
                {
                    oDataTable.SetValue("AddDownPaymentAmount", Row - 1, 0);
                    oDataTable.SetValue("PaymentOnAccount", Row - 1, Convert.ToDouble(amount));
                }

                oDataTable.SetValue("DocRateIN", Row - 1, 0);


                if (!SelectAllImportPressed)
                {
                    oForm.Freeze(true);
                    oMatrix.LoadFromDataSource();
                    oForm.Update();

                    oMatrix.SelectRow(Row, true, false);

                    oForm.Freeze(false);
                }
            }
        }

        public static void createFolder(SAPbouiCOM.Form oForm, out string errorText)
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
                Dictionary<string, object> formItems = new Dictionary<string, object>();
                formItems = new Dictionary<string, object>();
                string itemName = "FolderImp";
                formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_FOLDER);
                formItems.Add("Bound", true);
                formItems.Add("TableName", "");
                formItems.Add("Alias", "FolderDS");
                formItems.Add("Width", 400);
                formItems.Add("UID", itemName);
                formItems.Add("Caption", BDOSResources.getTranslate("ExportIntBank")); //Import Data into Bank System
                formItems.Add("Pane", 1);
                formItems.Add("ValOn", "0");
                formItems.Add("ValOff", itemName);
                formItems.Add("AffectsFormMode", false);
                formItems.Add("Description", BDOSResources.getTranslate("ExportIntBank"));

                FormsB1.createFormItem(oForm, formItems, out errorText);

                formItems = new Dictionary<string, object>();
                itemName = "FolderExp";
                formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_FOLDER);
                formItems.Add("Bound", true);
                formItems.Add("TableName", "");
                formItems.Add("Alias", "FolderDS");
                formItems.Add("Width", 400);
                formItems.Add("UID", itemName);
                formItems.Add("Caption", BDOSResources.getTranslate("ImportIntBank")); //Get Data from Bank System
                formItems.Add("Pane", 2);
                formItems.Add("ValOn", "0");
                formItems.Add("ValOff", itemName);
                formItems.Add("GroupWith", "FolderImp");
                formItems.Add("AffectsFormMode", false);
                formItems.Add("Description", BDOSResources.getTranslate("ImportIntBank"));

                FormsB1.createFormItem(oForm, formItems, out errorText);
            }
            catch
            {
                string errMsg;
                int errCode;
                Program.oCompany.GetLastError(out errCode, out errMsg);
                Program.uiApp.StatusBar.SetSystemMessage(errMsg);
            }
        }
    }
}
