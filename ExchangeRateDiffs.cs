using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Data;
using System.Xml;
using System.Xml.Linq;
using System.Text.RegularExpressions;
using SAPbobsCOM;
using SAPbouiCOM;

namespace BDO_Localisation_AddOn
{
    static partial class ExchangeRateDiffs
    {
        private static void CreateFormItems(Form oForm)
        {
            #region Auto Project Checkbox

            var oItem = oForm.Items.Item("26");

            var height = oItem.Height;
            var top = oItem.Top - height - 3;
            var width = oItem.Width;
            var left = oItem.Left;

            var formItems = new Dictionary<string, object>();
            var itemName = "AutoPrjCH"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "UserDataSources");
            formItems.Add("TableName", "");
            formItems.Add("Alias", itemName);
            formItems.Add("Bound", true);
            formItems.Add("Type", BoFormItemTypes.it_CHECK_BOX);
            formItems.Add("DataType", BoDataType.dt_SHORT_TEXT);
            formItems.Add("Length", 1);
            formItems.Add("Left", left);
            formItems.Add("Width", width);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("AutoProject"));
            formItems.Add("ValOff", "N");
            formItems.Add("ValOn", "Y");
            formItems.Add("ValueEx", "Y");

            FormsB1.createFormItem(oForm, formItems, out var errorText);
            if (errorText != null)
            {
                Program.uiApp.StatusBar.SetSystemMessage(errorText, BoMessageTime.bmt_Short);
                return;
            }

            #endregion

            #region Auto Blanket Agreement Checkbox

            oItem = oForm.Items.Item("27");
            left = oItem.Left;

            formItems = new Dictionary<string, object>();
            itemName = "AutoAgrCH"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "UserDataSources");
            formItems.Add("TableName", "");
            formItems.Add("Alias", itemName);
            formItems.Add("Bound", true);
            formItems.Add("Type", BoFormItemTypes.it_CHECK_BOX);
            formItems.Add("DataType", BoDataType.dt_SHORT_TEXT);
            formItems.Add("Length", 1);
            formItems.Add("Left", left);
            formItems.Add("Width", width);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("AutoAgrNo"));
            formItems.Add("ValOff", "N");
            formItems.Add("ValOn", "Y");
            formItems.Add("ValueEx", "Y");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                Program.uiApp.StatusBar.SetSystemMessage(errorText, BoMessageTime.bmt_Short);
            }

            #endregion
        }

        public static void UiApp_ItemEvent(string formUid, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (pVal.EventType == BoEventTypes.et_FORM_UNLOAD) return;

            var oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

            if (pVal.EventType == BoEventTypes.et_FORM_LOAD && pVal.BeforeAction)
            {
                CreateFormItems(oForm);
            }

            else if (pVal.EventType == BoEventTypes.et_ITEM_PRESSED)
            {
                if (pVal.ItemUID == "1")
                {
                    if (pVal.BeforeAction)
                    {
                        EditText oRemark = oForm.Items.Item("5").Specific;
                        oRemark.Value = DateTime.Now.ToString();
                    }

                    else
                    {
                        GetDataForUpdate(oForm);
                    }
                }
            }

            else if (pVal.EventType == BoEventTypes.et_FORM_CLOSE)
            {
                if (pVal.BeforeAction)
                {
                    GetDataForUpdate(oForm);
                }
            }
        }

        private static void GetDataForUpdate(Form oForm)
        {
            EditText oRemark = oForm.Items.Item("5").Specific;
            var memo = oRemark.Value;

            if (memo == "") return;

            double dayRate = 0;
            var docCur = "";

            var oRecordSetJDT1 = (Recordset) Program.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            var oRecordSetDoc = (Recordset) Program.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            var queryJDT1 = new StringBuilder();
            queryJDT1.Append(
                "select OJDT.\"BaseTrans\", OJDT.\"RevSource\", JDT1.\"Ref3Line\", JDT1.\"TransId\", JDT1.\"Debit\", JDT1.\"Credit\" \n");
            queryJDT1.Append("from JDT1 join OJDT on JDT1.\"TransId\" = OJDT.\"TransId\"\n");
            queryJDT1.Append("where \"LineMemo\" = '" + memo + "'");

            oRecordSetJDT1.DoQuery(queryJDT1
                .ToString()); // წამოიღებს გადაფასების ვიზარდიდან Journal Entry-ში შექმნილ დოკუმენტებს

            while (!oRecordSetJDT1.EoF)
            {
                string ref3 = oRecordSetJDT1.Fields.Item("Ref3Line").Value;
                int baseTrans = oRecordSetJDT1.Fields.Item("BaseTrans").Value;
                string revSource = oRecordSetJDT1.Fields.Item("RevSource").Value;

                if (ref3 != "") // ამით იფილტრება რომელი ანგარიშია JDT1-ში, ეს თუ ცარიელია ეგ ანგარიში არ გვაწყობს
                {
                    double debit = 0;
                    double credit = 0;

                    var queryDoc = new StringBuilder();

                    var docType = ref3.Substring(ref3.IndexOf('/') + 1, 2);
                    var docNum = Convert.ToInt32(ref3.Substring(ref3.LastIndexOf('/') + 1));
                    int transId = oRecordSetJDT1.Fields.Item("TransId").Value;

                    string docTable;
                    string docChildTable;

                    switch (docType) // რომელი დოკუმენტის მიხედვით იქმნება
                    {
                        case "PU":
                            docTable = "OPCH";
                            docChildTable = "PCH1";
                            break;
                        case "IN":
                            docTable = "OINV";
                            docChildTable = "INV1";
                            break;
                        case "PC":
                            docTable = "ORPC";
                            docChildTable = "RPC1";
                            break;
                        case "CN":
                            docTable = "ORIN";
                            docChildTable = "RIN1";
                            break;
                        case "CP":
                            docTable = "OCPI";
                            docChildTable = "CPI1";
                            break;
                        case "CS":
                            docTable = "OCSI";
                            docChildTable = "CSI1";
                            break;
                        case "JE":
                            docTable = "OJDT";
                            queryDoc.Append("select " + docTable + ".\"Project\", \"AgrNo\" \n");
                            queryDoc.Append("from " + docTable + " \n");
                            queryDoc.Append("where \"Number\"  = '" + docNum + "'");
                            goto shortCut;
                        case "PS":
                            docTable = "OVPM";

                            queryDoc.Append(
                                "select \"PrjCode\" as \"Project\", \"DocRate\", \"AgrNo\", \"U_UseBlaAgRt\", sum(\"OpenBalFc\") as \"Amount\" \n");
                            queryDoc.Append("from " + docTable + " \n");
                            queryDoc.Append("where \"DocNum\"  = '" + docNum + "'");
                            queryDoc.Append("group by \"PrjCode\", \"AgrNo\", \"DocRate\",\"U_UseBlaAgRt\"");
                            goto shortCut;
                        case "RC":
                            docTable = "ORCT";

                            queryDoc.Append(
                                "select \"PrjCode\" as \"Project\", \"DocRate\", \"AgrNo\", \"U_UseBlaAgRt\", sum(\"OpenBalFc\") as \"Amount\" \n");
                            queryDoc.Append("from " + docTable + " \n");
                            queryDoc.Append("where \"DocNum\"  = '" + docNum + "'");
                            queryDoc.Append("group by \"PrjCode\", \"AgrNo\", \"DocRate\", \"U_UseBlaAgRt\"");
                            goto shortCut;
                        default:
                            Program.uiApp.StatusBar.SetSystemMessage(
                                "Document with type - " + docType + " not supported", BoMessageTime.bmt_Short);
                            continue;
                    }

                    queryDoc.Append("select " + docTable + ".\"Project\", " + docTable + ".\"DocRate\", " + docTable +
                                    ".\"AgrNo\", " + docTable + ".\"U_UseBlaAgRt\", sum(" + docTable +
                                    ".\"DocTotalFC\") as \"Amount\" \n");
                    queryDoc.Append("from " + docTable + " inner join " + docChildTable + " on " + docTable +
                                    ".\"DocEntry\"= " + docChildTable + ".\"DocEntry\" \n");
                    queryDoc.Append("where \"DocNum\"  = '" + docNum + "'");
                    queryDoc.Append("group by " + docTable + ".\"Project\", " + docTable + ".\"AgrNo\", " + docTable +
                                    ".\"DocRate\", " + docTable + ".\"U_UseBlaAgRt\"");

                    shortCut:
                    oRecordSetDoc.DoQuery(queryDoc
                        .ToString()); // წამოიღებს იმ დოკუმენტს რის საფუძველზეც იქმნება Journal Entry
                    if (!oRecordSetDoc.EoF)
                    {
                        string project = oRecordSetDoc.Fields.Item("Project").Value;
                        int agrNo = oRecordSetDoc.Fields.Item("AgrNo").Value;

                        if (docTable != "OJDT") //kursis reinjebis gamoyeneba jer ar aris damatebuli jurnal entrishi
                        {
                            string useBlaAgrRt = oRecordSetDoc.Fields.Item("U_UseBlaAgRt").Value;

                            if (agrNo != 0)
                            {
                                if (useBlaAgrRt == "Y")
                                {
                                    var blaAgrRt = Convert.ToDouble(
                                        BlanketAgreement.GetBlAgremeentCurrencyRate(agrNo, out docCur, DateTime.Today));

                                    SBObob oSboBob = Program.oCompany.GetBusinessObject(BoObjectTypes.BoBridge);
                                    var rateRecordset = oSboBob.GetCurrencyRate(docCur, DateTime.Today);

                                    if (!rateRecordset.EoF)
                                    {
                                        dayRate = rateRecordset.Fields.Item("CurrencyRate").Value;
                                    }

                                    double docRate = oRecordSetDoc.Fields.Item("DocRate").Value;

                                    double amount =
                                        oRecordSetDoc.Fields.Item("Amount").Value; // * 1.18; //დღგ-ს გათვალისწინება

                                    debit = oRecordSetJDT1.Fields.Item("Debit").Value;
                                    credit = oRecordSetJDT1.Fields.Item("Credit").Value;

                                    if (blaAgrRt != dayRate)
                                    {
                                        if (credit != 0)
                                        {
                                            credit = amount * Math.Abs(blaAgrRt - docRate);
                                        }
                                        else if (debit != 0)
                                        {
                                            debit = amount * Math.Round(Math.Abs(blaAgrRt - docRate), 4);
                                        }
                                    }
                                    else
                                    {
                                        debit = 0;
                                        credit = 0;
                                    }

                                }
                            }
                        }

                        var updateProject = oForm.DataSources.UserDataSources.Item("AutoPrjCH").ValueEx;
                        var updateAgrNo = oForm.DataSources.UserDataSources.Item("AutoAgrCH").ValueEx;

                        UpdateJournalEntryTable(transId, memo, ref3, baseTrans, revSource, project, agrNo, credit,
                            debit, docCur, updateProject, updateAgrNo);
                    }
                }

                oRecordSetJDT1.MoveNext();
            }

            Marshal.ReleaseComObject(oRecordSetDoc);
            Marshal.ReleaseComObject(oRecordSetJDT1);

            oRemark.Value = "";
        }

        private static void UpdateJournalEntryTable(int transId, string memo, string ref3, int baseTrans,
            string revSource, string project, int agrNo, double credit, double debit, string docCur,
            string updateProject, string updateAgrNo)
        {
            #region Update Project and Blanket Agreement in old Journal Entry

            Recordset oRecordSetAgrNoAndProjectForOld = Program.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            try
            {
                var updateQuery = new StringBuilder();

                if (updateAgrNo == "Y")
                {
                    updateQuery.Append("UPDATE \"OJDT\" \n");
                    updateQuery.Append("SET \"AgrNo\" = '" + agrNo + "', \"U_BDOSAgrNo\" = '" + agrNo + "' \n");
                    updateQuery.Append("WHERE \"TransId\" = '" + transId + "';");
                    oRecordSetAgrNoAndProjectForOld.DoQuery(updateQuery.ToString());

                    updateQuery.Clear();
                }

                if (updateProject == "Y")
                {
                    updateQuery.Append("UPDATE \"JDT1\" \n");
                    updateQuery.Append("SET \"Project\" = '" + project + "' \n");
                    updateQuery.Append("WHERE \"TransId\" = '" + transId + "';");
                    oRecordSetAgrNoAndProjectForOld.DoQuery(updateQuery.ToString());
                }
            }
            catch (Exception ex)
            {
                Program.uiApp.StatusBar.SetSystemMessage(ex.Message, BoMessageTime.bmt_Short);
            }
            finally
            {
                Marshal.ReleaseComObject(oRecordSetAgrNoAndProjectForOld);
            }

            #endregion

            if (credit == 0 && debit == 0) return; // თუ თანხები ნოლია, მაკორექტირებელს აღარ ვქმნით

            #region Prepare old Journal Entry for new corrected journal entry

            //ვიღებთ საპის მიერ შექმნილ გატარებას, ვუცვლით თანხებს, გადაგვყავს ექსემელში და ამის მიხედვით ვქმნით მაკორექტირებელს, ბანძობაა ვიცი, 9ჯერ შეიცვალა ლოგიკა და ჩქარ-ჩქარა იყო დასაწერი (ფაცეპალმ)

            JournalEntries oJournalEntry = Program.oCompany.GetBusinessObject(BoObjectTypes.oJournalEntries);
            oJournalEntry.GetByKey(transId);

            Recordset oRecordSetTransType = Program.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            var queryTransType = new StringBuilder();
            queryTransType.Append("select \"TransType\" \n");
            queryTransType.Append("from OJDT \n");
            queryTransType.Append("where \"TransId\" = '" + transId + "'");

            oRecordSetTransType.DoQuery(queryTransType.ToString());
            if (!oRecordSetTransType.EoF)
            {
                oJournalEntry.Reference2 = oRecordSetTransType.Fields.Item("TransType").Value;
            }

            Marshal.ReleaseComObject(oRecordSetTransType);

            oJournalEntry.Reference = oJournalEntry.Original.ToString();

            #region Get sum amount of already transited entries

            Recordset oRecordSetOldEntriesAmount = Program.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            var amountColumn = "Debit";
            double oldAmount = 0;

            if (credit != 0)
            {
                amountColumn = "Credit";
            }

            var querySumAmount = new StringBuilder();
            querySumAmount.Append("SELECT SUM(\"JDT1\".\"" + amountColumn + "\") AS \"Amount\" \n");
            querySumAmount.Append("FROM \"JDT1\" \n");
            querySumAmount.Append("WHERE \"Ref3Line\" = '" + ref3 + "'");

            oRecordSetOldEntriesAmount.DoQuery(querySumAmount.ToString());
            if (!oRecordSetOldEntriesAmount.EoF)
            {
                oldAmount = oRecordSetOldEntriesAmount.Fields.Item("Amount").Value;
            }

            #endregion

            for (var line = 0; line < oJournalEntry.Lines.Count; line++)
            {
                oJournalEntry.Lines.SetCurrentLine(line);

                if (agrNo == 0) continue;
                
                if (oJournalEntry.Lines.AdditionalReference == ref3) //for bp
                {
                    if (credit != 0)
                    {
                        oJournalEntry.Lines.Credit = credit - oldAmount;
                    }
                    else if (debit != 0)
                    {
                        oJournalEntry.Lines.Debit = debit - oldAmount;
                    }
                }
                else if (oJournalEntry.Lines.AdditionalReference != ref3) //for second account
                {
                    oJournalEntry.Lines.FCCurrency = docCur;
                    if (credit != 0)
                    {
                        oJournalEntry.Lines.Debit = credit - oldAmount;
                    }
                    else if (debit != 0)
                    {
                        oJournalEntry.Lines.Credit = debit - oldAmount;
                    }
                }
            }

            Program.oCompany.XMLAsString = true;
            Program.oCompany.XmlExportType = BoXmlExportTypes.xet_ExportImportMode;

            var xml = oJournalEntry.GetAsXML();
            Marshal.ReleaseComObject(oJournalEntry);

            //ვშლით ისეთ ველებს რომლებიც არ შეიძლება რომ იყოს
            xml = Regex.Replace(xml,
                @"<\?.*?\?>|<DebitSys>.*?</DebitSys>|<CreditSys>.*?</CreditSys>|<JdtNum>.*?</JdtNum>|<SystemBaseAmount>.*?</SystemBaseAmount>|<VatAmount>.*?</VatAmount>|<SystemVatAmount>.*?</SystemVatAmount>|<GrossValue>.*?</GrossValue>",
                "");

            JournalEntries oJournalEntryNew = Program.oCompany.GetBusinessObject(BoObjectTypes.oJournalEntries);

            oJournalEntryNew.Browser.ReadXml(xml, 0);

            var updateCode = oJournalEntryNew.Add();
            Marshal.ReleaseComObject(oJournalEntryNew);

            #endregion

            if (updateCode != 0)
            {
                Program.oCompany.GetLastError(out updateCode, out var error);
                Program.uiApp.StatusBar.SetSystemMessage(error, BoMessageTime.bmt_Short);
            }
            else
            {
                #region Get new Journal Entry TransId

                var oRecordSetNewTransId = (Recordset) Program.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                var queryNewTransId = new StringBuilder();

                queryNewTransId.Append("select \"TransId\" \n");
                queryNewTransId.Append("from JDT1 \n");
                queryNewTransId.Append("where \"LineMemo\" = '" + memo + "'");

                oRecordSetNewTransId.DoQuery(queryNewTransId.ToString());

                var transIdNew = 0;

                while (!oRecordSetNewTransId.EoF)
                {
                    transIdNew = oRecordSetNewTransId.Fields.Item("TransId").Value;
                    if (transId == transIdNew)
                    {
                        transIdNew = 0;
                        oRecordSetNewTransId.MoveNext();
                    }
                    else
                    {
                        break;
                    }
                }

                Marshal.ReleaseComObject(oRecordSetNewTransId);

                #endregion

                #region Update new Journal Entry fields

                var oUpdateRecordSet = (Recordset) Program.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                var updateString = new StringBuilder();

                try
                {
                    updateString.Append("UPDATE \"OJDT\" \n");
                    updateString.Append("SET \"BaseTrans\" = '" + baseTrans + "', \"RevSource\" = '" + revSource + "'");

                    if (updateAgrNo == "Y")
                    {
                        updateString.Append(", \"AgrNo\" = '" + agrNo + "', \"U_BDOSAgrNo\" = '" + agrNo + "' \n");
                    }
                    else
                    {
                        updateString.Append(" \n");
                    }

                    updateString.Append("WHERE \"TransId\" = '" + transIdNew + "'; \n \n");
                    oUpdateRecordSet.DoQuery(updateString.ToString());

                    updateString.Clear();
                    updateString.Append("UPDATE \"JDT1\" \n");
                    updateString.Append("SET \"RevSource\" = '" + revSource + "' \n");
                    updateString.Append("WHERE \"TransId\" = '" + transIdNew + "';  \n \n");
                    oUpdateRecordSet.DoQuery(updateString.ToString());

                    updateString.Clear();
                    updateString.Append("UPDATE \"AJDT\" \n");
                    updateString.Append("SET \"BaseTrans\" = '" + baseTrans + "', \"RevSource\" = '" + revSource +
                                        "' \n");
                    updateString.Append("WHERE \"TransId\" = '" + transIdNew + "';  \n \n");
                    oUpdateRecordSet.DoQuery(updateString.ToString());

                    updateString.Clear();
                    updateString.Append("UPDATE \"AJD1\" \n");
                    updateString.Append("SET \"RevSource\" = '" + revSource + "' \n");
                    updateString.Append("WHERE \"TransId\" = '" + transIdNew + "';  \n \n");
                    oUpdateRecordSet.DoQuery(updateString.ToString());

                    updateString.Clear();
                    updateString.Append("UPDATE \"OBTF\" \n");
                    updateString.Append("SET \"BaseTrans\" = '" + baseTrans + "', \"RevSource\" = '" + revSource +
                                        "' \n");
                    updateString.Append("WHERE \"TransId\" = '" + transIdNew + "';  \n \n");
                    oUpdateRecordSet.DoQuery(updateString.ToString());

                    updateString.Clear();
                    updateString.Append("UPDATE \"BTF1\" \n");
                    updateString.Append("SET \"RevSource\" = '" + revSource + "' \n");
                    updateString.Append("WHERE \"TransId\" = '" + transIdNew + "';  \n \n");
                    oUpdateRecordSet.DoQuery(updateString.ToString());
                }
                catch (Exception ex)
                {
                    Program.uiApp.StatusBar.SetSystemMessage(ex.Message, BoMessageTime.bmt_Short);
                }

                finally
                {
                    Marshal.ReleaseComObject(oUpdateRecordSet);
                }

                #endregion
            }
        }
    }
}

