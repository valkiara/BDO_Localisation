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

namespace BDO_Localisation_AddOn
{
    static partial class ExchangeRateDiffs
    {

        public static void uiApp_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
                {
                    if (pVal.ItemUID == "1" && pVal.BeforeAction == true)
                    {
                        SAPbouiCOM.EditText oRemark = oForm.Items.Item("5").Specific;
                        oRemark.Value = DateTime.Now.ToString();
                    }
                    if (pVal.ItemUID == "1" && pVal.BeforeAction == false)
                    {
                        SAPbouiCOM.EditText oRemark = oForm.Items.Item("5").Specific;
                        UpdateJournalEntry(oRemark.Value);
                    }
                }

            }
        }


        //private static void UpdateRTM1Table(SAPbouiCOM.Form oForm)
        //{
        //    SAPbouiCOM.CheckBox oCheckBox;
        //    SAPbouiCOM.Matrix oMatrix;

        //    oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("18").Specific));

        //    int rowCount = oMatrix.RowCount;
        //    for (int row = 1; row <= rowCount; row++)
        //    {
        //        oCheckBox = oMatrix.Columns.Item("1").Cells.Item(row).Specific;
        //        bool checkedLine = oCheckBox.Checked;

        //        if (checkedLine)
        //        {
        //            string bpCode = oMatrix.Columns.Item("2").Cells.Item(row).Specific.Value;

        //        }
        //    }
        //}

        private static void UpdateJournalEntry(string memo)
        {
            string docType = "";
            int docNum = 0;
            string docTable = "";
            string project = "";
            int transId;
            int agrNo;
            string useBlaAgrRt = "N";
            double blaAgrRt = 0;
            double dayrt = 0;
            string docCur = "";
            string docChildTable = "";
            double docRate = 0;

            SAPbobsCOM.Recordset oRecordSetJDT1 = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset oRecordSetDoc = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            StringBuilder queryJDT1 = new StringBuilder();
            queryJDT1.Append("select \"Ref3Line\", \"TransId\", \"Debit\",\"Credit\" \n");
            queryJDT1.Append("from JDT1 \n");
            queryJDT1.Append("where \"LineMemo\" = '" + memo + "'");

            oRecordSetJDT1.DoQuery(queryJDT1.ToString());

            while (!oRecordSetJDT1.EoF)
            {

                string ref3 = oRecordSetJDT1.Fields.Item("Ref3Line").Value;
                if (ref3 != "")
                {
                    double debit = 0;
                    double credit = 0;

                    StringBuilder queryDoc = new StringBuilder();

                    docType = ref3.Substring(ref3.IndexOf('/') + 1, 2);
                    docNum = Convert.ToInt32(ref3.Substring(ref3.LastIndexOf('/') + 1));
                    transId = oRecordSetJDT1.Fields.Item("TransId").Value;

                    if (docType == "PU")
                    {
                        docTable = "OPCH";
                        docChildTable = "PCH1";
                    }
                    else if (docType == "IN")
                    {
                        docTable = "OINV";
                        docChildTable = "INV1";
                    }
                    else if (docType == "PC")
                    {
                        docTable = "ORPC";
                        docChildTable = "RPC1";
                    }
                    else if (docType == "CN")
                    {
                        docTable = "ORIN";
                        docChildTable = "RIN1";
                    }
                    else if (docType == "CP")
                    {
                        docTable = "OCPI";
                        docChildTable = "CPI1";
                    }
                    else if (docType == "CS")
                    {
                        docTable = "OCSI";
                        docChildTable = "CSI1";
                    }
                    else if (docType == "JE")
                    {
                        docTable = "OJDT";
                        queryDoc.Append("select " + docTable + ".\"Project\", \"AgrNo\" \n");
                        queryDoc.Append("from " + docTable + " \n");
                        queryDoc.Append("where \"Number\"  = '" + docNum + "'");
                        goto shortCut;
                    }
                    else if (docType == "PS")
                    {
                        docTable = "OVPM";

                        queryDoc.Append("select \"PrjCode\" as \"Project\", \"DocRate\", \"AgrNo\", \"U_UseBlaAgRt\", sum(\"OpenBalFc\") as \"Amount\" \n");
                        queryDoc.Append("from " + docTable + " \n");
                        queryDoc.Append("where \"DocNum\"  = '" + docNum + "'");
                        queryDoc.Append("group by \"PrjCode\", \"AgrNo\", \"DocRate\",\"U_UseBlaAgRt\"");
                        goto shortCut;
                    }
                    else if (docType == "RC")
                    {
                        docTable = "ORCT";

                        queryDoc.Append("select \"PrjCode\" as \"Project\", \"DocRate\", \"AgrNo\", \"U_UseBlaAgRt\", sum(\"OpenBalFc\") as \"Amount\" \n");
                        queryDoc.Append("from " + docTable + " \n");
                        queryDoc.Append("where \"DocNum\"  = '" + docNum + "'");
                        queryDoc.Append("group by \"PrjCode\", \"AgrNo\", \"DocRate\", \"U_UseBlaAgRt\"");
                        goto shortCut;
                    }
                    else
                    {
                        Program.uiApp.StatusBar.SetSystemMessage("Document with type - " + docType + " not supported", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        continue;
                    }

                    queryDoc.Append("select " + docTable + ".\"Project\", " + docTable + ".\"DocRate\", " + docTable + ".\"AgrNo\", " + docTable + ".\"U_UseBlaAgRt\", sum(" + docChildTable + ".\"OpenSumFC\") as \"Amount\" \n");
                    queryDoc.Append("from " + docTable + " inner join " + docChildTable + " on " + docTable + ".\"DocEntry\"= " + docChildTable + ".\"DocEntry\" \n");
                    queryDoc.Append("where \"DocNum\"  = '" + docNum + "'");
                    queryDoc.Append("group by " + docTable + ".\"Project\", " + docTable + ".\"AgrNo\", " + docTable + ".\"DocRate\", " + docTable + ".\"U_UseBlaAgRt\"");

                shortCut:
                    oRecordSetDoc.DoQuery(queryDoc.ToString());
                    if (!oRecordSetDoc.EoF)
                    {
                        project = oRecordSetDoc.Fields.Item("Project").Value;
                        agrNo = oRecordSetDoc.Fields.Item("agrNo").Value;
                        if (docTable != "OJDT") //kursis reinjebis gamoyeneba jer ar aris damatebuli jurnal entrishi
                        {
                            useBlaAgrRt = oRecordSetDoc.Fields.Item("U_UseBlaAgRt").Value;

                            if (agrNo != 0)
                            {
                                if (useBlaAgrRt == "Y")
                                {
                                    blaAgrRt = Convert.ToDouble(BlanketAgreement.GetBlAgremeentCurrencyRate(agrNo, out docCur, DateTime.Today));

                                    SAPbobsCOM.SBObob oSBOBob = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);
                                    SAPbobsCOM.Recordset RateRecordset = oSBOBob.GetCurrencyRate(docCur, DateTime.Today);

                                    if (!RateRecordset.EoF)
                                    {
                                        dayrt = RateRecordset.Fields.Item("CurrencyRate").Value;
                                    }

                                    docRate = oRecordSetDoc.Fields.Item("DocRate").Value;

                                    double amount = oRecordSetDoc.Fields.Item("Amount").Value * 1.18; //დღგ-ს გათვალისწინება

                                    debit = oRecordSetJDT1.Fields.Item("Debit").Value;
                                    credit = oRecordSetJDT1.Fields.Item("Credit").Value;

                                    if (blaAgrRt != dayrt)
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

                                }
                            }
                        }
                        UpdateJournalEntryTable(transId, ref3, project, agrNo, credit, debit, docCur);
                    }
                }
                oRecordSetJDT1.MoveNext();
            }
        }

        private static void UpdateJournalEntryTable(int transId, string ref3, string project, int agrNo, double credit, double debit, string docCur)
        {
            string error;
            SAPbobsCOM.JournalEntries oJounalEntry = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
            oJounalEntry.GetByKey(transId);


            for (int line = 0; line < oJounalEntry.Lines.Count; line++)
            {
                oJounalEntry.Lines.SetCurrentLine(line);

                if (!string.IsNullOrEmpty(project))
                {
                    oJounalEntry.Lines.ProjectCode = project;
                }
                if (agrNo != 0)
                {
                    //oJounalEntry.UserFields.Fields.Item("blanket agreementis columni").Value = agrNo;


                    if (oJounalEntry.Lines.AdditionalReference == ref3) //for bp
                    {
                        if (credit != 0)
                        {
                            oJounalEntry.Lines.Credit = credit;
                        }
                        else if (debit != 0)
                        {
                            oJounalEntry.Lines.Debit = debit;
                        }
                    }
                    else if (oJounalEntry.Lines.AdditionalReference != ref3)  //for second account
                    {
                        oJounalEntry.Lines.FCCurrency = docCur;
                        if (credit != 0)
                        {
                            oJounalEntry.Lines.Debit = credit;
                        }
                        else if (debit != 0)
                        {
                            oJounalEntry.Lines.Credit = debit;
                        }
                    }
                }
            }

            Program.oCompany.XMLAsString = true;
            Program.oCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode;

            string xml = oJounalEntry.GetAsXML();

            int updateCode = oJounalEntry.Cancel();
            if (updateCode != 0)
            {
                Program.oCompany.GetLastError(out updateCode, out error);
                Program.uiApp.StatusBar.SetSystemMessage(error, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            if(credit == 0 && debit == 0)
            {
                goto shortcut;
            }
            xml = Regex.Replace(xml, @"<\?.*?\?>|<DebitSys>.*?</DebitSys>|<CreditSys>.*?</CreditSys>|<JdtNum>.*?</JdtNum>|<SystemBaseAmount>.*?</SystemBaseAmount>|<VatAmount>.*?</VatAmount>|<SystemVatAmount>.*?</SystemVatAmount>|<GrossValue>.*?</GrossValue>", "");

            SAPbobsCOM.JournalEntries oJounalEntryNew = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);

            oJounalEntryNew.Browser.ReadXml(xml, 0);

            updateCode = oJounalEntryNew.Add();

            if (updateCode != 0)
            {
                Program.oCompany.GetLastError(out updateCode, out error);
                Program.uiApp.StatusBar.SetSystemMessage(error, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

            Marshal.ReleaseComObject(oJounalEntryNew);
            shortcut:
            Marshal.ReleaseComObject(oJounalEntry);
        }
    }

}

