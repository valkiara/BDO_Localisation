using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Data;

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
            string errorText;
            decimal blaAgrRt = 0;
            decimal dayrt = 0;
            string docCur = "";
            string docChildTable = "";

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
                    decimal debit = 0;
                    decimal credit = 0;

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

                        queryDoc.Append("select \"PrjCode\" as \"Project\", \"AgrNo\", \"U_UseBlaAgRt\", sum(\"OpenBalFc\") as \"Amount\" \n");
                        queryDoc.Append("from " + docTable + " \n");
                        queryDoc.Append("where \"DocNum\"  = '" + docNum + "'");
                        queryDoc.Append("group by \"PrjCode\", \"AgrNo\", \"U_UseBlaAgRt\"");
                        goto shortCut;
                    }
                    else if (docType == "RC")
                    {
                        docTable = "ORCT";

                        queryDoc.Append("select \"PrjCode\" as \"Project\", \"AgrNo\", \"U_UseBlaAgRt\", sum(\"OpenBalFc\") as \"Amount\" \n");
                        queryDoc.Append("from " + docTable + " \n");
                        queryDoc.Append("where \"DocNum\"  = '" + docNum + "'");
                        queryDoc.Append("group by \"PrjCode\", \"AgrNo\", \"U_UseBlaAgRt\"");
                        goto shortCut;
                    }
                    else
                    {
                        Program.uiApp.StatusBar.SetSystemMessage("Document with type - " + docType + " not supported", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        continue;
                    }

                    queryDoc.Append("select " + docTable + ".\"Project\", " + docTable + ".\"AgrNo\", " + docTable + ".\"U_UseBlaAgRt\", sum(" + docChildTable + ".\"OpenSumFC\") as \"Amount\" \n");
                    queryDoc.Append("from " + docTable + " inner join " + docChildTable + " on " + docTable + ".\"DocEntry\"= " + docChildTable + ".\"DocEntry\" \n");
                    queryDoc.Append("where \"DocNum\"  = '" + docNum + "'");
                    queryDoc.Append("group by " + docTable + ".\"Project\", " + docTable + ".\"AgrNo\", " + docTable + ".\"U_UseBlaAgRt\"");

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
                                    blaAgrRt = BlanketAgreement.GetBlAgremeentCurrencyRate(agrNo, DateTime.Today, out errorText, out docCur);

                                    SAPbobsCOM.SBObob oSBOBob = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);
                                    SAPbobsCOM.Recordset RateRecordset = oSBOBob.GetCurrencyRate(docCur, DateTime.Today);

                                    if (!RateRecordset.EoF)
                                    {
                                        dayrt = Convert.ToDecimal(RateRecordset.Fields.Item("CurrencyRate").Value);
                                    }

                                    decimal amount = Convert.ToDecimal(oRecordSetDoc.Fields.Item("Amount").Value);

                                    debit = oRecordSetJDT1.Fields.Item("Debit").Value;
                                    credit = oRecordSetJDT1.Fields.Item("Credit").Value;

                                    if (blaAgrRt != dayrt)
                                    {
                                        if (credit != 0)
                                        {
                                            credit = amount * blaAgrRt;
                                        }
                                        else if (debit != 0)
                                        {
                                            debit = amount * blaAgrRt;
                                        }
                                    }

                                }
                            }
                        }
                        UpdateJournalEntryTable(transId, ref3, project, agrNo, credit, debit);
                    }
                }
                oRecordSetJDT1.MoveNext();
            }
        }

        private static void UpdateJournalEntryTable(int transId, string ref3, string project, int agrNo, decimal credit, decimal debit)
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
                            oJounalEntry.Lines.Credit = Convert.ToDouble(credit);
                        }
                        else if (debit != 0)
                        {
                            oJounalEntry.Lines.Debit = Convert.ToDouble(debit);
                        }
                    }
                    else if (oJounalEntry.Lines.AdditionalReference != ref3)  //for second account
                    {
                        if (credit != 0)
                        {
                            oJounalEntry.Lines.Debit = Convert.ToDouble(debit);
                        }
                        else if (debit != 0)
                        {
                            oJounalEntry.Lines.Credit = Convert.ToDouble(credit);
                        }
                    }
                }
            }

            int updateCode = oJounalEntry.Update();

            if (updateCode != 0)
            {
                Program.oCompany.GetLastError(out updateCode, out error);
                Program.uiApp.StatusBar.SetSystemMessage(error, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

            Marshal.ReleaseComObject(oJounalEntry);
        }
    }

}

