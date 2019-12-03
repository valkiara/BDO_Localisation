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


        private static void UpdateRTM1Table(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.CheckBox oCheckBox;
            SAPbouiCOM.Matrix oMatrix;

            oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("18").Specific));

            int rowCount = oMatrix.RowCount;
            for (int row = 1; row <= rowCount; row++)
            {
                oCheckBox = oMatrix.Columns.Item("1").Cells.Item(row).Specific;
                bool checkedLine = oCheckBox.Checked;

                if (checkedLine)
                {
                    string bpCode = oMatrix.Columns.Item("2").Cells.Item(row).Specific.Value;

                }
            }
        }

        private static void UpdateJournalEntry(string memo)
        {
            string docType = "";
            int docNum = 0;
            string doc = "";
            string project = "";
            string prjColumn = "";
            string transId = "";
            string agrNo;
            string useBlaAgrRt = "N";
            string errorText;
            DateTime executionDate = DateTime.Today;
            decimal blaAgrRt = 0;
            decimal jdt1Rt = 0;

            SAPbobsCOM.Recordset oRecordSetJDT1 = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset oRecordSetDoc = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            StringBuilder queryJDT1 = new StringBuilder();
            queryJDT1.Append("select * \n");
            queryJDT1.Append("from JDT1 \n");
            queryJDT1.Append("where \"LineMemo\" = '" + memo + "'");

            oRecordSetJDT1.DoQuery(queryJDT1.ToString());

            while (!oRecordSetJDT1.EoF)
            {
                string ref3 = oRecordSetJDT1.Fields.Item("Ref3Line").Value;
                if (ref3 != "")
                {
                    StringBuilder queryDoc = new StringBuilder();

                    docType = ref3.Substring(ref3.IndexOf('/') + 1, 2);
                    docNum = Convert.ToInt32(ref3.Substring(ref3.LastIndexOf('/') + 1));
                    transId = oRecordSetJDT1.Fields.Item("TransId").Value.ToString();

                    prjColumn = "Project";

                    if (docType == "PU")
                    {
                        doc = "OPCH";
                    }
                    else if (docType == "IN")
                    {
                        doc = "OINV";
                    }
                    else if (docType == "PC")
                    {
                        doc = "ORPC";
                    }
                    else if (docType == "CN")
                    {
                        doc = "ORIN";
                    }
                    else if (docType == "CP")
                    {
                        doc = "OCPI";
                    }
                    else if (docType == "CS")
                    {
                        doc = "OCSI";
                    }
                    else if (docType == "JE")
                    {
                        doc = "OJDT";

                        queryDoc.Append("select " + doc + ".\"" + prjColumn + "\", \"AgrNo\" \n");
                        queryDoc.Append("from " + doc + " \n");
                        queryDoc.Append("where \"Number\"  = '" + docNum + "'");
                        goto shortCut;
                    }
                    else if (docType == "PS")
                    {
                        doc = "OVPM";
                        prjColumn = "PrjCode";
                    }
                    else if (docType == "RC")
                    {
                        doc = "ORCT";
                        prjColumn = "PrjCode";
                    }
                    else
                    {
                        Program.uiApp.StatusBar.SetSystemMessage("Document with type - " + docType + " not supported", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        continue;
                    }
                    
                    queryDoc.Append("select " + doc + ".\"" + prjColumn + "\" as \"Project\", \"AgrNo\", \"U_UseBlaAgRt\", \"DocDate\" \n");
                    queryDoc.Append("from " + doc + " \n");
                    queryDoc.Append("where \"DocNum\"  = '" + docNum + "'");

                    shortCut:
                    oRecordSetDoc.DoQuery(queryDoc.ToString());
                    if (!oRecordSetDoc.EoF)
                    {
                        project = oRecordSetDoc.Fields.Item("Project").Value;
                        agrNo = oRecordSetDoc.Fields.Item("agrNo").Value.ToString();
                        useBlaAgrRt = oRecordSetDoc.Fields.Item("U_UseBlaAgRt").Value;

                        if (project != "")
                        {
                            UpdateJournalEntryTable("JDT1", transId, "Project", project);
                        }
                        if (agrNo != "0")
                        {

                            if (useBlaAgrRt == "Y")
                            {
                                blaAgrRt = BlanketAgreement.GetBlAgremeentCurrencyRate(Convert.ToInt32(agrNo), executionDate, out errorText);
                                jdt1Rt = oRecordSetJDT1.Fields.Item("RevalRate").Value;

                                if (blaAgrRt < jdt1Rt)
                                {

                                }

                                else if (blaAgrRt > jdt1Rt)
                                {

                                }

                                else if (blaAgrRt == jdt1Rt)
                                {

                                }
                            }

                            UpdateJournalEntryTable("OJDT", transId, "AgrNo", agrNo);
                        }
                        
                    }
                }
                oRecordSetJDT1.MoveNext();
            }
        }

        private static void UpdateJournalEntryTable(string table, string transId, string colToUpdate,string Update)
        {
            string error;
            SAPbobsCOM.Recordset oRecordSetJDT = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            StringBuilder queryJDT = new StringBuilder();
            queryJDT.Append("select \"TransId\", \"" + colToUpdate + "\" \n");
            queryJDT.Append("from " + table + " \n");
            queryJDT.Append("where \"TransId\" = '" + transId + "'");

            oRecordSetJDT.DoQuery(queryJDT.ToString());

            while (!oRecordSetJDT.EoF)
            {
                SAPbobsCOM.JournalEntries oJounalEntry = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
                oJounalEntry.GetByKey(oRecordSetJDT.Fields.Item("TransId").Value);
                if(table == "JDT1")
                {
                    oJounalEntry.UserFields.Fields.Item("Project").Value = Update;
                }
                else if(table == "OJDT")
                {
                    oJounalEntry.UserFields.Fields.Item("AgrNo").Value = Update;
                }

                int updateCode = oJounalEntry.Update();

                if (updateCode != 0)
                {
                    Program.oCompany.GetLastError(out updateCode, out error);
                }

                Marshal.ReleaseComObject(oJounalEntry);

                oRecordSetJDT.MoveNext();
            }
        }


    }
}
