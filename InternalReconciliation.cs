using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
namespace BDO_Localisation_AddOn
{
    static partial class InternalReconciliation
    {
        public static string project;
        public static bool formGuess;
        public static int TransType;
        public static bool checker;
        public static string BP;
        public static void createFormItems(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            Dictionary<string, object> formItems = new Dictionary<string, object>();
            string itemName = "";
            formGuess = true;
            SAPbouiCOM.Item oItem = oForm.Items.Item("1470000041");
            int height = oItem.Height;
            int left_s = oItem.Left;
            int width = oItem.Width / 2;
            int left_e = oForm.Items.Item("120000015").Left;
            int top = oItem.Top;


            FormsB1.addChooseFromList(oForm, false, "63", "Prj_CFL");
            formItems = new Dictionary<string, object>();
            itemName = "PrjCodeS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width);
            formItems.Add("Top", top + height + 5);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Project"));
            formItems.Add("LinkTo", "PrjCodeE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }

            formItems = new Dictionary<string, object>();
            itemName = "PrjCodeE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OPRJ");
            formItems.Add("Alias", "PrjCode");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width);
            formItems.Add("Top", top + height + 5);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("ChooseFromListUID", "Prj_CFL");
            formItems.Add("ChooseFromListAlias", "PrjCode");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }
            GC.Collect();
        }

        public static void uiApp_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (pVal.FormTypeEx == "120060805")
                {
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE || pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE)
                        return;

                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD && pVal.BeforeAction)
                    {
                        createFormItems(oForm, out errorText);
                    }
                    else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST && !pVal.BeforeAction)
                    {
                        SAPbouiCOM.IChooseFromListEvent oCFLEvento = (SAPbouiCOM.IChooseFromListEvent)pVal;
                        chooseFromList(oForm, pVal, oCFLEvento, ref BubbleEvent);
                        filterTable(oForm);
                        SAPbouiCOM.Item itm = oForm.Items.Item("120000015");
                        itm.Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        itm = oForm.Items.Item("PrjCodeE");
                        itm.Enabled = false;
                    }
                    else if (pVal.ItemUID == "120000001" && pVal.BeforeAction)
                    {
                        SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("120000039").Specific));
                        for (int i = 1; i <= oMatrix.RowCount; i++)
                        {
                            SAPbouiCOM.CheckBox oCheck = (SAPbouiCOM.CheckBox)oMatrix.Columns.Item("120000002").Cells.Item(i).Specific;
                            if (oCheck.Checked)
                            {
                                SAPbouiCOM.EditText oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("120000033").Cells.Item(i).Specific;
                                SAPbouiCOM.EditText transt = (SAPbouiCOM.EditText)oMatrix.Columns.Item("120000003").Cells.Item(i).Specific;
                                SAPbouiCOM.EditText bpText = (SAPbouiCOM.EditText)oForm.Items.Item("120000015").Specific;
                                TransType = Int32.Parse(transt.Value.ToString());
                                BP = bpText.Value.ToString();
                                project = oEdit.Value;
                            }
                        }
                    }
                    if ((pVal.EventType == SAPbouiCOM.BoEventTypes.et_GOT_FOCUS || pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)  && checker)
                    {
                        SAPbobsCOM.Recordset oRecordset = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        try
                        {
                            StringBuilder query = new StringBuilder();
                            query.Append("SELECT \"OITR\".\"ReconJEId\" \n");
                            query.Append("FROM \"ITR1\" \n");
                            query.Append("INNER JOIN \"OITR\" ON \"OITR\".\"ReconNum\" = \"ITR1\".\"ReconNum\" \n");
                            query.Append("WHERE \n");
                            query.Append("\"ITR1\".\"ShortName\" = '" + BP + "' \n");
                            query.Append("AND \"ITR1\".\"TransId\" = '" + TransType + "' \n");
                            query.Append("AND \"OITR\".\"IsCard\" = 'C' \n");
                            query.Append("AND \"OITR\".\"Canceled\" = 'N' \n");
                            oRecordset.DoQuery(query.ToString());

                            if (!oRecordset.EoF)
                            {
                                SAPbobsCOM.JournalEntries journalEntry = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);

                                string transId = oRecordset.Fields.Item("ReconJEId").Value.ToString();
                                int transID = Int32.Parse(transId);
                                if (journalEntry.GetByKey(transID))
                                {
                                    journalEntry.ProjectCode = project;
                                }
                                if (journalEntry.Update() != 0)
                                {
                                    Program.oCompany.GetLastError(out int errCode, out string errMsg);
                                    Program.uiApp.StatusBar.SetSystemMessage(errMsg, SAPbouiCOM.BoMessageTime.bmt_Short);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            throw new Exception(ex.Message);
                        }
                        finally
                        {
                            checker = false;
                            Marshal.ReleaseComObject(oRecordset);
                        }
                    }
                }
                else if (pVal.FormTypeEx == "0")
                {
                    if (pVal.ItemUID == "1" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && !pVal.BeforeAction)
                        if (formGuess)
                        {
                            checker = true;
                            formGuess = false;
                        }
                }
            }
        }


        public static void chooseFromList(SAPbouiCOM.Form oForm, SAPbouiCOM.ItemEvent pVal, SAPbouiCOM.IChooseFromListEvent oCFLEvento, ref bool bubbleEvent)
        {
            try
            {
                string sCFL_ID = oCFLEvento.ChooseFromListUID;
                SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(sCFL_ID);

                if (!pVal.BeforeAction)
                {
                    SAPbouiCOM.DataTable oDataTable = null;
                    oDataTable = oCFLEvento.SelectedObjects;

                    if (oDataTable != null)
                    {
                        if (sCFL_ID == "Prj_CFL")
                        {
                            string PrjCode = Convert.ToString(oDataTable.GetValue("PrjCode", 0));

                            try
                            {
                                oForm.Items.Item("PrjCodeE").Specific.Value = PrjCode;
                            }
                            catch { }
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

        public static void filterTable(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("120000039").Specific));
            for (int i = 1; i <= oMatrix.RowCount; i++)
            {
                string ColId = "120000033";
                SAPbouiCOM.EditText oEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item(ColId).Cells.Item(i).Specific;
                string sValue = oEdit.Value;
                if (sValue != oForm.Items.Item("PrjCodeE").Specific.value)
                {
                    oMatrix.DeleteRow(i);
                    i--;
                }
            }
        }
    }
}