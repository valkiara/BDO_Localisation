using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Globalization;
using System.Data;


namespace BDO_Localisation_AddOn
{
    static partial class Retirement
    {
        public static void createUserFields( out string errorText)
        {
            errorText = null;
            Dictionary<string, object> fieldskeysMap;

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSExpAct"); //ხარჯის ანაგარიში
            fieldskeysMap.Add("TableName", "ORTI");
            fieldskeysMap.Add("Description", "ExpenseAccount");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);
            fieldskeysMap.Add("Mandatory", SAPbobsCOM.BoYesNoEnum.tNO);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSDstRl1");
            fieldskeysMap.Add("TableName", "ORTI");
            fieldskeysMap.Add("Description", "Distr.Rule 1");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);
            fieldskeysMap.Add("Mandatory", SAPbobsCOM.BoYesNoEnum.tNO);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSDstRl2");
            fieldskeysMap.Add("TableName", "ORTI");
            fieldskeysMap.Add("Description", "Distr.Rule 2");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);
            fieldskeysMap.Add("Mandatory", SAPbobsCOM.BoYesNoEnum.tNO);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSDstRl3");
            fieldskeysMap.Add("TableName", "ORTI");
            fieldskeysMap.Add("Description", "Distr.Rule 3");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);
            fieldskeysMap.Add("Mandatory", SAPbobsCOM.BoYesNoEnum.tNO);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSDstRl4");
            fieldskeysMap.Add("TableName", "ORTI");
            fieldskeysMap.Add("Description", "Distr.Rule 4");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);
            fieldskeysMap.Add("Mandatory", SAPbobsCOM.BoYesNoEnum.tNO);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSDstRl5");
            fieldskeysMap.Add("TableName", "ORTI");
            fieldskeysMap.Add("Description", "Distr.Rule 5");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);
            fieldskeysMap.Add("Mandatory", SAPbobsCOM.BoYesNoEnum.tNO);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSPrjCod");
            fieldskeysMap.Add("TableName", "ORTI");
            fieldskeysMap.Add("Description", "Project");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);
            fieldskeysMap.Add("Mandatory", SAPbobsCOM.BoYesNoEnum.tNO);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

        }

        public static void createFormItems(  SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            Dictionary<string, object> formItems;
            string itemName = "";

            SAPbouiCOM.Item oItem = null;
            int pane = 3;

            //ხარჯები (ჩანართი)
            SAPbouiCOM.Item oFolder = oForm.Items.Item("1470000055");
            formItems = new Dictionary<string, object>();
            itemName = "Expenses";
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_FOLDER);
            formItems.Add("Left", oFolder.Left + oFolder.Width);
            formItems.Add("Width", oFolder.Width);
            formItems.Add("Top", oFolder.Top);
            formItems.Add("Height", oFolder.Height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Expenses"));
            formItems.Add("Pane", pane);
            formItems.Add("ValOn", "0");
            formItems.Add("ValOff", itemName);
            formItems.Add("GroupWith", "1470000055");
            formItems.Add("AffectsFormMode", false);
            formItems.Add("Description", BDOSResources.getTranslate("Expenses"));

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            oItem = oForm.Items.Item("1470000053");
            int top = oItem.Top;
            int height = 15;
            int left_s = 6;
            int left_e = 127;
            int width_s = 121;
            int width_e = 148;
            string objectType = "";
            bool multiSelection = false;

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "ExpActS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("ExpenseAccount"));
            formItems.Add("LinkTo", "ExpActE");
            //formItems.Add("FontSize", fontSize);
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            objectType = "1";
            string uniqueID_lf_Acct_CFL = "Acct_CFL";
            FormsB1.addChooseFromList( oForm, multiSelection, objectType, uniqueID_lf_Acct_CFL);
            //პირობის დადება ანგარიშის არჩევის სიაზე
            SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(uniqueID_lf_Acct_CFL);
            SAPbouiCOM.Conditions oCons = oCFL.GetConditions();
            SAPbouiCOM.Condition oCon = oCons.Add();
            oCon.Alias = "Postable";
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCon.CondVal = "Y"; //Active Account
            oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

            oCon = oCons.Add();
            oCon.Alias = "FrozenFor";
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCon.CondVal = "N"; //not inactive
            oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_NONE;

            oCFL.SetConditions(oCons);

            formItems = new Dictionary<string, object>();
            itemName = "ExpActE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "ORTI");
            formItems.Add("Alias", "U_BDOSExpAct");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("ChooseFromListUID", uniqueID_lf_Acct_CFL);
            formItems.Add("ChooseFromListAlias", "AcctCode");
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "ExpActLB"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            formItems.Add("Left", left_e - 20);
            formItems.Add("Top", top);
            formItems.Add("Height", 14);
            formItems.Add("UID", itemName);
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);
            formItems.Add("LinkTo", "ExpActE");
            formItems.Add("LinkedObjectType", objectType);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "ProjectS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Project"));
            formItems.Add("LinkTo", "ProjectE");
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            objectType = "63";
            string uniqueID_lf_Project = "Project_CFL";
            FormsB1.addChooseFromList( oForm, multiSelection, objectType, uniqueID_lf_Project);

            formItems = new Dictionary<string, object>();
            itemName = "ProjectE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "ORTI");
            formItems.Add("Alias", "U_BDOSPrjCod");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("ChooseFromListUID", uniqueID_lf_Project);
            formItems.Add("ChooseFromListAlias", "PrjCode");
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "ProjectLB"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            formItems.Add("Left", left_e - 20);
            formItems.Add("Top", top);
            formItems.Add("Height", 14);
            formItems.Add("UID", itemName);
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);
            formItems.Add("LinkTo", "ProjectE");
            formItems.Add("LinkedObjectType", objectType);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            Dictionary<string, string> activeDimensionsList = CommonFunctions.getActiveDimensionsList( out errorText);

            for (int i = 1; i <= activeDimensionsList.Count; i++)
            {
                top = top + height + 1;

                formItems = new Dictionary<string, object>();
                itemName = "DistrRul" + i + "S"; //10 characters
                formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                formItems.Add("Left", left_s);
                formItems.Add("Width", width_s);
                formItems.Add("Top", top);
                formItems.Add("Height", height);
                formItems.Add("UID", itemName);
                formItems.Add("Caption", activeDimensionsList[i.ToString()]);
                formItems.Add("LinkTo", "DistrRul" + i + "E");
                //formItems.Add("Visible", false);
                formItems.Add("FromPane", pane);
                formItems.Add("ToPane", pane);

                FormsB1.createFormItem(oForm, formItems, out errorText);
                if (errorText != null)
                {
                    return;
                }

                objectType = "62";
                string uniqueID_lf_DistrRule = "Rule_CFL" + i.ToString() + "A";
                FormsB1.addChooseFromList( oForm, multiSelection, objectType, uniqueID_lf_DistrRule);

                formItems = new Dictionary<string, object>();
                itemName = "DistrRul" + i + "E"; //10 characters
                formItems.Add("isDataSource", true);
                formItems.Add("DataSource", "DBDataSources");
                formItems.Add("TableName", "ORTI");
                formItems.Add("Alias", "U_BDOSDstRl" + i);
                formItems.Add("Bound", true);
                formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                formItems.Add("Left", left_e);
                formItems.Add("Width", width_e);
                formItems.Add("Top", top);
                formItems.Add("Height", height);
                formItems.Add("UID", itemName);
                formItems.Add("DisplayDesc", true);
                formItems.Add("ChooseFromListUID", uniqueID_lf_DistrRule);
                formItems.Add("ChooseFromListAlias", "OcrCode");
                //formItems.Add("Visible", false);
                formItems.Add("FromPane", pane);
                formItems.Add("ToPane", pane);

                FormsB1.createFormItem(oForm, formItems, out errorText);
                if (errorText != null)
                {
                    return;
                }

                formItems = new Dictionary<string, object>();
                itemName = "DstrRul" + i + "LB"; //10 characters
                formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                formItems.Add("Left", left_e - 20);
                formItems.Add("Top", top);
                formItems.Add("Height", 14);
                formItems.Add("UID", itemName);
                formItems.Add("FromPane", pane);
                formItems.Add("ToPane", pane);
                formItems.Add("LinkTo", "DistrRul" + i + "E");
                formItems.Add("LinkedObjectType", objectType);

                FormsB1.createFormItem(oForm, formItems, out errorText);
                if (errorText != null)
                {
                    return;
                }
            }

            //foreach (KeyValuePair<string, string> activeDim in activeDimensionsList)
            //{
            //    oForm.Items.Item("DistrRul" + activeDim.Key + "S").Visible = true;
            //    oForm.Items.Item("DistrRul" + activeDim.Key + "E").Visible = true;
            //    oForm.Items.Item("DstrRul" + activeDim.Key + "LB").Visible = true;
            //}

            GC.Collect();
        }

        public static void chooseFromList( SAPbouiCOM.Form oForm, SAPbouiCOM.IChooseFromListEvent oCFLEvento, SAPbouiCOM.ItemEvent pVal, out string errorText)
        {
            errorText = null;
            string sCFL_ID = oCFLEvento.ChooseFromListUID;
            int row = pVal.Row;

            if (pVal.BeforeAction == false)
            {
                try
                {
                    SAPbouiCOM.DataTable oDataTable = null;
                    oDataTable = oCFLEvento.SelectedObjects;

                    if (oDataTable != null)
                    {
                        if (sCFL_ID == "Acct_CFL")
                        {
                            string account = Convert.ToString(oDataTable.GetValue("AcctCode", 0));

                            try
                            {
                                SAPbouiCOM.EditText oEdit = oForm.Items.Item("ExpActE").Specific;
                                oEdit.Value = account;
                            }
                            catch { }
                        }
                        else if (sCFL_ID == "Project_CFL")
                        {
                            string prjCode = Convert.ToString(oDataTable.GetValue("PrjCode", 0));

                            try
                            {
                                SAPbouiCOM.EditText oEdit = oForm.Items.Item("ProjectE").Specific;
                                oEdit.Value = prjCode;
                            }
                            catch { }
                        }
                        else if (sCFL_ID.Substring(0, sCFL_ID.Length - 2) == "Rule_CFL")
                        {
                            SAPbouiCOM.DataTable oDataTableSelectedObjects = oCFLEvento.SelectedObjects;
                            string val = oDataTableSelectedObjects.GetValue("OcrCode", 0);

                            SAPbouiCOM.EditText oEdit = oForm.Items.Item("DistrRul" + sCFL_ID.Substring(sCFL_ID.Length - 2, 1) + "E").Specific;
                            oEdit.Value = val;
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
            else
            {
                if (sCFL_ID.Substring(0, sCFL_ID.Length - 2) == "Rule_CFL")
                {
                    string dimensionCode = sCFL_ID.Substring(sCFL_ID.Length - 2, 1);

                    SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(sCFL_ID);
                    SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();

                    string startDateStr = oForm.Items.Item("1470000019").Specific.Value;
                    DateTime DocDate = DateTime.TryParseExact(startDateStr, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DocDate) == false ? DateTime.Now : DocDate;

                    SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    string query = @"SELECT
	                                     ""OCR1"".""OcrCode"",
	                                     ""OOCR"".""DimCode"" 
                                    FROM ""OCR1"" 
                                    LEFT JOIN ""OOCR"" ON ""OCR1"".""OcrCode"" = ""OOCR"".""OcrCode"" 
                                    WHERE ""OOCR"".""DimCode"" = " + dimensionCode + @" AND ""ValidFrom"" <= '" + DocDate.ToString("yyyyMMdd") +
                                                                                     @"' AND (""ValidTo"" > '" + DocDate.ToString("yyyyMMdd") + @"' OR " + @" ""ValidTo"" IS NULL)";

                    try
                    {
                        oRecordSet.DoQuery(query);
                        int recordCount = oRecordSet.RecordCount;
                        int i = 1;

                        while (!oRecordSet.EoF)
                        {
                            SAPbouiCOM.Condition oCon = oCons.Add();
                            oCon.Alias = "OcrCode";
                            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                            oCon.CondVal = oRecordSet.Fields.Item("OcrCode").Value.ToString();
                            oCon.Relationship = (i == recordCount) ? SAPbouiCOM.BoConditionRelationship.cr_NONE : SAPbouiCOM.BoConditionRelationship.cr_OR;

                            i = i + 1;
                            oRecordSet.MoveNext();
                        }

                        //თუ არცერთი შეესაბამება ცარიელზე გავიდეს
                        if (oCons.Count == 0)
                        {
                            SAPbouiCOM.Condition oCon = oCons.Add();
                            oCon.Alias = "OcrCode";
                            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                            oCon.CondVal = "";
                        }

                        oCFL.SetConditions(oCons);
                    }
                    catch (Exception ex)
                    {
                        errorText = ex.Message;
                    }
                }
            }

        }

        public static void formDataLoad( SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            oForm.Freeze(true);

            try
            {
                setVisibleFormItems(oForm, out errorText);
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
                oForm.Freeze(false);
                GC.Collect();
            }
        }

        public static void setVisibleFormItems(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            try
            {
                string docEntry = oForm.DataSources.DBDataSources.Item("ORTI").GetValue("DocEntry", 0).Trim();
                bool docEntryIsEmpty = string.IsNullOrEmpty(docEntry);
                oForm.Items.Item("1470000022").Click(SAPbouiCOM.BoCellClickType.ct_Regular); //მისაწვდომობის შეზღუდვისთვის

                bool Scrapping = oForm.DataSources.DBDataSources.Item("ORTI").GetValue("DocType", 0).Trim() == "SC";
                oForm.Items.Item("Expenses").Visible = Scrapping;

                SAPbouiCOM.Item oItem = null;

                oItem = oForm.Items.Item("ExpActE");
                oItem.Enabled = (docEntryIsEmpty == true);
                oItem.Visible = (Scrapping & oForm.PaneLevel == 3);
                oForm.Items.Item("ExpActS").Visible = (Scrapping & oForm.PaneLevel == 3);

                bool Prj = oForm.DataSources.DBDataSources.Item("ORTI").GetValue("PrjSmarz", 0) == "Y";
                oItem = oForm.Items.Item("ProjectE");
                oItem.Enabled = (docEntryIsEmpty == true & Prj);
                oItem.Specific.ChooseFromListUID = "Project_CFL";
                oItem.Specific.ChooseFromListAlias = "PrjCode";
                oItem.Visible = (Scrapping & oForm.PaneLevel == 3);
                oForm.Items.Item("ProjectS").Visible = (Scrapping & oForm.PaneLevel == 3);

                bool DstRl = oForm.DataSources.DBDataSources.Item("ORTI").GetValue("DstRlSmarz", 0) == "Y";
                for (int i = 1; i <= 5; i++)
                {
                    oItem = oForm.Items.Item("DistrRul" + i + "E");
                    oItem.Enabled = (docEntryIsEmpty == true);
                    oItem.Specific.ChooseFromListUID = "Rule_CFL" + i.ToString() + "A";
                    oItem.Specific.ChooseFromListAlias = "OcrCode";
                    oItem.Visible = (Scrapping & oForm.PaneLevel == 3);
                    oForm.Items.Item("DistrRul" + i + "S").Visible = (Scrapping & oForm.PaneLevel == 3);
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

        public static void uiApp_FormDataEvent(  ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(BusinessObjectInfo.FormTypeEx, Program.currentFormCount);

            if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD)
            {
                if (BusinessObjectInfo.BeforeAction == true)
                {
                    SAPbouiCOM.DBDataSource DocDBSourceOCRD = oForm.DataSources.DBDataSources.Item(0);

                    // შემოწმება
                    if (oForm.DataSources.DBDataSources.Item("ORTI").GetValue("DocType", 0).Trim() == "SC")
                    {
                        if (oForm.DataSources.DBDataSources.Item("ORTI").GetValue("U_BDOSExpAct", 0).Trim() == "")
                        {
                            Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("ExpenseAccount") + " " + BDOSResources.getTranslate("YouCantLeaveEmpty"));
                            Program.uiApp.MessageBox(BDOSResources.getTranslate("OperationUnsuccesfullSeeLog"));
                            BubbleEvent = false;
                        }
                    }

                }


                if (BusinessObjectInfo.ActionSuccess & BusinessObjectInfo.BeforeAction == false)
                {
                    //დოკუმენტის გატარების დროს გატარდეს ბუღლტრული გატარება
                    if (oForm.DataSources.DBDataSources.Item("ORTI").GetValue("DocType", 0).Trim() == "SC")
                    {
                        SAPbouiCOM.DBDataSource DocDBSourceTAXP = oForm.DataSources.DBDataSources.Item(0);

                        string DocEntry = DocDBSourceTAXP.GetValue("DocEntry", 0);
                        string DocNum = DocDBSourceTAXP.GetValue("DocNum", 0);
                        decimal DocRate = 0;
                        string DocCurr = "";
                        DateTime DocDate = DateTime.ParseExact(DocDBSourceTAXP.GetValue("DocDate", 0), "yyyyMMdd", CultureInfo.InvariantCulture);

                        JrnEntry( DocEntry, DocNum, DocDate, DocRate, DocCurr, out errorText);
                        if (errorText != null)
                        {
                            Program.uiApp.MessageBox(errorText);
                            BubbleEvent = false;
                        }
                    }
                }
            }

            if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE & BusinessObjectInfo.BeforeAction == false & BusinessObjectInfo.ActionSuccess == true)
            {
                if (Program.cancellationTrans == true & Program.canceledDocEntry != 0)
                {
                    cancellation( oForm, Program.canceledDocEntry, out errorText);
                    Program.canceledDocEntry = 0;
                }
            }

            if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD & BusinessObjectInfo.BeforeAction == false)
            {
                formDataLoad( oForm, out errorText);
            }
        }

        public static void cancellation(  SAPbouiCOM.Form oForm, int docEntry, out string errorText)
        {
            errorText = null;

            try
            {
                JournalEntry.cancellation( oForm, docEntry, "1470000094", out errorText);
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

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (pVal.ItemUID == "Expenses" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED & pVal.BeforeAction == true)
                {
                    oForm.Freeze(true);
                    oForm.PaneLevel = 3;
                    setVisibleFormItems(oForm, out errorText);
                    oForm.Freeze(false);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
                {
                    if ((pVal.ItemUID == "1470000042" || pVal.ItemUID == "1470000043") && pVal.BeforeAction == false)
                    {
                        oForm.Freeze(true);
                        setVisibleFormItems(oForm, out errorText);
                        oForm.Freeze(false);
                    }
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_COMBO_SELECT & pVal.ItemUID == "1470000040" & pVal.BeforeAction == false)
                {
                    oForm.Freeze(true);
                    setVisibleFormItems(oForm, out errorText);
                    oForm.Freeze(false);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                {
                    string CurrentitemName = pVal.ItemUID;
                    if (CurrentitemName == "ProjectE" || CurrentitemName == "ExpActE" || CurrentitemName.Substring(0, 8) == "DistrRul")
                    {
                        {
                            SAPbouiCOM.ChooseFromListEvent oCFLEvento = null;
                            oCFLEvento = ((SAPbouiCOM.ChooseFromListEvent)(pVal));

                            chooseFromList( oForm, oCFLEvento, pVal, out errorText);
                        }
                    }
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD & pVal.BeforeAction == true)
                {
                    createFormItems( oForm, out errorText);
                    formDataLoad( oForm, out errorText);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_RESIZE && pVal.BeforeAction == false)
                {
                    oForm.Freeze(true);
                    resizeForm( oForm, out errorText);
                    oForm.Freeze(false);
                }
            }
        }

        public static void resizeForm(  SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            try
            {
                reArrangeFormItems( oForm);
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
        }

        public static void reArrangeFormItems( SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.Item oItem = null;

            oItem = oForm.Items.Item("1470000053");
            int top = oItem.Top;
            int height = 15;

            top = top + height + 1;

            oItem = oForm.Items.Item("ExpActS");
            oItem.Top = top;
            oItem = oForm.Items.Item("ExpActE");
            oItem.Top = top;
            oItem = oForm.Items.Item("ExpActLB");
            oItem.Top = top;

            top = top + height + 1;

            oItem = oForm.Items.Item("ProjectS");
            oItem.Top = top;
            oItem = oForm.Items.Item("ProjectE");
            oItem.Top = top;
            oItem = oForm.Items.Item("ProjectLB");
            oItem.Top = top;

            string errorText = "";
            Dictionary<string, string> activeDimensionsList = CommonFunctions.getActiveDimensionsList( out errorText);

            for (int i = 1; i <= activeDimensionsList.Count; i++)
            {
                top = top + height + 1;

                oItem = oForm.Items.Item("DistrRul" + i + "S");
                oItem.Top = top;
                oItem = oForm.Items.Item("DistrRul" + i + "E");
                oItem.Top = top;
                oItem = oForm.Items.Item("DstrRul" + i + "LB");
                oItem.Top = top;
            }
        }

        public static void JrnEntry( string DocEntry, string DocNum, DateTime DocDate, Decimal rate, string currency, out string errorText)
        {
            errorText = null;

            try
            {
                DataTable jeLines = JournalEntry.JournalEntryTable();
                DataRow jeLinesRow = null;

                DataTable reLines = ProfitTax.ProfitTaxTable();

                SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                string query = @"SELECT SUM(""JDT1"".""Debit"") AS ""Amount"", 
	                           ""ORTI"".""PrjSmarz"",
	                           ""ORTI"".""DstRlSmarz"",
	                           ""JDT1"".""Account"" AS ""Account"",
	                           ""JDT1"".""ProfitCode"" AS ""DistrRule1"",
	                           ""JDT1"".""OcrCode2"" AS ""DistrRule2"",
	                           ""JDT1"".""OcrCode3"" AS ""DistrRule3"",
	                           ""JDT1"".""OcrCode4"" AS ""DistrRule4"",
	                           ""JDT1"".""OcrCode5"" AS ""DistrRule5"",
	                           ""JDT1"".""Project"" AS ""PrjCode"",
                               ""OACT"".""ActType"" AS ""BDOSExpAccount_ActType"",
	                           ""ORTI"".""U_BDOSExpAct"" AS ""BDOSExpAccount"",
	                           ""ORTI"".""U_BDOSDstRl1"" AS ""BDOSDistrRule1"",
	                           ""ORTI"".""U_BDOSDstRl2"" AS ""BDOSDistrRule2"",
	                           ""ORTI"".""U_BDOSDstRl3"" AS ""BDOSDistrRule3"",
	                           ""ORTI"".""U_BDOSDstRl4"" AS ""BDOSDistrRule4"",
	                           ""ORTI"".""U_BDOSDstRl5"" AS ""BDOSDistrRule5"",
	                           ""ORTI"".""U_BDOSPrjCod"" AS ""BDOSPrjCode""
		
                        FROM ""RTI1""
                        LEFT JOIN ""ORTI"" ON ""RTI1"".""DocEntry"" = ""ORTI"".""DocEntry""
                        LEFT JOIN ""JDT1"" ON ""RTI1"".""DocEntry"" = ""JDT1"".""CreatedBy""
                        LEFT JOIN ""OITM"" ON ""RTI1"".""ItemCode"" = ""OITM"".""ItemCode""
                        LEFT JOIN ""ACS1"" ON ""OITM"".""AssetClass"" = ""ACS1"".""Code""
                        LEFT JOIN ""OADT"" ON ""ACS1"".""AcctDtn"" = ""OADT"".""Code""
                        LEFT JOIN ""OACT"" ON ""ORTI"".""U_BDOSExpAct"" = ""OACT"".""AcctCode""

                        WHERE ""ORTI"".""DocType"" = 'SC' AND ""RTI1"".""DocEntry"" =  '" + DocEntry + @"' AND ""JDT1"".""TransType"" = 1470000094 
		                        AND (""JDT1"".""Account"" = ""OADT"".""ReNBVeAct"" OR ""JDT1"".""Account"" = ""OADT"".""ReNBVrAct"" OR ""JDT1"".""Account"" = ""OADT"".""ReExpNAct"" OR ""JDT1"".""Account"" = ""OADT"".""ReRevNAct"")
                        GROUP BY ""ORTI"".""PrjSmarz"",
                                ""ORTI"".""DstRlSmarz"",
                                ""JDT1"".""Account"",
                        	    ""JDT1"".""ProfitCode"",
	   							""JDT1"".""OcrCode2"",
	   							""JDT1"".""OcrCode3"",
	   							""JDT1"".""OcrCode4"",
	   							""JDT1"".""OcrCode5"",
	   							""JDT1"".""Project"",
                                ""OACT"".""ActType"",
		                       	""ORTI"".""U_BDOSExpAct"",
	                            ""ORTI"".""U_BDOSDstRl1"",
	                            ""ORTI"".""U_BDOSDstRl2"",
	                            ""ORTI"".""U_BDOSDstRl3"",
	                            ""ORTI"".""U_BDOSDstRl4"",
	                            ""ORTI"".""U_BDOSDstRl5"",
	                            ""ORTI"".""U_BDOSPrjCod""";

                oRecordSet.DoQuery(query);

                int i = 0;

                while (!oRecordSet.EoF)
                {
                    jeLinesRow = jeLines.Rows.Add(i);
                    jeLinesRow["AccountCode"] = oRecordSet.Fields.Item("BDOSExpAccount").Value;
                    jeLinesRow["ShortName"] = oRecordSet.Fields.Item("BDOSExpAccount").Value;
                    jeLinesRow["ContraAccount"] = oRecordSet.Fields.Item("Account").Value;
                    jeLinesRow["Debit"] = oRecordSet.Fields.Item("Amount").Value;
                    jeLinesRow["Credit"] = 0;
                    if (oRecordSet.Fields.Item("BDOSExpAccount_ActType").Value == "E")
                    {
                        if (oRecordSet.Fields.Item("PrjSmarz").Value == "Y")
                        {
                            jeLinesRow["ProjectCode"] = oRecordSet.Fields.Item("BDOSPrjCode").Value;
                        }
                        if (oRecordSet.Fields.Item("DstrlSmarz").Value == "Y")
                        {
                            jeLinesRow["CostingCode"] = oRecordSet.Fields.Item("BDOSDistrRule1").Value;
                            jeLinesRow["CostingCode2"] = oRecordSet.Fields.Item("BDOSDistrRule2").Value;
                            jeLinesRow["CostingCode3"] = oRecordSet.Fields.Item("BDOSDistrRule3").Value;
                            jeLinesRow["CostingCode4"] = oRecordSet.Fields.Item("BDOSDistrRule4").Value;
                            jeLinesRow["CostingCode5"] = oRecordSet.Fields.Item("BDOSDistrRule5").Value;
                        }
                    }
                    i++;

                    jeLinesRow = jeLines.Rows.Add(i);
                    jeLinesRow["AccountCode"] = oRecordSet.Fields.Item("Account").Value;
                    jeLinesRow["ShortName"] = oRecordSet.Fields.Item("Account").Value;
                    jeLinesRow["ContraAccount"] = oRecordSet.Fields.Item("BDOSExpAccount").Value;
                    jeLinesRow["Debit"] = 0;
                    jeLinesRow["Credit"] = oRecordSet.Fields.Item("Amount").Value;
                    jeLinesRow["ProjectCode"] = oRecordSet.Fields.Item("PrjCode").Value;
                    jeLinesRow["CostingCode"] = oRecordSet.Fields.Item("DistrRule1").Value;
                    jeLinesRow["CostingCode2"] = oRecordSet.Fields.Item("DistrRule2").Value;
                    jeLinesRow["CostingCode3"] = oRecordSet.Fields.Item("DistrRule3").Value;
                    jeLinesRow["CostingCode4"] = oRecordSet.Fields.Item("DistrRule4").Value;
                    jeLinesRow["CostingCode5"] = oRecordSet.Fields.Item("DistrRule5").Value;
                    i++;

                    oRecordSet.MoveNext();
                }

                JournalEntry.JrnEntry( DocEntry, "1470000094", "Retirement: " + DocNum, DocDate, jeLines, out errorText);

                if (errorText != null)
                {
                    return;
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
        }

    }
}
