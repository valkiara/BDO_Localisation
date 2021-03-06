using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Data;
using System.Globalization;
using System.Windows.Forms;
using System.Xml;
using DataTable = System.Data.DataTable;


namespace BDO_Localisation_AddOn
{
    static partial class JournalEntry
    {
        public static void createUserFields(out string errorText)
        {
            errorText = null;
            Dictionary<string, object> fieldskeysMap;

            #region Employee ID

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSEmpID");
            fieldskeysMap.Add("TableName", "JDT1");
            fieldskeysMap.Add("Description", "Employee ID");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            #endregion

            #region AC Number

            fieldskeysMap = new Dictionary<string, object>(); //  A/C Number ივსება Good Receipt PO, AP Invoice, AP Credit memo, Landed Cost დოკუმენტებიდან
            fieldskeysMap.Add("Name", "BDOSACNum");
            fieldskeysMap.Add("TableName", "OJDT");
            fieldskeysMap.Add("Description", "A/C Number");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            #endregion

            #region Blanket Agreement Number

            fieldskeysMap = new Dictionary<string, object>(); //  Blanket Agreement Number
            fieldskeysMap.Add("Name", "BDOSAgrNo");
            fieldskeysMap.Add("TableName", "OJDT");
            fieldskeysMap.Add("Description", "Blanket Agreement Number");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            #endregion

            #region Use Blanket Agreement Rate Ranges

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "UseBlaAgRt");
            fieldskeysMap.Add("TableName", "OJDT");
            fieldskeysMap.Add("Description", "Use Blanket Agreement Rates");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 1);
            fieldskeysMap.Add("DefaultValue", "N");

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            #endregion
        }

        public static void JrnEntry(string jeReference, string jeReference2, string remark, DateTime jeDate, DataTable jeLines, out string errorText)
        {
            errorText = null;

            if (jeLines.Rows.Count == 0)
            {
                return;
            }

            SAPbobsCOM.JournalEntries oJounalEntry = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);

            oJounalEntry.Reference = jeReference;
            oJounalEntry.Reference2 = jeReference2;
            oJounalEntry.Memo = remark;
            oJounalEntry.DueDate = jeDate;
            oJounalEntry.ReferenceDate = jeDate;

            DataRow jeLine;
            SAPbobsCOM.JournalEntries_Lines oJLines = oJounalEntry.Lines;
            bool isFCCurr = false;

            for (int i = 0; i < jeLines.Rows.Count; i++)
            {
                jeLine = jeLines.Rows[i];
                if (jeLine["FCCurrency"].ToString() != "" && jeLine["FCCurrency"].ToString() != Program.LocalCurrency)
                {
                    isFCCurr = true;
                }
            }

            for (int i = 0; i < jeLines.Rows.Count; i++)
            {
                jeLine = jeLines.Rows[i];

                oJLines.AccountCode = jeLine["AccountCode"].ToString();
                oJLines.ShortName = jeLine["ShortName"].ToString();
                oJLines.ContraAccount = jeLine["ContraAccount"].ToString();
                oJLines.TaxGroup = jeLine["VatGroup"].ToString();

                if (jeLine["FCCurrency"].ToString() != "")
                {
                    oJLines.FCCredit = Convert.ToDouble(jeLine["FCCredit"]);
                    oJLines.FCDebit = Convert.ToDouble(jeLine["FCDebit"]);
                    oJLines.FCCurrency = jeLine["FCCurrency"].ToString();
                }
                else
                {
                    oJLines.FCCurrency = isFCCurr ? Program.LocalCurrency : "";
                    oJLines.FCCredit = 0;
                    oJLines.FCDebit = 0;
                }

                oJLines.Credit = Convert.ToDouble(jeLine["Credit"]);
                oJLines.Debit = Convert.ToDouble(jeLine["Debit"]);

                oJLines.CostingCode = jeLine["CostingCode"].ToString();
                oJLines.CostingCode2 = jeLine["CostingCode2"].ToString();
                oJLines.CostingCode3 = jeLine["CostingCode3"].ToString();
                oJLines.CostingCode4 = jeLine["CostingCode4"].ToString();
                oJLines.CostingCode5 = jeLine["CostingCode5"].ToString();
                oJLines.ProjectCode = jeLine["ProjectCode"].ToString();

                oJLines.UserFields.Fields.Item("U_BDOSEmpID").Value = jeLine["U_BDOSEmpID"].ToString().Trim();

                oJLines.Add();
            }

            int lRetCode = 0;
            lRetCode = oJounalEntry.Add();

            if (lRetCode != 0)
            {
                Program.oCompany.GetLastError(out lRetCode, out errorText);
            }
        }

        public static DataTable JournalEntryTable()
        {
            DataTable jeLines = new DataTable();
            jeLines.Columns.Add("AccountCode", typeof(string));
            jeLines.Columns.Add("ShortName", typeof(string));
            jeLines.Columns.Add("ContraAccount", typeof(string));
            jeLines.Columns.Add("Debit", typeof(double)).DefaultValue = 0;
            jeLines.Columns.Add("Credit", typeof(double)).DefaultValue = 0;
            jeLines.Columns.Add("FCDebit", typeof(double)).DefaultValue = 0;
            jeLines.Columns.Add("FCCredit", typeof(double)).DefaultValue = 0;
            jeLines.Columns.Add("FCCurrency", typeof(string));

            jeLines.Columns.Add("ProjectCode", typeof(string));
            jeLines.Columns.Add("CostingCode", typeof(string));
            jeLines.Columns.Add("CostingCode2", typeof(string));
            jeLines.Columns.Add("CostingCode3", typeof(string));
            jeLines.Columns.Add("CostingCode4", typeof(string));
            jeLines.Columns.Add("CostingCode5", typeof(string));

            jeLines.Columns.Add("VatGroup", typeof(string));
            jeLines.Columns.Add("U_BDOSEmpID", typeof(string));

            return jeLines;
        }

        public static void AddJournalEntryRow(DataTable accountCodes, DataTable jeLines, string entryType, string debAccount, string credAccount, decimal amount, decimal fcAmount, string currency, string distrRule1, string distrRule2, string distrRule3, string distrRule4, string distrRule5, string prjCode, string vatGroup, string bdoSEmpID)
        {
            string emptyAccountTxt = null;

            if (entryType == "Full")
            {
                if (string.IsNullOrEmpty(debAccount) && string.IsNullOrEmpty(credAccount))
                    emptyAccountTxt = $"{BDOSResources.getTranslate("Debit")}, {BDOSResources.getTranslate("Credit")}";
                else if (string.IsNullOrEmpty(debAccount))
                    emptyAccountTxt = $"{BDOSResources.getTranslate("Debit")}";
                else if (string.IsNullOrEmpty(credAccount))
                    emptyAccountTxt = $"{BDOSResources.getTranslate("Credit")}";
            }
            else if (entryType == "OnlyCredit")
            {
                if (string.IsNullOrEmpty(credAccount))
                    emptyAccountTxt = $"{BDOSResources.getTranslate("Credit")}";
            }
            else if (entryType == "OnlyDebit")
            {
                if (string.IsNullOrEmpty(debAccount))
                    emptyAccountTxt = $"{BDOSResources.getTranslate("Debit")}";
            }

            if (!string.IsNullOrEmpty(emptyAccountTxt))
                throw new Exception($"{BDOSResources.getTranslate("AccountIsNotCompleted")} - {emptyAccountTxt}");

            DataRow jeLinesRow = null;
            DataRowCollection jeLinesRows = jeLines.Rows;
            //დებეტი
            if (entryType != "OnlyCredit")
            {
                jeLinesRow = jeLinesRows.Add();
                jeLinesRow["AccountCode"] = debAccount;
                jeLinesRow["ShortName"] = debAccount;
                jeLinesRow["ContraAccount"] = credAccount;
                jeLinesRow["Debit"] = Convert.ToDouble(amount);
                jeLinesRow["FCDebit"] = Convert.ToDouble(fcAmount);
                jeLinesRow["FCCurrency"] = currency;
                jeLinesRow["VatGroup"] = vatGroup;
                jeLinesRow["ProjectCode"] = prjCode;
                jeLinesRow["Credit"] = 0;
                jeLinesRow["FCCredit"] = 0;

                DataRow[] oAccountCode = accountCodes.Select("AcctCode = '" + debAccount + "'");
                string AccountType = oAccountCode[0]["ActType"].ToString();
                string U_BDOSEmpAct = oAccountCode[0]["U_BDOSEmpAct"].ToString();

                if (AccountType != "N")
                {
                    jeLinesRow["CostingCode"] = distrRule1;
                    jeLinesRow["CostingCode2"] = distrRule2;
                    jeLinesRow["CostingCode3"] = distrRule3;
                    jeLinesRow["CostingCode4"] = distrRule4;
                    jeLinesRow["CostingCode5"] = distrRule5;
                }
                if (U_BDOSEmpAct == "Y")
                {
                    jeLinesRow["U_BDOSEmpID"] = bdoSEmpID;
                }
            }

            //კტ
            if (entryType != "OnlyDebit")
            {
                jeLinesRow = jeLinesRows.Add();
                jeLinesRow["AccountCode"] = credAccount;
                jeLinesRow["ShortName"] = credAccount;
                jeLinesRow["ContraAccount"] = debAccount;
                jeLinesRow["Credit"] = Convert.ToDouble(amount);
                jeLinesRow["FCCredit"] = Convert.ToDouble(fcAmount);
                jeLinesRow["FCCurrency"] = currency;
                jeLinesRow["VatGroup"] = vatGroup;
                jeLinesRow["ProjectCode"] = prjCode;
                jeLinesRow["Debit"] = 0;
                jeLinesRow["FCDebit"] = 0;

                DataRow[] oAccountCode = accountCodes.Select("AcctCode = '" + credAccount + "'");
                string AccountType = oAccountCode[0]["ActType"].ToString();
                string U_BDOSEmpAct = oAccountCode[0]["U_BDOSEmpAct"].ToString();
                if (AccountType != "N")
                {
                    jeLinesRow["CostingCode"] = distrRule1;
                    jeLinesRow["CostingCode2"] = distrRule2;
                    jeLinesRow["CostingCode3"] = distrRule3;
                    jeLinesRow["CostingCode4"] = distrRule4;
                    jeLinesRow["CostingCode5"] = distrRule5;
                }
                if (U_BDOSEmpAct == "Y")
                {
                    jeLinesRow["U_BDOSEmpID"] = bdoSEmpID;
                }
            }
        }

        public static void createFormItems(SAPbouiCOM.Form oForm)
        {
            string errorText;

            SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("76").Specific;
            SAPbouiCOM.Column oColumn = oMatrix.Columns.Item("U_BDOSEmpID");
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("EmployeeNo");

            SAPbouiCOM.Item oItem = oForm.Items.Item("9");

            int height = oItem.Height;
            int top = oItem.Top + oItem.Height + 5;
            int left = oItem.Left;
            int width = oItem.Width;

            Dictionary<string, object> formItems = new Dictionary<string, object>();
            string itemName = "BDOSJrnEnS";

            try
            {
                oForm.Items.Item(itemName);
            }
            catch
            {
                formItems.Add("Size", 20);
                formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                formItems.Add("Left", left);
                formItems.Add("Width", width);
                formItems.Add("Top", top);
                formItems.Add("Height", height);
                formItems.Add("UID", itemName);
                formItems.Add("Caption", BDOSResources.getTranslate("AdditionalEntry"));

                FormsB1.createFormItem(oForm, formItems, out errorText);
                if (errorText != null)
                {
                    throw new Exception(errorText);
                }
            }

            formItems = new Dictionary<string, object>();
            itemName = "BDOSJrnEnt";
            try
            {
                oForm.Items.Item(itemName);
            }
            catch
            {
                formItems.Add("Size", 20);
                formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                formItems.Add("Left", left);
                formItems.Add("Width", width);
                formItems.Add("Top", top + height + 2);
                formItems.Add("Height", height);
                formItems.Add("UID", itemName);
                formItems.Add("Enabled", false);
                formItems.Add("AffectsFormMode", false);

                FormsB1.createFormItem(oForm, formItems, out errorText);
                if (errorText != null)
                {
                    throw new Exception(errorText);
                }
            }

            formItems = new Dictionary<string, object>();
            itemName = "BDOSJEntLB";
            try
            {
                oForm.Items.Item(itemName);
            }
            catch
            {
                formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                formItems.Add("Left", left - 20);
                formItems.Add("Top", top + height + 2);
                formItems.Add("UID", itemName);
                formItems.Add("LinkTo", "BDOSJrnEnt");
                formItems.Add("LinkedObjectType", "30");

                FormsB1.createFormItem(oForm, formItems, out errorText);
                if (errorText != null)
                {
                    throw new Exception(errorText);
                }
            }

            oItem = oForm.Items.Item("7");
            left = oItem.Left;
            width = oItem.Width;

            formItems = new Dictionary<string, object>();
            itemName = "BDOSACNumS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left);
            formItems.Add("Width", width);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("ACNumber"));
            formItems.Add("LinkTo", "BDOSACNumE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }

            formItems = new Dictionary<string, object>();
            itemName = "BDOSACNumE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OJDT");
            formItems.Add("Alias", "U_BDOSACNum");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left);
            formItems.Add("Width", width);
            formItems.Add("Top", top + height + 2);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }

            #region Blanket Agreement Number string and element

            oItem = oForm.Items.Item("1980002040");
            top = oItem.Top;
            width = oItem.Width;

            formItems = new Dictionary<string, object>();
            itemName = "BDOSAgrNoS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left + 20);
            formItems.Add("Width", width);
            formItems.Add("Top", top - height - 1);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("BlanketAgreement"));
            formItems.Add("LinkTo", "BDOSAgrNoE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }

            var objectType = "1250000025"; //Blanket Agreement

            FormsB1.addChooseFromList(oForm, false, objectType, "BlanketAgreement_Cfl");

            formItems = new Dictionary<string, object>();
            itemName = "BDOSAgrNoE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OJDT");
            formItems.Add("Alias", "U_BDOSAgrNo");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left + 20);
            formItems.Add("Width", width);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("ChooseFromListUID", "BlanketAgreement_Cfl");
            formItems.Add("ChooseFromListAlias", "AbsID");
            formItems.Add("SetAutoManaged", true);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }

            formItems = new Dictionary<string, object>();
            itemName = "BDOSAgrNoL"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            formItems.Add("Left", left + 1);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("LinkTo", "BDOSAgrNoE");
            formItems.Add("LinkedObjectType", objectType);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            #endregion

            #region Use Blanket Agreement Rate Ranges

            formItems = new Dictionary<string, object>();
            itemName = "UsBlaAgRtS"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OJDT");
            formItems.Add("Alias", "U_UseBlaAgRt");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
            formItems.Add("Length", 1);
            formItems.Add("Left", left + width + 20);
            formItems.Add("Width", 100);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("UseBlAgrRt"));
            formItems.Add("ValOff", "N");
            formItems.Add("ValOn", "Y");
            formItems.Add("DisplayDesc", true);
            formItems.Add("SetAutoManaged", true);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            GC.Collect();


            #endregion

            formItems = new Dictionary<string, object>();
            itemName = "BDOSAddEnt";
            try
            {
                oForm.Items.Item(itemName);
            }
            catch
            {
                formItems.Add("isDataSource", true);
                formItems.Add("Length", 1);
                formItems.Add("DataSource", "UserDataSources");
                formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                formItems.Add("TableName", "");
                formItems.Add("Alias", itemName);
                formItems.Add("Bound", true);
                formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
                formItems.Add("Width", oForm.Items.Item("37").Width);
                formItems.Add("Left", oForm.Items.Item("37").Left - oForm.Items.Item("37").Width - 20);
                formItems.Add("Top", oForm.Items.Item("37").Top);
                formItems.Add("Caption", BDOSResources.getTranslate("DisplayAE"));
                formItems.Add("AffectsFormMode", false);
                formItems.Add("UID", itemName);
                formItems.Add("Enabled", true);
                formItems.Add("ValueOn", "Y");
                formItems.Add("ValueOff", "N");

                FormsB1.createFormItem(oForm, formItems, out errorText);
                if (errorText != null)
                {
                    throw new Exception(errorText);
                }
            }
            //ShowAdditionalEntries(oForm);

            SAPbouiCOM.DBDataSource JDT1 = oForm.DataSources.DBDataSources.Item("JDT1");
            SAPbouiCOM.DBDataSource OJDT = oForm.DataSources.DBDataSources.Item("OJDT");
            SAPbouiCOM.Item MatrixItem = oForm.Items.Item("76");
            oMatrix = MatrixItem.Specific;

            SAPbouiCOM.Columns oColumns = null;

            SAPbouiCOM.DataTable JDT1_BDO = oForm.DataSources.DataTables.Add("JDT1_BDO");

            for (int i = 0; i < JDT1.Fields.Count; i++)
            {
                JDT1_BDO.Columns.Add(JDT1.Fields.Item(i).Name, JDT1.Fields.Item(i).Type, JDT1.Fields.Item(i).Size);
            }

            JDT1_BDO.Columns.Add("AcctName", SAPbouiCOM.BoFieldsType.ft_Text, 100);
            JDT1_BDO.Columns.Add("prFrmItm", SAPbouiCOM.BoFieldsType.ft_Text, 100);

            formItems = new Dictionary<string, object>();
            itemName = "JDT1BDOS";
            formItems = new Dictionary<string, object>();
            formItems.Add("isDataSource", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
            formItems.Add("Left", MatrixItem.Left);
            formItems.Add("Width", MatrixItem.Width);
            formItems.Add("Top", MatrixItem.Top);
            formItems.Add("Height", MatrixItem.Height);
            formItems.Add("UID", itemName);
            formItems.Add("FromPane", MatrixItem.FromPane);
            formItems.Add("ToPane", MatrixItem.ToPane);
            formItems.Add("Visible", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }

            SAPbouiCOM.Matrix MatrixJDT1BDOS = oForm.Items.Item("JDT1BDOS").Specific;
            oColumns = MatrixJDT1BDOS.Columns;

            //სტანდარტული ველების ხელოვნურად გამოჩენა (ანგარიში, თანხები, ვალუტა...)
            oColumn = oColumns.Add("Account", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oColumn.TitleObject.Caption = oMatrix.Columns.Item("1").TitleObject.Caption;
            oColumn.Width = 10;
            oColumn.Editable = false;
            SAPbouiCOM.LinkedButton oLink = oColumn.ExtendedObject;
            oLink.LinkedObjectType = "1";
            oColumn.DataBind.Bind("JDT1_BDO", "ShortName");

            oColumn = oColumns.Add("2", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = oMatrix.Columns.Item("2").TitleObject.Caption;
            oColumn.Width = 10;
            oColumn.Editable = false;
            oColumn.DataBind.Bind("JDT1_BDO", oMatrix.Columns.Item("2").DataBind.Alias);

            oColumn = oColumns.Add("ContrlAct", oMatrix.Columns.Item("37").Type);
            oColumn.TitleObject.Caption = oMatrix.Columns.Item("37").TitleObject.Caption;
            oColumn.Width = 10;
            oColumn.Editable = false;
            oLink = oColumn.ExtendedObject;
            oLink.LinkedObjectType = "1";
            oColumn.DataBind.Bind("JDT1_BDO", "Account");

            //Deb Cr FCAmounts
            oColumn = oColumns.Add("FCCurrency", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("FCCurrency");
            oColumn.Width = 10;
            oColumn.Editable = false;
            oColumn.DataBind.Bind("JDT1_BDO", "FCCurrency");

            oColumn = oColumns.Add("FCDebit", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("FCDebit");
            oColumn.Width = 10;
            oColumn.Editable = false;
            oColumn.DataBind.Bind("JDT1_BDO", "FCDebit");
            oColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;

            oColumn = oColumns.Add("FCCredit", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("FCCredit");
            oColumn.Width = 10;
            oColumn.Editable = false;
            oColumn.DataBind.Bind("JDT1_BDO", "FCCredit");
            oColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;

            //Deb Cr Amounts
            oColumn = oColumns.Add("Debit", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Debit");
            oColumn.Width = 10;
            oColumn.Editable = false;
            oColumn.DataBind.Bind("JDT1_BDO", "Debit");
            oColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;

            oColumn = oColumns.Add("Credit", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Credit");
            oColumn.Width = 10;
            oColumn.Editable = false;
            oColumn.DataBind.Bind("JDT1_BDO", "Credit");
            oColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;

            //Deb Cr Amounts
            oColumn = oColumns.Add("SYSDeb", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("SCDebit");
            oColumn.Width = 10;
            oColumn.Editable = false;
            oColumn.DataBind.Bind("JDT1_BDO", "SYSDeb");
            oColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;

            oColumn = oColumns.Add("SYSCred", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("SCCredit");
            oColumn.Width = 10;
            oColumn.Editable = false;
            oColumn.DataBind.Bind("JDT1_BDO", "SYSCred");
            oColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;

            oColumn = oColumns.Add("9", oMatrix.Columns.Item("9").Type);
            oColumn.TitleObject.Caption = oMatrix.Columns.Item("9").TitleObject.Caption;
            oColumn.Width = 10;
            oColumn.Editable = false;
            oColumn.DataBind.Bind("JDT1_BDO", oMatrix.Columns.Item("9").DataBind.Alias);

            oColumn = oColumns.Add("17", oMatrix.Columns.Item("17").Type);
            oColumn.TitleObject.Caption = oMatrix.Columns.Item("17").TitleObject.Caption;
            oColumn.Width = 10;
            oColumn.Editable = false;
            oColumn.DataBind.Bind("JDT1_BDO", oMatrix.Columns.Item("17").DataBind.Alias);

            //oColumn = oColumns.Add("480002020", oMatrix.Columns.Item("480002020").Type);
            //oColumn.TitleObject.Caption = oMatrix.Columns.Item("480002020").TitleObject.Caption;
            //oColumn.Width = 10;
            //oColumn.Editable = false;
            //oColumn.DataBind.Bind("JDT1_BDO", oMatrix.Columns.Item("480002020").DataBind.Alias);

            oColumn = oColumns.Add("2006", oMatrix.Columns.Item("2006").Type);
            oColumn.TitleObject.Caption = oMatrix.Columns.Item("2006").TitleObject.Caption;
            oColumn.Width = 10;
            oColumn.Editable = false;
            oColumn.DataBind.Bind("JDT1_BDO", oMatrix.Columns.Item("2006").DataBind.Alias);

            oColumn = oColumns.Add("2001", oMatrix.Columns.Item("2001").Type);
            oColumn.TitleObject.Caption = oMatrix.Columns.Item("2001").TitleObject.Caption;
            oColumn.Width = 10;
            oColumn.Editable = false;
            oColumn.DataBind.Bind("JDT1_BDO", oMatrix.Columns.Item("2001").DataBind.Alias);

            oColumn = oColumns.Add("2003", oMatrix.Columns.Item("2003").Type);
            oColumn.TitleObject.Caption = oMatrix.Columns.Item("2003").TitleObject.Caption;
            oColumn.Width = 10;
            oColumn.Editable = false;
            oColumn.DataBind.Bind("JDT1_BDO", oMatrix.Columns.Item("2003").DataBind.Alias);

            oColumn = oColumns.Add("2004", oMatrix.Columns.Item("2004").Type);
            oColumn.TitleObject.Caption = oMatrix.Columns.Item("2004").TitleObject.Caption;
            oColumn.Width = 10;
            oColumn.Editable = false;
            oColumn.DataBind.Bind("JDT1_BDO", oMatrix.Columns.Item("2004").DataBind.Alias);

            oColumn = oColumns.Add("2005", oMatrix.Columns.Item("2005").Type);
            oColumn.TitleObject.Caption = oMatrix.Columns.Item("2005").TitleObject.Caption;
            oColumn.Width = 10;
            oColumn.Editable = false;
            oColumn.DataBind.Bind("JDT1_BDO", oMatrix.Columns.Item("2005").DataBind.Alias);

            oColumn = oColumns.Add("16", oMatrix.Columns.Item("16").Type);
            oColumn.TitleObject.Caption = oMatrix.Columns.Item("16").TitleObject.Caption;
            oColumn.Width = 10;
            oColumn.Editable = false;
            oColumn.DataBind.Bind("JDT1_BDO", oMatrix.Columns.Item("16").DataBind.Alias);

            oColumn = oColumns.Add("prFrmItm", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("prFormItem");
            oColumn.Width = 10;
            oColumn.Editable = false;
            oColumn.DataBind.Bind("JDT1_BDO", "prFrmItm");

            oColumn = oColumns.Add("BDOSEmpID", oMatrix.Columns.Item("U_BDOSEmpID").Type);
            oColumn.TitleObject.Caption = oMatrix.Columns.Item("U_BDOSEmpID").TitleObject.Caption;
            oColumn.Width = 10;
            oColumn.Editable = false;
            oColumn.DataBind.Bind("JDT1_BDO", oMatrix.Columns.Item("U_BDOSEmpID").DataBind.Alias);

            //for new HR & Payroll Add-On
            try
            {
                oColumn = oColumns.Add("SlryRule", oMatrix.Columns.Item("U_slryRuleCode").Type);
                oColumn.TitleObject.Caption = oMatrix.Columns.Item("U_slryRuleCode").TitleObject.Caption;
                oColumn.Width = 10;
                oColumn.Editable = false;
                oColumn.DataBind.Bind("JDT1_BDO", oMatrix.Columns.Item("U_slryRuleCode").DataBind.Alias);

                oColumn = oColumns.Add("EmpCode", oMatrix.Columns.Item("U_empCode").Type);
                oColumn.TitleObject.Caption = oMatrix.Columns.Item("U_empCode").TitleObject.Caption;
                oColumn.Width = 10;
                oColumn.Editable = false;
                oColumn.DataBind.Bind("JDT1_BDO", oMatrix.Columns.Item("U_empCode").DataBind.Alias);

                oColumn = oColumns.Add("AccrMnth", oMatrix.Columns.Item("U_accrMnth").Type);
                oColumn.TitleObject.Caption = oMatrix.Columns.Item("U_accrMnth").TitleObject.Caption;
                oColumn.Width = 10;
                oColumn.Editable = false;
                oColumn.DataBind.Bind("JDT1_BDO", oMatrix.Columns.Item("U_accrMnth").DataBind.Alias);

                oColumn = oColumns.Add("WTaxCode", oMatrix.Columns.Item("U_wTaxCode").Type);
                oColumn.TitleObject.Caption = oMatrix.Columns.Item("U_wTaxCode").TitleObject.Caption;
                oColumn.Width = 10;
                oColumn.Editable = false;
                oColumn.DataBind.Bind("JDT1_BDO", oMatrix.Columns.Item("U_wTaxCode").DataBind.Alias);
            }
            catch { }
        }

        private static void ChooseFromList(SAPbouiCOM.Form oForm, SAPbouiCOM.ItemEvent pVal, SAPbouiCOM.ChooseFromListEvent oCflEventO, out bool bubbleEvent)
        {
            bubbleEvent = true;

            var oCflId = oCflEventO.ChooseFromListUID;
            var oCfl = oForm.ChooseFromLists.Item(oCflId);

            if (!pVal.BeforeAction)
            {
                try
                {
                    oForm.Freeze(true);

                    var oDataTable = oCflEventO.SelectedObjects;

                    if (oDataTable == null) return;

                    if (oCflId == "BlanketAgreement_Cfl")
                    {
                        string blanketAgreementNumber = Convert.ToString(oDataTable.GetValue("AbsID", 0));
                        LanguageUtils.IgnoreErrors<string>(() => oForm.Items.Item("BDOSAgrNoE").Specific.Value = blanketAgreementNumber); //ერორს აგდებს, რატო კაცმა არ იცის
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

            else
            {
                if (oCflId == "BlanketAgreement_Cfl")
                {
                    if (IsMultiBp(oForm, out var bpCode))
                    {
                        Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("CantChooseBlaAgrForMultiBp"), SAPbouiCOM.BoMessageTime.bmt_Short);
                        bubbleEvent = false;
                        return;
                    }

                    var oCons = new SAPbouiCOM.Conditions();

                    if (bpCode.Length != 0)
                    {
                        var oCon = oCons.Add();
                        oCon.Alias = "BpCode";
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = bpCode;
                    }
                    else
                    {
                        var oCon = oCons.Add();
                        oCon.Alias = "AbsID";
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = "";
                    }
                    oCfl.SetConditions(oCons);
                }
            }
        }

        private static void ShowAdditionalEntries(SAPbouiCOM.Form oForm)
        {
            oForm.Freeze(true);
            try
            {
                SAPbouiCOM.DBDataSource JDT1 = oForm.DataSources.DBDataSources.Item("JDT1");
                SAPbouiCOM.DBDataSource OJDT = oForm.DataSources.DBDataSources.Item("OJDT");
                SAPbouiCOM.Item MatrixItem = oForm.Items.Item("76");
                SAPbouiCOM.Matrix oMatrix = MatrixItem.Specific;

                //ჩვენი ცხრილის შევსება
                //ძირითადი გატარება
                SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                SAPbouiCOM.DataTable JDT1_BDO = oForm.DataSources.DataTables.Item("JDT1_BDO");
                JDT1_BDO.Rows.Clear();

                for (int i = 0; i < JDT1.Size; i++)
                {
                    if (JDT1.GetValue("Account", i).Trim() == "")
                    {
                        continue;
                    }

                    JDT1_BDO.Rows.Add();

                    for (int j = 0; j < JDT1.Fields.Count; j++)
                    {
                        if (JDT1.GetValue(j, i) != "")
                        {
                            JDT1_BDO.SetValue(j, i, JDT1.GetValue(j, i));
                        }

                    }
                    JDT1_BDO.SetValue("AcctName", i, oMatrix.Columns.Item("2").Cells.Item(i + 1).Specific.Value);
                    JDT1_BDO.SetValue("prFrmItm", i, getCashFlow(JDT1_BDO.GetValue("TransId", i).ToString(), JDT1_BDO.GetValue("Line_ID", i).ToString()));
                }

                int count = JDT1_BDO.Rows.Count;
                SAPbobsCOM.SBObob vObj = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);

                //დამატებითი გატარება
                if (Program.JrnLinesGlobal.Rows.Count > 0)
                {
                    oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    string query = "SELECT \"MainCurncy\" , \"SysCurrncy\" FROM \"OADM\"";
                    oRecordSet.DoQuery(query);
                    double SYSRate = 0;
                    if (!oRecordSet.EoF)
                    {
                        DateTime DocDate = DateTime.ParseExact(OJDT.GetValue("RefDate", 0), "yyyyMMdd", CultureInfo.InvariantCulture);
                        string SysCurrncy = oRecordSet.Fields.Item("SysCurrncy").Value.ToString();
                        string MainCurncy = oRecordSet.Fields.Item("MainCurncy").Value.ToString();

                        if (SysCurrncy == MainCurncy)
                        {
                            SYSRate = 1;
                        }
                        else
                        {
                            try
                            {
                                SYSRate = vObj.GetCurrencyRate(SysCurrncy, DocDate).Fields.Item("CurrencyRate").Value;
                            }
                            catch
                            {
                                SYSRate = 1;

                            }
                        }
                    }

                    for (int i = 0; i < Program.JrnLinesGlobal.Rows.Count; i++)
                    {
                        int GlCount = count + i;
                        JDT1_BDO.Rows.Add();

                        DataRow dtRow = Program.JrnLinesGlobal.Rows[i];

                        JDT1_BDO.SetValue("ShortName", GlCount, dtRow["AccountCode"]);
                        JDT1_BDO.SetValue("Account", GlCount, dtRow["AccountCode"]);
                        JDT1_BDO.SetValue("AcctName", GlCount, dtRow["ShortName"]);
                        JDT1_BDO.SetValue("ContraAct", GlCount, dtRow["ContraAccount"]);
                        JDT1_BDO.SetValue("Credit", GlCount, dtRow["Credit"]);
                        JDT1_BDO.SetValue("Debit", GlCount, dtRow["Debit"]);
                        JDT1_BDO.SetValue("SYSCred", GlCount, SYSRate > 0 ? Convert.ToDouble(dtRow["Credit"]) / SYSRate : 0);
                        JDT1_BDO.SetValue("SYSDeb", GlCount, SYSRate > 0 ? Convert.ToDouble(dtRow["Debit"]) / SYSRate : 0);
                        JDT1_BDO.SetValue("FCCredit", GlCount, dtRow["FCCredit"]);
                        JDT1_BDO.SetValue("FCDebit", GlCount, dtRow["FCDebit"]);
                        JDT1_BDO.SetValue("FCCurrency", GlCount, dtRow["FCCurrency"].ToString());
                        JDT1_BDO.SetValue("AcctName", GlCount, getAcctName(JDT1_BDO.GetValue("Account", GlCount)));
                        JDT1_BDO.SetValue("ProfitCode", GlCount, dtRow["CostingCode"].ToString());
                        JDT1_BDO.SetValue("OcrCode2", GlCount, dtRow["CostingCode2"].ToString());
                        JDT1_BDO.SetValue("OcrCode3", GlCount, dtRow["CostingCode3"].ToString());
                        JDT1_BDO.SetValue("OcrCode4", GlCount, dtRow["CostingCode4"].ToString());
                        JDT1_BDO.SetValue("OcrCode5", GlCount, dtRow["CostingCode5"].ToString());
                        JDT1_BDO.SetValue("Project", GlCount, dtRow["ProjectCode"].ToString());
                        JDT1_BDO.SetValue("VatGroup", GlCount, dtRow["VatGroup"].ToString());
                        JDT1_BDO.SetValue("U_BDOSEmpID", GlCount, dtRow["U_BDOSEmpID"].ToString());
                    }

                    Program.JrnLinesGlobal = new DataTable();
                }
                else
                {
                    string transId = OJDT.GetValue("TransId", 0);

                    if (transId != "")
                    {
                        string TransType = OJDT.GetValue("TransType", 0).Trim();
                        string CreatedBy = OJDT.GetValue("CreatedBy", 0).Trim();

                        oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                        StringBuilder query = new StringBuilder();
                        query.Append("SELECT T1.* \n");
                        query.Append("FROM   \"JDT1\" T1 \n");
                        query.Append("       LEFT JOIN \"OJDT\" T0 \n");
                        query.Append("              ON T0.\"TransId\" = T1.\"TransId\" \n");
                        query.Append($"WHERE  ( T0.\"Ref1\" = '{CreatedBy}' \n");
                        query.Append($"         AND T0.\"Ref2\" = '{TransType}' ) \n");
                        query.Append($"        OR ( T0.\"Ref3\" = '{CreatedBy}' \n");
                        query.Append($"             AND T0.\"Ref2\" = '{TransType}' \n");
                        query.Append("             AND T0.\"U_byAddOn\" = 'Y' )");

                        oRecordSet.DoQuery(query.ToString());

                        int GlCount = count;

                        while (!oRecordSet.EoF)
                        {
                            JDT1_BDO.Rows.Add();

                            JDT1_BDO.SetValue("ShortName", GlCount, oRecordSet.Fields.Item("Account").Value);
                            JDT1_BDO.SetValue("Account", GlCount, oRecordSet.Fields.Item("Account").Value);
                            JDT1_BDO.SetValue("ContraAct", GlCount, oRecordSet.Fields.Item("ContraAct").Value);
                            JDT1_BDO.SetValue("Credit", GlCount, oRecordSet.Fields.Item("Credit").Value);
                            JDT1_BDO.SetValue("Debit", GlCount, oRecordSet.Fields.Item("Debit").Value);
                            JDT1_BDO.SetValue("SYSCred", GlCount, oRecordSet.Fields.Item("SYSCred").Value);
                            JDT1_BDO.SetValue("SYSDeb", GlCount, oRecordSet.Fields.Item("SYSDeb").Value);
                            JDT1_BDO.SetValue("FCCredit", GlCount, oRecordSet.Fields.Item("FCCredit").Value);
                            JDT1_BDO.SetValue("FCDebit", GlCount, oRecordSet.Fields.Item("FCDebit").Value);
                            JDT1_BDO.SetValue("FCCurrency", GlCount, oRecordSet.Fields.Item("FCCurrency").Value);
                            JDT1_BDO.SetValue("AcctName", GlCount, getAcctName(JDT1_BDO.GetValue("Account", GlCount)));
                            JDT1_BDO.SetValue("ProfitCode", GlCount, oRecordSet.Fields.Item("ProfitCode").Value);
                            JDT1_BDO.SetValue("OcrCode2", GlCount, oRecordSet.Fields.Item("OcrCode2").Value);
                            JDT1_BDO.SetValue("OcrCode3", GlCount, oRecordSet.Fields.Item("OcrCode3").Value);
                            JDT1_BDO.SetValue("OcrCode4", GlCount, oRecordSet.Fields.Item("OcrCode4").Value);
                            JDT1_BDO.SetValue("OcrCode5", GlCount, oRecordSet.Fields.Item("OcrCode5").Value);
                            JDT1_BDO.SetValue("Project", GlCount, oRecordSet.Fields.Item("Project").Value);
                            JDT1_BDO.SetValue("prFrmItm", GlCount, getCashFlow(JDT1_BDO.GetValue("TransId", GlCount).ToString(), JDT1_BDO.GetValue("Line_ID", GlCount).ToString()));
                            JDT1_BDO.SetValue("U_BDOSEmpID", GlCount, oRecordSet.Fields.Item("U_BDOSEmpID").Value);

                            //for new HR & Payroll Add-On
                            try
                            {
                                JDT1_BDO.SetValue("U_slryRuleCode", GlCount, oRecordSet.Fields.Item("U_slryRuleCode").Value);
                                JDT1_BDO.SetValue("U_empCode", GlCount, oRecordSet.Fields.Item("U_empCode").Value);
                                JDT1_BDO.SetValue("U_accrMnth", GlCount, oRecordSet.Fields.Item("U_accrMnth").Value);
                                JDT1_BDO.SetValue("U_wTaxCode", GlCount, oRecordSet.Fields.Item("U_wTaxCode").Value);
                            }
                            catch { }

                            GlCount++;
                            oRecordSet.MoveNext();
                        }
                    }
                }
                SAPbouiCOM.Matrix MatrixJDT1BDOS = oForm.Items.Item("JDT1BDOS").Specific;
                MatrixJDT1BDOS.Clear();
                MatrixJDT1BDOS.LoadFromDataSource();
                MatrixJDT1BDOS.AutoResizeColumns();
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

        public static string getAcctName(string Account)
        {

            SAPbobsCOM.Recordset oRecordSetOACT = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string queryOACT = @"SELECT ""AcctName"" FROM ""OACT""
                                             WHERE ""AcctCode"" = '" + Account + "'";

            oRecordSetOACT.DoQuery(queryOACT);

            if (!oRecordSetOACT.EoF)
            {
                return oRecordSetOACT.Fields.Item("AcctName").Value;
            }
            return "";

        }

        public static string getCashFlow(string TransId, string Line_ID)
        {
            SAPbobsCOM.Recordset oRecordSetOACT = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string queryOACT = @"SELECT
	                            ""CFWName"" 
                                FROM ""OCFT"" 
                                INNER JOIN ""OCFW"" ON ""OCFT"".""CFWId"" =""OCFW"".""CFWId"" 
                                WHERE ""OCFT"".""JDTId"" = " + TransId + @" AND ""OCFT"".""JDTLineId"" = " + Line_ID;

            oRecordSetOACT.DoQuery(queryOACT);

            if (!oRecordSetOACT.EoF)
            {
                return oRecordSetOACT.Fields.Item("CFWName").Value;
            }
            return "";
        }

        public static void formDataLoad(SAPbouiCOM.Form oForm)
        {
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                string TransType = oForm.DataSources.DBDataSources.Item("OJDT").GetValue("TransType", 0).Trim();
                string CreatedBy = oForm.DataSources.DBDataSources.Item("OJDT").GetValue("CreatedBy", 0).Trim();
                string StornoToTr = oForm.DataSources.DBDataSources.Item("OJDT").GetValue("StornoToTr", 0).Trim();

                string query;
                if (string.IsNullOrEmpty(StornoToTr))
                {
                    query = @"SELECT 
                            ""TransId""  
                            FROM ""OJDT"" 
                            WHERE ""StornoToTr"" IS NULL   
                            AND ""Ref1"" = '" + CreatedBy + @"'  
                            AND ""Ref2"" = '" + TransType + "' ";
                }
                else
                {
                    query = @"SELECT 
                            ""TransId""   
                            FROM ""OJDT"" 
                            WHERE ""StornoToTr"" IS NOT NULL   
                            AND ""Ref1"" = '" + CreatedBy + @"'
                            AND ""Ref2"" = '" + TransType + "' ";
                }

                oRecordSet.DoQuery(query);

                if (!oRecordSet.EoF)
                    oForm.Items.Item("BDOSJrnEnt").Specific.Value = oRecordSet.Fields.Item("TransId").Value;
                else
                    oForm.Items.Item("BDOSJrnEnt").Specific.Value = "";
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

        public static void uiApp_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;

            SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(BusinessObjectInfo.FormTypeEx, Program.currentFormCount);

            if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD && !BusinessObjectInfo.BeforeAction)
            {
                if (oForm.TypeEx == "392")
                {
                    formDataLoad(oForm);
                    oForm.Items.Item("BDOSAddEnt").Enabled = true;
                }
            }

            if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD &
                BusinessObjectInfo.BeforeAction == false & BusinessObjectInfo.ActionSuccess)
            {
                UpdateBlanketAgreementNumber(BusinessObjectInfo.ObjectKey, oForm);

                if (Program.canceledDocEntry == 0) return;
                try
                {
                    int transId =
                        Convert.ToInt32(oForm.Items.Item("BDOSJrnEnt").Specific.Value);

                    SAPbobsCOM.JournalEntries oJournalDoc =
                        Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
                    oJournalDoc.GetByKey(transId);

                    int response = oJournalDoc.Cancel();

                    if (response != 0)
                    {
                        Program.oCompany.GetLastError(out response, out string errorText);
                    }
                    else
                    {
                        Program.canceledDocEntry = 0;
                    }
                }
                catch (Exception ex)
                {
                    Program.uiApp.StatusBar.SetSystemMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short);
                }
            }
        }

        public static void cancellation(SAPbouiCOM.Form oForm, int Ref1, string Ref2, out string errorText)
        {
            errorText = null;
            int transId;

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                string queryHeader = "SELECT " +
                                "\"TransId\" " +
                                "FROM \"OJDT\"  " +
                                "WHERE \"Ref1\" = '" + Ref1.ToString() + "' " +
                                "AND \"Ref2\" = '" + Ref2 + "' ";

                oRecordSet.DoQuery(queryHeader);
                if (!oRecordSet.EoF)
                {
                    transId = oRecordSet.Fields.Item("TransId").Value;

                    SAPbobsCOM.JournalEntries oJournalDoc = (SAPbobsCOM.JournalEntries)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
                    oJournalDoc.GetByKey(transId);

                    int response = oJournalDoc.Cancel();

                    if (response != 0)
                    {
                        int errCode;
                        string errMsg;

                        Program.oCompany.GetLastError(out errCode, out errMsg);
                        errorText = BDOSResources.getTranslate("ErrorDescription") + ": " + errMsg + "! " + BDOSResources.getTranslate("Code") + ": " + errCode + "! ";
                    }
                }
            }
            catch (Exception ex)
            {
                errorText = BDOSResources.getTranslate("ErrorDescription") + ": " + ex.Message + "! ";
            }
            finally
            {
                Marshal.ReleaseComObject(oRecordSet);
            }
        }

        public static void setVisibleFormItems(SAPbouiCOM.Form oForm)
        {
            oForm.Freeze(true);
            try
            {
                SAPbouiCOM.CheckBox oCheckBoxFC = (SAPbouiCOM.CheckBox)oForm.Items.Item("37").Specific;
                SAPbouiCOM.CheckBox oCheckBoxLC = (SAPbouiCOM.CheckBox)oForm.Items.Item("36").Specific;

                bool visibleFC = oCheckBoxFC.Checked;
                bool visibleLC = oCheckBoxLC.Checked;

                bool additionalEntry = oForm.Items.Item("BDOSAddEnt").Specific.Checked;
                oForm.Items.Item("76").Visible = !additionalEntry;
                oForm.Items.Item("JDT1BDOS").Visible = additionalEntry;

                if (additionalEntry)
                {
                    SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("JDT1BDOS").Specific;
                    oMatrix.Columns.Item("FCDebit").Visible = visibleFC;
                    oMatrix.Columns.Item("FCCredit").Visible = visibleFC;
                    oMatrix.Columns.Item("FCCurrency").Visible = visibleFC;
                    oMatrix.Columns.Item("SYSCred").Visible = visibleLC;
                    oMatrix.Columns.Item("SYSDeb").Visible = visibleLC;

                    oMatrix.AutoResizeColumns();
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

        public static void uiApp_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD && pVal.BeforeAction)
                {
                    createFormItems(oForm);

                    oForm.Items.Item("BDOSAgrNoE").Enabled = false;
                    oForm.Items.Item("UsBlaAgRtS").Enabled = false;

                    Program.FORM_LOAD_FOR_VISIBLE = true;
                    Program.FORM_LOAD_FOR_ACTIVATE = true;
                }

                else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE && !pVal.BeforeAction)
                {
                    if (Program.FORM_LOAD_FOR_VISIBLE)
                    {
                        Program.FORM_LOAD_FOR_VISIBLE = false;
                        setVisibleFormItems(oForm);
                    }
                }

                else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE && !pVal.BeforeAction)
                {
                    if (Program.FORM_LOAD_FOR_ACTIVATE)
                    {
                        oForm.Items.Item("BDOSAddEnt").Enabled = true;
                        Program.FORM_LOAD_FOR_ACTIVATE = false;
                    }
                }

                else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && !pVal.BeforeAction)
                {
                    if (pVal.ItemUID == "BDOSAddEnt")
                    {
                        if (!pVal.InnerEvent)
                        {
                            ShowAdditionalEntries(oForm);
                            setVisibleFormItems(oForm);
                        }
                    }
                    else if (pVal.ItemUID == "36" || pVal.ItemUID == "37") //Display in და da Display in SYSC
                    {
                        setVisibleFormItems(oForm);
                    }

                    else if (pVal.ItemUID == "UsBlaAgRtS")
                    {
                        CorrectJERateWithBlanketAgreementRateRanges(oForm, true);
                    }
                }

                else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED && pVal.BeforeAction)
                {
                    if (pVal.ItemUID == "JDT1BDOS")
                    {
                        SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("JDT1BDOS").Specific;
                        SAPbouiCOM.Column oAccount = oMatrix.Columns.Item("Account");
                        SAPbouiCOM.Column oControlAccount = oMatrix.Columns.Item("ContrlAct");

                        if (oAccount.Cells.Item(pVal.Row).Specific.Value != oControlAccount.Cells.Item(pVal.Row).Specific.Value)
                        {
                            SAPbouiCOM.LinkedButton oLink = oAccount.ExtendedObject;
                            oLink.LinkedObjectType = "2";
                        }
                        else
                        {
                            SAPbouiCOM.LinkedButton oLink = oAccount.ExtendedObject;
                            oLink.LinkedObjectType = "1";
                        }
                    }
                }

                else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                {
                    var oCflEventO = (SAPbouiCOM.ChooseFromListEvent)pVal;
                    ChooseFromList(oForm, pVal, oCflEventO, out BubbleEvent);
                }

                else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_VALIDATE && !pVal.BeforeAction)
                {
                    if (pVal.ItemUID == "BDOSAgrNoE")
                    {
                        try
                        {
                            oForm.Freeze(true);

                            if (oForm.Items.Item("BDOSAgrNoE").Specific.Value.Length != 0)
                            {
                                oForm.Items.Item("UsBlaAgRtS").Enabled = true;
                            }
                            else
                            {
                                var currentEventFilters = Program.uiApp.GetFilter();
                                Program.uiApp.SetFilter(); //ივენთებს ჩახსნის რომ სხვა აითემზე დაჭერისას არ დააუსასრულლუპოს

                                oForm.Items.Item("UsBlaAgRtS").Specific.Checked = false;
                                oForm.Items.Item("7").Click();

                                Program.uiApp.SetFilter(currentEventFilters); //აქ ვაბრუნებ ისევ

                                oForm.Items.Item("UsBlaAgRtS").Enabled = false;
                            }

                            if (!pVal.InnerEvent)
                            {
                                CorrectJERateWithBlanketAgreementRateRanges(oForm, true);
                            }
                        }
                        finally
                        {
                            oForm.Freeze(false);
                        }
                    }

                    else if (pVal.ItemUID == "76" && !pVal.InnerEvent)
                    {
                        if (pVal.ColUID == "1")
                        {
                            try
                            {
                                oForm.Freeze(true);

                                var blaAgr = oForm.Items.Item("BDOSAgrNoE");
                                var isMultiBp = IsMultiBp(oForm, out var bpCode);

                                if (blaAgr.Specific.Value.Length > 0 && isMultiBp)
                                {
                                    Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("CantChooseMultiBp"), SAPbouiCOM.BoMessageTime.bmt_Short);
                                    oForm.Items.Item(pVal.ItemUID).Specific.GetCellSpecific(pVal.ColUID, pVal.Row).Value = "";
                                }

                                else if (bpCode != BlanketAgreement.GetBPByBlAgreement(blaAgr.Specific.Value) && blaAgr.Specific.Value.Length > 0)
                                {
                                    blaAgr.Specific.Value = "";
                                }

                                if (blaAgr.Enabled && blaAgr.Specific.Value.Length == 0 && bpCode.Length == 0)
                                {
                                    blaAgr.Enabled = false;
                                }

                                else if (!blaAgr.Enabled && bpCode.Length > 0)
                                {
                                    blaAgr.Enabled = true;
                                    blaAgr.Specific.ChooseFromListUID = "BlanketAgreement_Cfl";
                                    blaAgr.Specific.ChooseFromListAlias = "AbsID";
                                }

                            }
                            finally
                            {
                                oForm.Freeze(false);
                            }
                        }

                        else if (pVal.ColUID == "3" || pVal.ColUID == "4")
                        {
                            CorrectJERateWithBlanketAgreementRateRanges(oForm);
                        }
                    }

                    else if (pVal.ItemUID == "6")
                    {
                        CorrectJERateWithBlanketAgreementRateRanges(oForm);
                    }
                }
            }
        }

        public static void uiApp_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            SAPbouiCOM.Form oForm = Program.uiApp.Forms.ActiveForm;

            if (pVal.BeforeAction && pVal.MenuUID == "1284")
            {
                if (oForm.DataSources.DBDataSources.Item("OJDT").GetValue("DataSource", 0) == "O" &&
                    oForm.DataSources.DBDataSources.Item("OJDT").GetValue("Ref2", 0) != "Reconcilation" &&
                    string.IsNullOrEmpty(oForm.DataSources.DBDataSources.Item("OJDT").GetValue("VatDate", 0)))
                {
                    BubbleEvent = false;
                    throw new Exception(BDOSResources.getTranslate("YouCantCancelJournalEntry") + "!");
                }
            }
        }

        public static void UpdateJournalEntryACNumber(string DocEntry, string TransType, string ACNumber, out string errorText)
        {
            errorText = "";

            if (!string.IsNullOrEmpty(DocEntry) && !string.IsNullOrEmpty(TransType))
            {
                SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                StringBuilder query = new StringBuilder();
                query.Append("SELECT \"TransId\" \n");
                query.Append("FROM \"OJDT\" \n");
                query.Append("WHERE \"StornoToTr\" IS NULL \n");
                query.Append("AND \"TransType\" = '" + TransType + "' \n");
                query.Append("AND \"CreatedBy\" = '" + DocEntry + "' \n");
                query.Append("UNION ALL \n");
                query.Append("SELECT \"TransId\" \n");
                query.Append("FROM \"OJDT\" \n");
                query.Append("WHERE \"StornoToTr\" IS NULL \n");
                query.Append("AND \"Ref2\" = '" + TransType + "' \n");
                query.Append("AND \"Ref1\" = '" + DocEntry + "'");

                oRecordSet.DoQuery(query.ToString());

                while (!oRecordSet.EoF)
                {
                    ACNumber = string.IsNullOrEmpty(ACNumber) ? "" : ACNumber;

                    SAPbobsCOM.JournalEntries oJounalEntry = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
                    oJounalEntry.GetByKey(oRecordSet.Fields.Item("TransId").Value);
                    oJounalEntry.UserFields.Fields.Item("U_BDOSACNum").Value = ACNumber.Trim();

                    int updateCode = oJounalEntry.Update();

                    if (updateCode != 0)
                    {
                        Program.oCompany.GetLastError(out updateCode, out errorText);
                    }
                    Marshal.ReleaseComObject(oJounalEntry);

                    oRecordSet.MoveNext();
                }
            }
        }

        private static void UpdateBlanketAgreementNumber(string objectKeyXml, SAPbouiCOM.Form oForm)
        {
            var blanketAgreementNumber = oForm.Items.Item("BDOSAgrNoE").Specific.Value;
            if (string.IsNullOrEmpty(blanketAgreementNumber)) return;

            var objectKeyXmlDoc = new XmlDocument();
            objectKeyXmlDoc.LoadXml(objectKeyXml);

            var docEntry = Convert.ToInt32(objectKeyXmlDoc.InnerText);

            var oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            var updateQuery = new StringBuilder();
            try
            {
                updateQuery.Append("UPDATE OJDT \n");
                updateQuery.Append("SET \"AgrNo\" = '" + Convert.ToInt32(blanketAgreementNumber) + "' \n");
                updateQuery.Append("WHERE \"TransId\" = '" + docEntry + "'");

                oRecordSet.DoQuery(updateQuery.ToString());
            }
            catch (Exception ex)
            {
                Program.uiApp.StatusBar.SetSystemMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short);
            }
            finally
            {
                Marshal.ReleaseComObject(oRecordSet);
            }
        }

        private static bool IsMultiBp(SAPbouiCOM.Form oForm, out string bpCode)
        {
            var oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("76").Specific;
            bpCode = "";
            var bpCount = 0;

            for (var row = 1; row < oMatrix.RowCount; row++)
            {
                if (oMatrix.GetCellSpecific("1", row).Value.ToString() != oMatrix.GetCellSpecific("37", row).Value.ToString())
                {
                    if (bpCode != oMatrix.GetCellSpecific("1", row).Value.ToString())
                    {
                        bpCode = oMatrix.GetCellSpecific("1", row).Value.ToString();
                        bpCount++;
                    }
                }

                if (bpCount > 1)
                {
                    bpCode = "";
                    return true;
                }
            }

            return false;
        }

        private static void CorrectJERateWithBlanketAgreementRateRanges(SAPbouiCOM.Form oForm, bool fromCheckBox = false)
        {
            try
            {
                oForm.Freeze(true);

                if (oForm.Items.Item("UsBlaAgRtS").Specific.Checked)
                {
                    var postDate = DateTime.ParseExact(oForm.Items.Item("6").Specific.Value, "yyyyMMdd", CultureInfo.InvariantCulture);
                    var blanketAgreement = oForm.Items.Item("BDOSAgrNoE").Specific.Value;
                    var oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("76").Specific;
                    var rate = BlanketAgreement.GetBlAgremeentCurrencyRate(Convert.ToInt32(blanketAgreement), out string _, postDate);

                    for (var row = 1; row < oMatrix.RowCount; row++)
                    {
                        var debitFc = oMatrix.GetCellSpecific("3", row).Value.Length > 0
                            ? FormsB1.cleanStringOfNonDigits(oMatrix.GetCellSpecific("3", row).Value) : decimal.Zero;
                        var creditFc = oMatrix.GetCellSpecific("4", row).Value.Length > 0
                            ? FormsB1.cleanStringOfNonDigits(oMatrix.GetCellSpecific("4", row).Value) : decimal.Zero;

                        if (debitFc != decimal.Zero)
                        {
                            oMatrix.GetCellSpecific("5", row).Value = FormsB1.ConvertDecimalToStringForEditboxStrings(debitFc * rate); //Debit
                        }
                        else if (creditFc != decimal.Zero)
                        {
                            oMatrix.GetCellSpecific("6", row).Value = FormsB1.ConvertDecimalToStringForEditboxStrings(creditFc * rate); //Credit
                        }
                    }
                }

                else if (!oForm.Items.Item("UsBlaAgRtS").Specific.Checked && fromCheckBox)
                {
                    var postDate = DateTime.ParseExact(oForm.Items.Item("6").Specific.Value, "yyyyMMdd", CultureInfo.InvariantCulture);
                    var oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("76").Specific;
                    var isMultiBp = IsMultiBp(oForm, out var bpCode);

                    var currency = CommonFunctions.getBPBankInfo(bpCode)?.Fields.Item("Currency").Value;
                    if (string.IsNullOrEmpty(currency) || currency == "##" || currency == "GEL") return;

                    var oSboBob = (SAPbobsCOM.SBObob)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);
                    var rate = Convert.ToDecimal(oSboBob.GetCurrencyRate(currency, postDate).Fields.Item("CurrencyRate").Value);

                    for (var row = 1; row < oMatrix.RowCount; row++)
                    {
                        var debitFc = oMatrix.GetCellSpecific("3", row).Value.Length > 0
                            ? FormsB1.cleanStringOfNonDigits(oMatrix.GetCellSpecific("3", row).Value) : decimal.Zero;
                        var creditFc = oMatrix.GetCellSpecific("4", row).Value.Length > 0
                            ? FormsB1.cleanStringOfNonDigits(oMatrix.GetCellSpecific("4", row).Value) : decimal.Zero;

                        if (debitFc != decimal.Zero)
                        {
                            oMatrix.GetCellSpecific("5", row).Value = FormsB1.ConvertDecimalToStringForEditboxStrings(debitFc * rate); //Debit
                        }
                        else if (creditFc != decimal.Zero)
                        {
                            oMatrix.GetCellSpecific("6", row).Value = FormsB1.ConvertDecimalToStringForEditboxStrings(creditFc * rate); //Credit
                        }
                    }
                }
            }
            finally
            {
                oForm.Freeze(false);
            }
        }
    }
}