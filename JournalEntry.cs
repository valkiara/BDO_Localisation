﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Data;
using System.Globalization;

namespace BDO_Localisation_AddOn
{
    static partial class JournalEntry
    {
        public static void createUserFields( out string errorText)
        {
            errorText = null;
            Dictionary<string, object> fieldskeysMap;

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSEmpID");
            fieldskeysMap.Add("TableName", "JDT1");
            fieldskeysMap.Add("Description", "Employee ID");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            GC.Collect();

            fieldskeysMap = new Dictionary<string, object>(); //  A/C Number ივსება Good Receipt PO, AP Invoice, AP Credit memo, Landed Cost დოკუმენტებიდან
            fieldskeysMap.Add("Name", "BDOSACNum");
            fieldskeysMap.Add("TableName", "OJDT");
            fieldskeysMap.Add("Description", "A/C Number");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

        }

        public static void JrnEntry( string jeReference, string jeReference2, string remark, DateTime jeDate, DataTable jeLines, out string errorText)
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
                if(jeLine["FCCurrency"].ToString() != "" && jeLine["FCCurrency"].ToString() != Program.LocalCurrency)
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
                    oJLines.FCCurrency = isFCCurr?Program.LocalCurrency:"";
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

        public static void AddJournalEntryRow(DataTable AccountCodes, DataTable jeLines, string EntryType, string DebAccount, string CredAccount, decimal Amount, decimal FCAmount, string Currency, string DistrRule1, string DistrRule2, string DistrRule3, string DistrRule4, string DistrRule5, string PrjCode, string VatGroup, string U_BDOSEmpID)
        {

            DataRow jeLinesRow = null;
            DataRowCollection jeLinesRows = jeLines.Rows;
            //დებეტი
            if (EntryType != "OnlyCredit")
            {
                jeLinesRow = jeLinesRows.Add();
                jeLinesRow["AccountCode"] = DebAccount;
                jeLinesRow["ShortName"] = DebAccount;
                jeLinesRow["ContraAccount"] = CredAccount;
                jeLinesRow["Debit"] = Convert.ToDouble(Amount);
                jeLinesRow["FCDebit"] = Convert.ToDouble(FCAmount);
                jeLinesRow["FCCurrency"] = Currency;
                jeLinesRow["VatGroup"] = VatGroup;
                jeLinesRow["ProjectCode"] = PrjCode;
                jeLinesRow["Credit"] = 0;
                jeLinesRow["FCCredit"] = 0;

                DataRow[] oAccountCode = AccountCodes.Select("AcctCode = '" + DebAccount + "'");
                string AccountType = oAccountCode[0]["ActType"].ToString();
                string U_BDOSEmpAct = oAccountCode[0]["U_BDOSEmpAct"].ToString();

                if (AccountType != "N")
                {
                    jeLinesRow["CostingCode"] = DistrRule1;
                    jeLinesRow["CostingCode2"] = DistrRule2;
                    jeLinesRow["CostingCode3"] = DistrRule3;
                    jeLinesRow["CostingCode4"] = DistrRule4;
                    jeLinesRow["CostingCode5"] = DistrRule5;
                }
                if (U_BDOSEmpAct == "Y")
                {
                    jeLinesRow["U_BDOSEmpID"] = U_BDOSEmpID;
                }
            }

            //კტ
            if (EntryType != "OnlyDebit")
            {
                jeLinesRow = jeLinesRows.Add();
                jeLinesRow["AccountCode"] = CredAccount;
                jeLinesRow["ShortName"] = CredAccount;
                jeLinesRow["ContraAccount"] = DebAccount;
                jeLinesRow["Credit"] = Convert.ToDouble(Amount);
                jeLinesRow["FCCredit"] = Convert.ToDouble(FCAmount);
                jeLinesRow["FCCurrency"] = Currency;
                jeLinesRow["VatGroup"] = VatGroup;
                jeLinesRow["ProjectCode"] = PrjCode;
                jeLinesRow["Debit"] = 0;
                jeLinesRow["FCDebit"] = 0;

                DataRow[] oAccountCode = AccountCodes.Select("AcctCode = '" + CredAccount + "'");
                string AccountType = oAccountCode[0]["ActType"].ToString();
                string U_BDOSEmpAct = oAccountCode[0]["U_BDOSEmpAct"].ToString();
                if (AccountType != "N")
                {
                    jeLinesRow["CostingCode"] = DistrRule1;
                    jeLinesRow["CostingCode2"] = DistrRule2;
                    jeLinesRow["CostingCode3"] = DistrRule3;
                    jeLinesRow["CostingCode4"] = DistrRule4;
                    jeLinesRow["CostingCode5"] = DistrRule5;
                }
                if (U_BDOSEmpAct == "Y")
                {
                    jeLinesRow["U_BDOSEmpID"] = U_BDOSEmpID;
                }
            }
        }

        public static void createFormItems(  SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("76").Specific;
            SAPbouiCOM.Column oColumn = oMatrix.Columns.Item("U_BDOSEmpID");
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("EmployeeNo");

            SAPbouiCOM.Item oItem = oForm.Items.Item("9");

            int height = oItem.Height;
            int top = oItem.Top + oItem.Height + 5;
            int left = oItem.Left;
            int width = oItem.Width;

            //////////////////////

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
                    return;
                }
            }

            //////////////////////

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

                FormsB1.createFormItem(oForm, formItems, out errorText);
                if (errorText != null)
                {
                    return;
                }
            }

            //////////////////////

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
                    return;
                }
            }

            //AC Number
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
                return;
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
                return;
            }
            //AC Number

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
                    return;
                }
            }

            ShowAdditionalEntries( oForm, out errorText);

        }

        private static void ShowAdditionalEntries(  SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            SAPbouiCOM.DBDataSource JDT1 = oForm.DataSources.DBDataSources.Item("JDT1");
            SAPbouiCOM.DBDataSource OJDT = oForm.DataSources.DBDataSources.Item("OJDT");
            SAPbouiCOM.Item MatrixItem = oForm.Items.Item("76");
            SAPbouiCOM.Matrix oMatrix = MatrixItem.Specific;

            SAPbouiCOM.Columns oColumns = null;
            SAPbouiCOM.Column oColumn;

            SAPbobsCOM.SBObob vObj;
            vObj = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);

            if (oForm.Items.Item("BDOSAddEnt").Specific.Checked == false)
            {
                MatrixItem.Visible = true;
                //oForm.Items.Item("32").Visible = true;
                //oForm.Items.Item("34").Visible = true;
                try
                {
                    oForm.Items.Item("JDT1BDOS").Visible = false;
                }
                catch
                { }

                return;

            }
            //oForm.Items.Item("32").Visible = false;
            //oForm.Items.Item("34").Visible = false;
            oForm.Freeze(true);



            try
            {
                SAPbouiCOM.DataTable JDT1_BDO = oForm.DataSources.DataTables.Add("JDT1_BDO");

                for (int i = 0; i < JDT1.Fields.Count; i++)
                {
                    JDT1_BDO.Columns.Add(JDT1.Fields.Item(i).Name, JDT1.Fields.Item(i).Type, JDT1.Fields.Item(i).Size);
                }

                JDT1_BDO.Columns.Add("AcctName", SAPbouiCOM.BoFieldsType.ft_Text, 100);
                JDT1_BDO.Columns.Add("prFrmItm", SAPbouiCOM.BoFieldsType.ft_Text, 100);



                Dictionary<string, object> formItems = new Dictionary<string, object>();
                string itemName = "";


                itemName = "JDT1BDOS";
                formItems = new Dictionary<string, object>();
                formItems.Add("isDataSource", true);
                formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
                //formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_GRID);
                formItems.Add("Left", MatrixItem.Left);
                formItems.Add("Width", MatrixItem.Width);
                formItems.Add("Top", MatrixItem.Top);
                formItems.Add("Height", MatrixItem.Height);
                formItems.Add("UID", itemName);
                formItems.Add("FromPane", MatrixItem.FromPane);
                formItems.Add("ToPane", MatrixItem.ToPane);

                FormsB1.createFormItem(oForm, formItems, out errorText);

                SAPbouiCOM.Matrix MatrixJDT1BDOS = oForm.Items.Item("JDT1BDOS").Specific;

                oColumns = MatrixJDT1BDOS.Columns;

                oForm.Items.Item("BDOSAddEnt").Enabled = true;

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

                //ჩვენი ცხრილის შევსება
                //ძირითადი გატარება
                SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
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
                    JDT1_BDO.SetValue("prFrmItm", i, getCashFlow( JDT1_BDO.GetValue("TransId", i).ToString(), JDT1_BDO.GetValue("Line_ID", i).ToString()));

                }

                int count = JDT1_BDO.Rows.Count;

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
                        JDT1_BDO.SetValue("FCCurrency", GlCount, dtRow["FCCurrency"]);
                        JDT1_BDO.SetValue("AcctName", GlCount, getAcctName( JDT1_BDO.GetValue("Account", GlCount)));

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
                    string TransId = OJDT.GetValue("TransId", 0);

                    if (TransId != "")
                    {

                        string TransType = OJDT.GetValue("TransType", 0).Trim();
                        string CreatedBy = OJDT.GetValue("CreatedBy", 0).Trim();

                        oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        string query = @"SELECT
                                             ""JDT1"".*
                                        FROM ""JDT1""
                                        LEFT JOIN ""OJDT"" on ""OJDT"".""TransId"" = ""JDT1"".""TransId""
                                        WHERE ""OJDT"".""Ref1"" = '" + CreatedBy + @"' 
                                        AND ""OJDT"".""Ref2"" = '" + TransType + "' ";



                        oRecordSet.DoQuery(query);


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
                            JDT1_BDO.SetValue("AcctName", GlCount, getAcctName( JDT1_BDO.GetValue("Account", GlCount)));
                            JDT1_BDO.SetValue("ProfitCode", GlCount, oRecordSet.Fields.Item("ProfitCode").Value);
                            JDT1_BDO.SetValue("OcrCode2", GlCount, oRecordSet.Fields.Item("OcrCode2").Value);
                            JDT1_BDO.SetValue("OcrCode3", GlCount, oRecordSet.Fields.Item("OcrCode3").Value);
                            JDT1_BDO.SetValue("OcrCode4", GlCount, oRecordSet.Fields.Item("OcrCode4").Value);
                            JDT1_BDO.SetValue("OcrCode5", GlCount, oRecordSet.Fields.Item("OcrCode5").Value);
                            JDT1_BDO.SetValue("Project", GlCount, oRecordSet.Fields.Item("Project").Value);
                            JDT1_BDO.SetValue("prFrmItm", GlCount, getCashFlow( JDT1_BDO.GetValue("TransId", GlCount).ToString(), JDT1_BDO.GetValue("Line_ID", GlCount).ToString()));
                            JDT1_BDO.SetValue("U_BDOSEmpID", GlCount, oRecordSet.Fields.Item("U_BDOSEmpID").Value);
                            GlCount++;
                            oRecordSet.MoveNext();
                        }

                    }
                }

                MatrixJDT1BDOS = oForm.Items.Item("JDT1BDOS").Specific;
                MatrixJDT1BDOS.Clear();
                MatrixJDT1BDOS.LoadFromDataSource();
                MatrixJDT1BDOS.AutoResizeColumns();


            }
            catch (Exception ex)
            {
                 errorText = ex.Message;
                //Program.uiApp.StatusBar.SetSystemMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);

            }


            SAPbouiCOM.Item MatrixItemJDT1BDOS = oForm.Items.Item("JDT1BDOS");

            MatrixItem.Visible = false;
            MatrixItemJDT1BDOS.Visible = true;
            setVisibility(oForm, out errorText);

            oForm.Freeze(false);

        }

        public static string getAcctName( string Account)
        {

            SAPbobsCOM.Recordset oRecordSetOACT = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string queryOACT = @"SELECT * FROM ""OACT""
                                             WHERE ""AcctCode"" = '" + Account + "'";

            oRecordSetOACT.DoQuery(queryOACT);

            if (!oRecordSetOACT.EoF)
            {
                return oRecordSetOACT.Fields.Item("AcctName").Value;
            }
            return "";

        }

        public static string getCashFlow( string TransId, string Line_ID)
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

        public static void formDataLoad( SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            string TransType = oForm.DataSources.DBDataSources.Item("OJDT").GetValue("TransType", 0).Trim();
            string CreatedBy = oForm.DataSources.DBDataSources.Item("OJDT").GetValue("CreatedBy", 0).Trim();
            string StornoToTr = oForm.DataSources.DBDataSources.Item("OJDT").GetValue("StornoToTr", 0).Trim();

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string query = "";
            if (StornoToTr == "")
            {
                query = @"SELECT 
                            *  
                            FROM ""OJDT"" 
                            WHERE ""StornoToTr"" IS NULL   
                            AND ""Ref1"" = '" + CreatedBy + @"'  
                            AND ""Ref2"" = '" + TransType + "' ";
            }
            else
            {
                query = @"SELECT 
                            *  
                            FROM ""OJDT"" 
                            WHERE ""StornoToTr"" IS NOT NULL   
                            AND ""Ref1"" = '" + CreatedBy + @"'
                            AND ""Ref2"" = '" + TransType + "' ";
            }

            oRecordSet.DoQuery(query);

            if (!oRecordSet.EoF)
            {
                oForm.Items.Item("BDOSJrnEnt").Specific.Value = oRecordSet.Fields.Item("TransId").Value;
            }
            else
            {
                oForm.Items.Item("BDOSJrnEnt").Specific.Value = "";
            }
        }

        public static void uiApp_FormDataEvent(  ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(BusinessObjectInfo.FormTypeEx, Program.currentFormCount);

            if (oForm.TypeEx == "392" & BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD & BusinessObjectInfo.BeforeAction == false)
            {
                JournalEntry.formDataLoad( oForm, out errorText);
            }
        }

        public static void cancellation( SAPbouiCOM.Form oForm, int Ref1, string Ref2, out string errorText)
        {
            errorText = null;
            SAPbobsCOM.JournalEntries oJounalEntry = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);

            int TransId = 0;

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string queryHeader = "SELECT " +
                            "*  " +
                            "FROM \"OJDT\"  " +
                            "WHERE \"Ref1\" = '" + Ref1.ToString() + "' " +
                            "AND \"Ref2\" = '" + Ref2 + "' ";

            oRecordSet.DoQuery(queryHeader);
            if (!oRecordSet.EoF)
            {
                TransId = oRecordSet.Fields.Item("TransId").Value;

                SAPbobsCOM.JournalEntries oJournalDoc = (SAPbobsCOM.JournalEntries)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
                oJournalDoc.GetByKey(TransId);
                int response = oJournalDoc.Cancel();

            }
        }

        public static void setVisibility(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            oForm.Freeze(true);

            SAPbouiCOM.CheckBox oCheckBoxFC = (SAPbouiCOM.CheckBox)oForm.Items.Item("37").Specific;
            SAPbouiCOM.CheckBox oCheckBoxLC = (SAPbouiCOM.CheckBox)oForm.Items.Item("36").Specific;

            bool visibleFC = oCheckBoxFC.Checked;
            bool visibleLC = oCheckBoxLC.Checked;

            try
            {
                SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("JDT1BDOS").Specific;
                oMatrix.Columns.Item("FCDebit").Visible = visibleFC;
                oMatrix.Columns.Item("FCCredit").Visible = visibleFC;
                oMatrix.Columns.Item("FCCurrency").Visible = visibleFC;

                oMatrix.Columns.Item("SYSCred").Visible = visibleLC;
                oMatrix.Columns.Item("SYSDeb").Visible = visibleLC;

                oMatrix.AutoResizeColumns();
            }
            catch
            { }



            oForm.Freeze(false);

        }

        public static void uiApp_ItemEvent(  string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD & pVal.BeforeAction == true)
                {
                    JournalEntry.createFormItems( oForm, out errorText);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE & pVal.BeforeAction == false)
                {
                    JournalEntry.formDataLoad( oForm, out errorText);
                    setVisibility(oForm, out errorText);
                }

                if (pVal.ItemUID == "BDOSAddEnt")
                {
                    if (pVal.ItemUID == "BDOSAddEnt" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.InnerEvent == false && pVal.BeforeAction == false)
                    {

                        ShowAdditionalEntries( oForm, out errorText);
                    }
                }

                try
                {
                    oForm.Items.Item("BDOSAddEnt").Enabled = true;
                }
                catch
                { }


                //Display in და da Display in SYSC
                if ((pVal.ItemUID == "36" || pVal.ItemUID == "37") && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.BeforeAction == false)
                {
                    setVisibility(oForm, out errorText);
                }

                if (pVal.ItemUID == "JDT1BDOS" & pVal.ColUID == "Account" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED && pVal.BeforeAction == true)
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
        }

        public static void UpdateJournalEntryACNumber(string DocEntry, string TransType, string ACNumber, out string errorText)
        {
            errorText = "";

            if (DocEntry != "" && TransType != "" && ACNumber != "")
            {
                SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string query = @"SELECT 
                            *  
                            FROM ""OJDT"" 
                            WHERE ""StornoToTr"" IS NULL   
                            AND ""TransType"" = '" + TransType + @"'  
                            AND ""CreatedBy"" = '" + DocEntry + "' ";

                oRecordSet.DoQuery(query);

                if (!oRecordSet.EoF)
                {
                    SAPbobsCOM.JournalEntries oJounalEntry = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
                    oJounalEntry.GetByKey(oRecordSet.Fields.Item("TransId").Value);

                    oJounalEntry.UserFields.Fields.Item("U_BDOSACNum").Value = ACNumber.Trim();

                    int updateCode = 0;
                    updateCode = oJounalEntry.Update();

                    if (updateCode != 0)
                    {
                        Program.oCompany.GetLastError(out updateCode, out errorText);
                    }
                }
            }
        }

    }
}