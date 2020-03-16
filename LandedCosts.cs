using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Resources;
using System.Globalization;
using System.Data;

namespace BDO_Localisation_AddOn
{
    static partial class LandedCosts
    {
        public static CultureInfo cultureInfo = null;

        public static void CheckAccounts(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("51").Specific));
            int rowCount = oMatrix.RowCount;
            string vatGrp = null;
            bool isError = false;

            for (int row = 1; row <= rowCount; row++)
            {
                vatGrp = oMatrix.Columns.Item("BDOSVatGrp").Cells.Item(row).Specific.Value;

                SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string query = "SELECT " +
                                "* " +
                                "FROM \"OVTG\" " +
                                "WHERE \"OVTG\".\"Code\"='" + vatGrp + "'";

                oRecordSet.DoQuery(query);

                if (!oRecordSet.EoF)
                {
                    if (oRecordSet.Fields.Item("U_BDOSAccF").Value == "" || oRecordSet.Fields.Item("Account").Value == "" || oRecordSet.Fields.Item("U_BDOSAccCVt").Value == "")
                    {
                        isError = true;
                    }
                }

                if (isError == true)
                {
                    errorText = BDOSResources.getTranslate("CheckVatGroupAccounts");
                }
            }
        }

        public static DataTable createAdditionalEntries(SAPbouiCOM.Form oForm, SAPbobsCOM.GeneralData oGeneralData, DataTable DTSource, out DataTable reLines)
        {
            DataTable jeLines = JournalEntry.JournalEntryTable();
            DataRow jeLinesRow = null;

            reLines = ProfitTax.ProfitTaxTable();


            SAPbouiCOM.DBDataSource DBDataSourceTable = null;
            SAPbouiCOM.DBDataSource DBDataSource = null;
            int JEcount = 0;

            if (oForm == null)
            {
                JEcount = DTSource.Rows.Count;
            }
            else
            {
                DBDataSourceTable = oForm.DataSources.DBDataSources.Item("IPF1");
                DBDataSource = oForm.DataSources.DBDataSources.Item("OIPF");
                JEcount = DBDataSourceTable.Size;
            }

            decimal VatAmount = FormsB1.cleanStringOfNonDigits(CommonFunctions.getChildOrDbDataSourceValue(DBDataSource, null, DTSource, "U_BDOSAVtAmt", 0).ToString());
            string CreditAccount = "";
            string DebitAccount = "";
            string CustomVatAccount = "";

            string vatCode = "";
            string BaseEntry = "";
            string BaseType = "";


            //დღგ-ის გატარება
            for (int i = 0; i < JEcount; i++)
            {
                vatCode = CommonFunctions.getChildOrDbDataSourceValue(DBDataSourceTable, null, DTSource, "U_BDOSVatGrp", i).ToString();
                BaseEntry = CommonFunctions.getChildOrDbDataSourceValue(DBDataSourceTable, null, DTSource, "BaseEntry", i).ToString();
                BaseType = CommonFunctions.getChildOrDbDataSourceValue(DBDataSourceTable, null, DTSource, "BaseType", i).ToString();

                SAPbobsCOM.VatGroups oVatCode;
                oVatCode = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVatGroups);
                oVatCode.GetByKey(vatCode);

                CreditAccount = oVatCode.UserFields.Fields.Item("U_BDOSAccF").Value;
                DebitAccount = oVatCode.TaxAccount;
                CustomVatAccount = oVatCode.UserFields.Fields.Item("U_BDOSAccCVt").Value;


            }

            if (BaseType == "69")
            {
                SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string query = @"SELECT ""U_BDOSAVtAmt"" FROM ""OIPF""
                                WHERE ""DocEntry"" = " + BaseEntry;

                oRecordSet.DoQuery(query);

                if (!oRecordSet.EoF)
                {
                    VatAmount = VatAmount - Convert.ToDecimal(oRecordSet.Fields.Item("U_BDOSAVtAmt").Value);
                }
            }

            if (VatAmount == 0)
            {
                return jeLines;
            }

            jeLinesRow = jeLines.Rows.Add();
            jeLinesRow["AccountCode"] = CreditAccount;
            jeLinesRow["ShortName"] = CreditAccount;
            jeLinesRow["ContraAccount"] = DebitAccount;
            jeLinesRow["Credit"] = 0;
            jeLinesRow["Debit"] = VatAmount;
            jeLinesRow["FCCurrency"] = "";

            jeLinesRow = jeLines.Rows.Add();
            jeLinesRow["AccountCode"] = DebitAccount;
            jeLinesRow["ShortName"] = DebitAccount;
            jeLinesRow["ContraAccount"] = CreditAccount;
            jeLinesRow["Credit"] = VatAmount;
            jeLinesRow["Debit"] = 0;
            jeLinesRow["FCCurrency"] = "";


            jeLinesRow = jeLines.Rows.Add();
            jeLinesRow["AccountCode"] = DebitAccount;
            jeLinesRow["ShortName"] = DebitAccount;
            jeLinesRow["ContraAccount"] = CustomVatAccount;
            jeLinesRow["Credit"] = 0;
            jeLinesRow["Debit"] = VatAmount;
            jeLinesRow["FCCurrency"] = "";

            jeLinesRow = jeLines.Rows.Add();
            jeLinesRow["AccountCode"] = CustomVatAccount;
            jeLinesRow["ShortName"] = CustomVatAccount;
            jeLinesRow["ContraAccount"] = DebitAccount;
            jeLinesRow["Credit"] = VatAmount;
            jeLinesRow["Debit"] = 0;
            jeLinesRow["FCCurrency"] = "";

            return jeLines;

        }

        public static void ChangeVatPercent(SAPbouiCOM.Form oForm, int row, out string errorText)
        {
            errorText = null;

            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("51").Specific));

            oMatrix.Columns.Item("BDOSVatPrc").Cells.Item(row).Specific.Value = FormsB1.ConvertDecimalToString(GetVatGroupRate(oMatrix.Columns.Item("BDOSVatGrp").Cells.Item(row).Specific.Value));
            RefillRowVatAmount(oForm, oMatrix, row);
        }

        public static decimal GetVatGroupRate(string VatGroup)
        {

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string query = "SELECT " +
                            "* " +
                            "FROM \"OVTG\" " +
                            "WHERE \"OVTG\".\"Code\"='" + VatGroup + "'";

            oRecordSet.DoQuery(query);

            if (!oRecordSet.EoF)
            {
                return Convert.ToDecimal(oRecordSet.Fields.Item("Rate").Value);
            }
            else
            {
                return 0;
            }
        }

        public static void RefillVatAmounts(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;


            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("51").Specific));
            int rowCount = oMatrix.RowCount;

            for (int row = 1; row <= rowCount; row++)
            {
                RefillRowVatAmount(oForm, oMatrix, row);
            }

            RefillTotalVatAmounts(oForm, out errorText);

        }

        public static void RefillTotalVatAmounts(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;


            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("51").Specific));
            int rowCount = oMatrix.RowCount;
            decimal totalVatamount = 0;

            for (int row = 1; row <= rowCount; row++)
            {
                totalVatamount = totalVatamount + FormsB1.cleanStringOfNonDigits(oMatrix.Columns.Item("BDOSVatAmt").Cells.Item(row).Specific.Value);
            }

            oForm.Items.Item("BDOSVatAmt").Specific.Value = FormsB1.ConvertDecimalToString(totalVatamount);
            oForm.Items.Item("BDOSAVtAmt").Specific.Value = FormsB1.ConvertDecimalToString(totalVatamount);
        }

        public static void FillTaxCodes(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("51").Specific));
            int rowCount = oMatrix.RowCount;

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string query = "SELECT " +
                            "\"OITM\".\"ItemCode\", " +
                            "\"OITM\".\"VatGroupPu\", " +
                            "\"OVTG\".\"Rate\" " +
                            "FROM \"OITM\" " +
                            "LEFT JOIN \"OVTG\" ON \"OITM\".\"VatGroupPu\" = \"OVTG\".\"Code\" ";

            query = query + "WHERE \"ItemCode\" IN ( ";
            for (int row = 1; row <= rowCount; row++)
            {
                query = query + "'" + oMatrix.Columns.Item("1").Cells.Item(row).Specific.Value + "'";

                if (row < rowCount)
                {
                    query = query + ",";
                }
            }
            if (rowCount == 0)
            {
                query = query + "'0'";
            }

            query = query + ")";
            oRecordSet.DoQuery(query);


            for (int row = 1; row <= rowCount; row++)
            {
                string itemCode = oMatrix.Columns.Item("1").Cells.Item(row).Specific.Value;

                while (!oRecordSet.EoF && itemCode != oRecordSet.Fields.Item("ItemCode").Value)
                {
                    oRecordSet.MoveNext();
                }

                if (oRecordSet.Fields.Item("ItemCode").Value == itemCode)
                {
                    oMatrix.Columns.Item("BDOSVatGrp").Cells.Item(row).Specific.Select(oRecordSet.Fields.Item("VatGroupPu").Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                    oMatrix.Columns.Item("BDOSVatPrc").Cells.Item(row).Specific.Value = oRecordSet.Fields.Item("Rate").Value;

                    oRecordSet.MoveFirst();
                }
            }

            RefillVatAmounts(oForm, out errorText);

        }

        public static void RefillRowVatAmount(SAPbouiCOM.Form oForm, SAPbouiCOM.Matrix oMatrix, int row)
        {

            try
            {

                //-------- get currency and date

                SAPbouiCOM.DBDataSource DBDataSourceO = oForm.DataSources.DBDataSources.Item("OIPF");
                string DocDateStr = DBDataSourceO.GetValue("DocDate", 0);
                DateTime DocDate = DateTime.TryParseExact(DocDateStr, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DocDate) == false ? DateTime.Now : DocDate;

                string DocCurr = "";
                string BaseDocPrice = oMatrix.Columns.Item("7").Cells.Item(row).Specific.String;

                SAPbobsCOM.Recordset oRecordSetCur = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string query = "SELECT " +
                                "\"OCRN\".\"CurrCode\" " +
                                "FROM \"OCRN\" ";
                oRecordSetCur.DoQuery(query);

                while (!oRecordSetCur.EoF)
                {
                    string curr = oRecordSetCur.Fields.Item("CurrCode").Value;

                    if (BaseDocPrice.Contains(curr))
                    {
                        DocCurr = curr;
                    }

                    oRecordSetCur.MoveNext();
                }

                //-------- get rate
                decimal DocRate = 0;
                SAPbobsCOM.SBObob oSBOBob = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);
                if (DocCurr == "GEL")
                {
                    DocRate = 1;
                }
                else
                {
                    SAPbobsCOM.Recordset RateRecordset = oSBOBob.GetCurrencyRate(DocCurr, DocDate);
                    while (!RateRecordset.EoF)
                    {
                        DocRate = Convert.ToDecimal(RateRecordset.Fields.Item("CurrencyRate").Value);
                        RateRecordset.MoveNext();
                    }

                }
                //=========================================

                string itemCode = oMatrix.Columns.Item("1").Cells.Item(row).Specific.Value;
                decimal duty = 0;

                SAPbobsCOM.Recordset oRecordSetDut = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                string queryDut =
                                "SELECT " +
                                "\"OARG\".\"TotalTax\" " +
                                "FROM \"OARG\" " +
                                "JOIN \"OITM\" " +
                                "ON \"OARG\".\"CstGrpCode\" = \"OITM\".\"CstGrpCode\" " +
                                "WHERE \"OITM\".\"ItemCode\" = '" + itemCode + "'";

                oRecordSetDut.DoQuery(queryDut);

                while (!oRecordSetDut.EoF)
                {
                    duty = (decimal)oRecordSetDut.Fields.Item("TotalTax").Value;
                    oRecordSetDut.MoveNext();
                }

                decimal dutyValue = ((FormsB1.cleanStringOfNonDigits(oMatrix.Columns.Item("3").Cells.Item(row).Specific.Value)
                                * FormsB1.cleanStringOfNonDigits(oMatrix.Columns.Item("7").Cells.Item(row).Specific.Value)
                                * DocRate)
                                + FormsB1.cleanStringOfNonDigits(oMatrix.Columns.Item("10000102").Cells.Item(row).Specific.Value)
                                - FormsB1.cleanStringOfNonDigits(oMatrix.Columns.Item("10000058").Cells.Item(row).Specific.Value))
                                * (duty / 100);
                decimal duTyRounded = decimal.Round(dutyValue, 2);

                if (oMatrix.Columns.Item("10000074").Visible)
                    oMatrix.Columns.Item("10000074").Cells.Item(row).Specific.Value = FormsB1.ConvertDecimalToStringForEditboxStrings(duTyRounded);

                //---------------------------------
                decimal LineTotal = FormsB1.cleanStringOfNonDigits(oMatrix.Columns.Item("10000074").Cells.Item(row).Specific.Value)
                    + FormsB1.cleanStringOfNonDigits(oMatrix.Columns.Item("10000102").Cells.Item(row).Specific.Value)
                    - FormsB1.cleanStringOfNonDigits(oMatrix.Columns.Item("10000058").Cells.Item(row).Specific.Value)
                    + (FormsB1.cleanStringOfNonDigits(oMatrix.Columns.Item("3").Cells.Item(row).Specific.Value) * DocRate * FormsB1.cleanStringOfNonDigits(oMatrix.Columns.Item("7").Cells.Item(row).Specific.Value));
                decimal VatPercent = FormsB1.cleanStringOfNonDigits(oMatrix.Columns.Item("BDOSVatPrc").Cells.Item(row).Specific.Value);
                decimal VatAmount = LineTotal * (VatPercent / 100);

                oMatrix.Columns.Item("BDOSVatAmt").Cells.Item(row).Specific.Value = FormsB1.ConvertDecimalToString(VatAmount);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        public static void FillCostsAmounts(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string query = "";

            SAPbouiCOM.DBDataSources dBDataSources = oForm.DataSources.DBDataSources;

            try
            {
                string acNumber = dBDataSources.Item("OIPF").GetValue("U_BDOSACNum", 0);
                if (string.IsNullOrEmpty(acNumber))
                {
                    if (dBDataSources.Item("IPF1").Size > 0)
                    {
                        string baseEntry = dBDataSources.Item("IPF1").GetValue("BaseEntry", 0);
                        string baseType = dBDataSources.Item("IPF1").GetValue("BaseType", 0);
                        if (baseEntry != "" && (baseType == "18" || baseType == "20" || baseType == "69"))
                        {
                            string tableName = (baseType == "18" ? "OPCH" : (baseType == "20" ? "OPDN" : "OIPF"));

                            query = @"SELECT
                                 ""U_BDOSACNum""
                            FROM """ + tableName + @"""
                            WHERE ""DocEntry"" = '" + baseEntry + "'";

                            oRecordSet.DoQuery(query);
                            if (!oRecordSet.EoF)
                            {
                                acNumber = oRecordSet.Fields.Item("U_BDOSACNum").Value;
                                SAPbouiCOM.EditText oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("BDOSACNumE").Specific;
                                oEditText.Value = acNumber;
                            }
                        }
                    }
                }

                if (!string.IsNullOrEmpty(acNumber))
                {
                    SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("54").Specific));
                    SAPbouiCOM.Column oColumnAlcCode = oMatrix.Columns.Item("5");
                    int rowCount = oMatrix.RowCount;

                    query = @"SELECT
                            ""OALC"".""AlcCode"",
	                        ""JDT1"".""Account"",
	                        SUM(""JDT1"".""Debit""),
	                        SUM(""JDT1"".""Credit""),
	                        SUM(""JDT1"".""Debit"") - SUM(""JDT1"".""Credit"") AS ""Amount""
                        FROM ""JDT1""
                        INNER JOIN ""OJDT"" ON ""JDT1"".""TransId"" = ""OJDT"".""TransId"" AND ""OJDT"".""U_BDOSACNum"" = '" + acNumber.Trim() + @"'
                        INNER JOIN ""OALC"" ON ""JDT1"".""Account"" = ""OALC"".""LaCAllcAcc""

                        GROUP BY ""OALC"".""AlcCode"", ""JDT1"".""Account"" ";

                    oRecordSet.DoQuery(query);
                    while (!oRecordSet.EoF)
                    {
                        decimal amount = Convert.ToDecimal(oRecordSet.Fields.Item("Amount").Value);
                        string alcCode = oRecordSet.Fields.Item("AlcCode").Value;
                        if (amount > 0)
                        {
                            for (int row = 1; row <= rowCount; row++)
                            {
                                if (alcCode == oColumnAlcCode.Cells.Item(row).Specific.Value)
                                {
                                    oMatrix.Columns.Item("3").Cells.Item(row).Specific.Value = FormsB1.ConvertDecimalToStringForEditboxStrings(amount);
                                }

                            }
                        }

                        oRecordSet.MoveNext();
                    }
                }
            }
            catch { }
            { }
        }

        public static void createFormItems(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            Dictionary<string, object> formItems = new Dictionary<string, object>();

            SAPbouiCOM.Item oItem = oForm.Items.Item("62");
            int height = oItem.Height;
            int top = oForm.Items.Item("18").Top;
            int left_s = oItem.Left;
            int width_s = oItem.Width;

            int left_e = oForm.Items.Item("61").Left;
            int width_e = oForm.Items.Item("61").Width;

            string itemName = "BDOSVtAmS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("ProjectVatAmount"));
            //formItems.Add("TextStyle", 4);
            formItems.Add("FontSize", 10);
            formItems.Add("Enabled", true);
            formItems.Add("LinkTo", "BDOSVatAmt");
            formItems.Add("FromPane", 1);
            formItems.Add("ToPane", 1);


            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "BDOSVatAmt"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OIPF");
            formItems.Add("Alias", "U_BDOSVatAmt");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", false);
            formItems.Add("FromPane", 1);
            formItems.Add("ToPane", 1);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            itemName = "BDOSFllVAT";
            formItems = new Dictionary<string, object>();
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            formItems.Add("Left", left_s - width_s - 2);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top + height);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("FillVatRates"));
            formItems.Add("FromPane", 1);
            formItems.Add("ToPane", 1);


            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "BDOSVatAcS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top + height);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("ActualVatAmount"));
            //formItems.Add("TextStyle", 4);
            formItems.Add("FontSize", 10);
            formItems.Add("Enabled", true);
            formItems.Add("LinkTo", "BDOSVatAmt");
            formItems.Add("FromPane", 1);
            formItems.Add("ToPane", 1);


            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "BDOSAVtAmt"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OIPF");
            formItems.Add("Alias", "U_BDOSAVtAmt");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top + height);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", true);
            formItems.Add("FromPane", 1);
            formItems.Add("ToPane", 1);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("51").Specific));
            SAPbouiCOM.Columns oColumns = oMatrix.Columns;
            SAPbouiCOM.Column oColumn;

            oMatrix.Clear();

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string query = "select * " +
            "FROM  \"OVTG\" " +
            "WHERE \"Category\"='I'";

            oRecordSet.DoQuery(query);

            oColumn = oColumns.Add("BDOSVatGrp", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("VatGroup");
            oColumn.DataBind.SetBound(true, "IPF1", "U_BDOSVatGrp");
            oColumn.Width = 60;
            oColumn.Editable = true;
            oColumn.Visible = true;
            oColumn.DisplayDesc = true;
            oColumn.ExpandType = SAPbouiCOM.BoExpandType.et_ValueDescription;
            while (!oRecordSet.EoF)
            {
                oColumn.ValidValues.Add(oRecordSet.Fields.Item("Code").Value, oRecordSet.Fields.Item("Name").Value);
                oRecordSet.MoveNext();
            }

            oColumn = oColumns.Add("BDOSVatPrc", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("VatPercent");
            oColumn.DataBind.SetBound(true, "IPF1", "U_BDOSVatPrc");
            oColumn.Width = 60;
            oColumn.Editable = false;
            oColumn.Visible = true;

            oColumn = oColumns.Add("BDOSVatAmt", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("VatAmount");
            oColumn.DataBind.SetBound(true, "IPF1", "U_BDOSVatAmt");
            oColumn.Width = 60;
            oColumn.Editable = true;
            oColumn.Visible = true;

            oItem = oForm.Items.Item("36");
            top = oItem.Top + height * 2 + 1;
            left_s = oItem.Left;
            width_s = oItem.Width;
            oItem = oForm.Items.Item("3");
            left_e = oItem.Left;
            width_e = oItem.Width;

            formItems = new Dictionary<string, object>();
            itemName = "BDOSACNumS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
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
            formItems.Add("TableName", "OIPF");
            formItems.Add("Alias", "U_BDOSACNum");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", true);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }


            oItem = oForm.Items.Item("BDOSACNumS");
            top = oItem.Top + height + 5;

            string caption = BDOSResources.getTranslate("StockRevaluation");
            formItems = new Dictionary<string, object>();
            itemName = "BDOSStRev"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s * 1.5);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", caption);
            formItems.Add("TextStyle", 4);
            formItems.Add("FontSize", 10);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            bool multiSelection = false;
            string objectType = "162"; //Waybill document
            string uniqueID_WaybillCFL = "UDO_F_BDO_STRV_D";
            FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_WaybillCFL);

            formItems = new Dictionary<string, object>();
            itemName = "StockRevE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "UserDataSources");
            formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
            formItems.Add("Length", 11);
            formItems.Add("TableName", "");
            formItems.Add("Alias", itemName);
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e + width_e - 40);
            formItems.Add("Width", 40);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("AffectsFormMode", false);
            formItems.Add("Enabled", false);
            formItems.Add("ChooseFromListUID", uniqueID_WaybillCFL);
            formItems.Add("ChooseFromListAlias", "DocEntry");
            
            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }
            oForm.Items.Item(itemName).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);


            formItems = new Dictionary<string, object>();
            itemName = "BDO_WblLB"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            formItems.Add("Left", left_e + width_e - 40 - 20);
            formItems.Add("Top", top);
            formItems.Add("Height", 13);
            formItems.Add("UID", itemName);
            formItems.Add("LinkTo", "StockRevE");
            formItems.Add("LinkedObjectType", "162");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            oForm.DataSources.UserDataSources.Add("BDO_WblID", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
            oForm.DataSources.UserDataSources.Add("BDO_WblNum", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
            oForm.DataSources.UserDataSources.Add("BDO_WblSts", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);

            oItem = oForm.Items.Item("68");
            itemName = "BDOSFllAmt";
            formItems = new Dictionary<string, object>();
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            formItems.Add("Left", oItem.Left);
            formItems.Add("Width", oItem.Width);
            formItems.Add("Top", oItem.Top + oItem.Height + 1);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Fill"));
            formItems.Add("FromPane", 6);
            formItems.Add("ToPane", 6);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            GC.Collect();
        }

        public static void createUserFields(out string errorText)
        {
            errorText = null;
            Dictionary<string, object> fieldskeysMap;

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSVatAmt");
            fieldskeysMap.Add("TableName", "OIPF");
            fieldskeysMap.Add("Description", "Project Vat Amount");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSAVtAmt");
            fieldskeysMap.Add("TableName", "OIPF");
            fieldskeysMap.Add("Description", "Actual Vat Amount");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //  A/C Number
            fieldskeysMap.Add("Name", "BDOSACNum");
            fieldskeysMap.Add("TableName", "OIPF");
            fieldskeysMap.Add("Description", "A/C Number");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSVatGrp");
            fieldskeysMap.Add("TableName", "IPF1");
            fieldskeysMap.Add("Description", "Vat Group");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 100);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSVatPrc");
            fieldskeysMap.Add("TableName", "IPF1");
            fieldskeysMap.Add("Description", "Vat Percent");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);


            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSVatAmt");
            fieldskeysMap.Add("TableName", "IPF1");
            fieldskeysMap.Add("Description", "Vat Amount");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            GC.Collect();
        }

        public static void resizeForm(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            try
            {
                reArrangeFormItems(oForm);
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
        }

        public static void reArrangeFormItems(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.Item oItem = null;
            int height = 15;
            int top = 6;
            top = top + height + 1;

            oItem = oForm.Items.Item("BDOSVtAmS");
            oItem.Left = oForm.Items.Item("62").Left;

            oItem = oForm.Items.Item("BDOSFllVAT");
            oItem.Left = oForm.Items.Item("BDOSVatAcS").Left - oItem.Width - 2;

        }

        public static void uiApp_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(BusinessObjectInfo.FormTypeEx, Program.currentFormCount);

            if (oForm.TypeEx == "992")
            {             
                if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD)
                {
                    CheckAccounts(oForm, out errorText);
                    if (errorText != null)
                    {
                        Program.uiApp.MessageBox(errorText);
                        BubbleEvent = false;
                    }


                    //დოკუმენტის გატარების დროს გატარდეს ბუღლტრული გატარება
                    //if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD && BusinessObjectInfo.ActionSuccess && !BusinessObjectInfo.BeforeAction)
                    //{
                    //    SAPbouiCOM.DBDataSource DocDBSourceOCRD = oForm.DataSources.DBDataSources.Item(0);
                    //    string DocEntry = DocDBSourceOCRD.GetValue("DocEntry", 0);
                    //    string DocNum = DocDBSourceOCRD.GetValue("DocNum", 0);
                    //    string DocCurr = DocDBSourceOCRD.GetValue("DocCurr", 0);
                    //    decimal DocRate = FormsB1.cleanStringOfNonDigits( DocDBSourceOCRD.GetValue("DocRate", 0));
                    //    DateTime DocDate = DateTime.ParseExact(DocDBSourceOCRD.GetValue("DocDate", 0), "yyyyMMdd", CultureInfo.InvariantCulture);
                    //    JrnEntry( DocEntry, DocNum, DocDate, DocRate, DocCurr, out errorText);
                    //    if (errorText != null)
                    //    {
                    //        Program.uiApp.MessageBox(errorText);
                    //        BubbleEvent = false;
                    //    }
                    //}
                    if (BusinessObjectInfo.ActionSuccess != BusinessObjectInfo.BeforeAction)
                    {
                        //დოკუმენტის გატარების დროს გატარდეს ბუღლტრული გატარება
                        SAPbouiCOM.DBDataSource DocDBSourcePAYR = oForm.DataSources.DBDataSources.Item(0);

                        if (DocDBSourcePAYR.GetValue("CANCELED", 0) == "N")
                        {

                            string DocEntry = DocDBSourcePAYR.GetValue("DocEntry", 0);


                            string DocNum = DocDBSourcePAYR.GetValue("DocNum", 0);
                            DateTime DocDate = DateTime.ParseExact(DocDBSourcePAYR.GetValue("DocDate", 0), "yyyyMMdd", CultureInfo.InvariantCulture);

                            CommonFunctions.StartTransaction();

                            Program.JrnLinesGlobal = new DataTable();
                            DataTable reLines = null;
                            DataTable JrnLinesDT = createAdditionalEntries(oForm, null, null, out reLines);

                            JrnEntry(DocEntry, DocNum, DocDate, JrnLinesDT, reLines, out errorText);
                            if (errorText != null)
                            {
                                Program.uiApp.MessageBox(errorText);
                                BubbleEvent = false;
                            }
                            else
                            {
                                if (BusinessObjectInfo.ActionSuccess == false)
                                {
                                    Program.JrnLinesGlobal = JrnLinesDT;
                                }
                            }

                            //თუ დოკუმენტი გატარდა, მერე ვაკეთებს ბუღალტრულ გატარებას
                            if (BusinessObjectInfo.ActionSuccess == true && !BusinessObjectInfo.BeforeAction)
                            {
                                CommonFunctions.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                            }
                            else
                            {
                                CommonFunctions.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                            }
                        }
                    }
                }

                if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD & !BusinessObjectInfo.BeforeAction)
                {
                    formDataLoad(oForm);
                    setVisibleFormItems(oForm);
                } 

                //A/C Number Update
                if ((BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD || BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE)
                    && BusinessObjectInfo.ActionSuccess == true && !BusinessObjectInfo.BeforeAction)
                {
                    CommonFunctions.StartTransaction();

                    SAPbouiCOM.DBDataSource DocDBSource = oForm.DataSources.DBDataSources.Item(0);
                    string DocEntry = DocDBSource.GetValue("DocEntry", 0);
                    string ObjType = DocDBSource.GetValue("ObjType", 0);
                    string ACNumber = DocDBSource.GetValue("U_BDOSACNum", 0);

                    JournalEntry.UpdateJournalEntryACNumber(DocEntry, ObjType, ACNumber, out errorText);
                    if (string.IsNullOrEmpty(errorText))
                    {
                        CommonFunctions.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                    }
                    else
                    {
                        Program.uiApp.MessageBox(errorText);
                        CommonFunctions.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                    }
                }
            }
        }

        public static void JrnEntry(string DocEntry, string DocNum, DateTime DocDate, DataTable JrnLinesDT, DataTable reLines, out string errorText)
        {
            errorText = null;

            try
            {

                //JournalEntry.JrnEntry( DocEntry, "69", "Landed costs: " + DocNum, DocDate, jeLines, rate, currency, out errorText);
                JournalEntry.JrnEntry(DocEntry, "69", "Landed costs: " + DocNum, DocDate, JrnLinesDT, out errorText);


            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
        }

        public static void setVisibleFormItems(SAPbouiCOM.Form oForm)
        {
            oForm.Freeze(true);

            try
            {
                //   if (oForm.PaneLevel == 6 || oForm.PaneLevel == 7)
                oForm.Items.Item("68").Visible = false;
                string docEntrySTR = oForm.DataSources.DBDataSources.Item("OIPF").GetValue("DocEntry", 0).Trim();
                bool docEntryIsEmpty = string.IsNullOrEmpty(docEntrySTR);

                oForm.Items.Item("68").Enabled = docEntryIsEmpty;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                oForm.Freeze(false);
                oForm.Update();
            }
        }

        public static void uiApp_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;
            
            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DRAW)
                {
                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                    {
                        try
                        {
                            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("51").Specific;
                            if (oMatrix.RowCount == 0)
                            {
                                FormsB1.SimulateRefresh();
                            }
                        }
                        catch
                        {

                        }
                    }
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD && pVal.BeforeAction)
                {
                    createFormItems(oForm, out errorText);
                    setVisibleFormItems(oForm);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_RESIZE && !pVal.BeforeAction)
                {
                    oForm.Freeze(true);
                    resizeForm(oForm, out errorText);
                    oForm.Freeze(false);
                }

                //დღგს განაკვეთის შეცვლისას თანხის გადათვლა
                if (pVal.ItemUID == "51" && pVal.ColUID == "BDOSVatGrp" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_COMBO_SELECT && !pVal.BeforeAction)
                {
                    oForm.Freeze(true);
                    ChangeVatPercent(oForm, pVal.Row, out errorText);
                    oForm.Freeze(false);
                }

                //დღგს თანხის შეცვლისას ჯამური თანხის შევსება
                if (pVal.ItemUID == "51" && pVal.ColUID == "BDOSVatAmt" && pVal.ItemChanged && !pVal.BeforeAction)
                {
                    oForm.Freeze(true);
                    RefillTotalVatAmounts(oForm, out errorText);
                    oForm.Freeze(false);
                }

                ////დანახარჯების თანხის შევსებისას დღგს თანხების გადათვლა
                //if (pVal.ItemUID == "54" && pVal.ItemChanged && !pVal.BeforeAction)
                //{
                //    oForm.Freeze(true);
                //    RefillVatAmounts(oForm, out errorText);
                //    oForm.Freeze(false);
                //}

                ////დანახარჯების თანხის შევსებისას დღგს თანხების გადათვლა
                //if (pVal.ItemUID == "24" && pVal.ItemChanged && !pVal.BeforeAction)
                //{
                //    oForm.Freeze(true);
                //    RefillVatAmounts(oForm, out errorText);
                //    oForm.Freeze(false);
                //}

                //რეკალკულაციის დროს დღგს განაკვეთების შევსება
                if ((pVal.ItemUID == "68" || pVal.ItemUID == "BDOSFllVAT") && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && !pVal.BeforeAction)
                {
                    oForm.Freeze(true);
                    FillTaxCodes(oForm, out errorText);
                    oForm.Freeze(false);
                }
                string uidd = pVal.ItemUID;
                //აწყობილი ანგარიშების მიხედვით ბუღ.თანხებით შევსება
                if (pVal.ItemUID == "BDOSFllAmt" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && !pVal.BeforeAction)
                {
                    if (oForm.Items.Item("68").Enabled)
                    {
                        oForm.Freeze(true);
                        FillCostsAmounts(oForm, out errorText);
                        oForm.Freeze(false);
                    }
                }

                else if (pVal.ItemUID == "53" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && !pVal.BeforeAction)
                {
                    setVisibleFormItems(oForm);
                }
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED & !pVal.BeforeAction)
                {
                    if (pVal.ItemUID == "BDOSStRev")
                    {
                        SAPbouiCOM.DBDataSource DocDBSource = oForm.DataSources.DBDataSources.Item(0);
                        string docNum = DocDBSource.GetValue("DocNum", 0);
                        string docEntry = "";
                        if (!stockExists(docNum))
                        {
                            StockRevaluation.fillStockRevaluation(docNum, out docEntry);
                            formDataLoad(oForm);                       
                            if (docEntry != "")
                            {
                                Program.uiApp.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_StockRevaluation, "162", docEntry);
                            }
                        }
                    }
                }
            }
        }

        public static void formDataLoad(SAPbouiCOM.Form oForm)
        {
            try
            {
                oForm.Freeze(true);

                SAPbouiCOM.DBDataSource DocDBSource = oForm.DataSources.DBDataSources.Item(0);
                string docNum = DocDBSource.GetValue("DocNum", 0);

                oForm.Items.Item("StockRevE").Specific.Value = StockRevaluation.getDocEntry(docNum);

                oForm.Update();
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

        public static bool stockExists(string docNum)
        {
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                string query = "select \"DocNum\" from OMRV " + "\n"
                + "where \"U_BaseDocNum\" = '" + docNum + "'";

                oRecordSet.DoQuery(query);

                if (!oRecordSet.EoF) return true;

                return false;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
            }
        }
        
    }
}