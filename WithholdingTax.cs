﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Globalization;
using System.Data;

namespace BDO_Localisation_AddOn
{
    static partial class WithholdingTax
    {
        public static CultureInfo cultureInfo = null;


        public static void JrnEntryAPInvoiceCredidNoteCheck(SAPbouiCOM.Form oForm, string DocType, out string errorText)
        {
            errorText = null;

            try
            {
                SAPbouiCOM.DBDataSource DocDBSource = DocType == "18" ? oForm.DataSources.DBDataSources.Item("PCH1") : oForm.DataSources.DBDataSources.Item("RPC1");

                if (DocDBSource.Size == 0)
                {
                    return;
                }

                for (int i = 0; i < DocDBSource.Size; i++)
                {
                    string VatGroup = DocDBSource.GetValue("VatGroup", i);

                    SAPbobsCOM.VatGroups oVG;
                    oVG = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVatGroups);
                    oVG.GetByKey(VatGroup);

                    string BDOSAccF = oVG.UserFields.Fields.Item("U_BDOSAccF").Value;
                    string TaxAccount = oVG.TaxAccount;

                    if (TaxAccount == "")
                    {
                        errorText = BDOSResources.getTranslate("CheckVatGroupAccounts");
                    }
                    if (BDOSAccF == "")
                    {
                        errorText = BDOSResources.getTranslate("CheckVatGroupAccounts");
                    }
                }


            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
        }

        public static void createUserFields(out string errorText)
        {
            errorText = null;

            Dictionary<string, object> fieldskeysMap;

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BdgtDbtAcc");
            fieldskeysMap.Add("TableName", "OWHT");
            fieldskeysMap.Add("Description", "Debt Account");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);


            Dictionary<string, string> listValidValuesDict = new Dictionary<string, string>();
            listValidValuesDict.Add("N", "N");
            listValidValuesDict.Add("Y", "Y");

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSOvStTx");
            fieldskeysMap.Add("TableName", "OWHT");
            fieldskeysMap.Add("Description", "Overstandart tax posting");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 1);
            fieldskeysMap.Add("DefaultValue", "N");
            fieldskeysMap.Add("ValidValues", listValidValuesDict);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSPhisTx");
            fieldskeysMap.Add("TableName", "OWHT");
            fieldskeysMap.Add("Description", "Physical Entity Tax");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 1);
            fieldskeysMap.Add("DefaultValue", "N");
            fieldskeysMap.Add("ValidValues", listValidValuesDict);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            GC.Collect();
        }

        public static void createFormItems(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("3").Specific;
            SAPbouiCOM.Columns oColumns = oMatrix.Columns;

            //////////

            SAPbouiCOM.Column oColumnCB = oMatrix.Columns.Item("U_BDOSOvStTx");
            oColumnCB.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;

            //////////

            SAPbouiCOM.Column oColumn = oMatrix.Columns.Item("U_BdgtDbtAcc");
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("DebtAccount");

            bool multiSelection = false;
            string objectType = "1";
            string uniqueID_lf_AccCode_CFL = "acc_CFL";
            //HR-შიც ემატება და შეცდომა რო არ გამოიწვიოს
            try
            {
                FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_lf_AccCode_CFL);
                oColumn.ChooseFromListUID = uniqueID_lf_AccCode_CFL;
                oColumn.ChooseFromListAlias = "AcctCode";

                SAPbouiCOM.ChooseFromList oCFL;
                SAPbouiCOM.Conditions oCons;
                SAPbouiCOM.Condition oCon;

                oCFL = oForm.ChooseFromLists.Item("acc_CFL");
                oCons = oCFL.GetConditions();
                oCon = oCons.Add();
                oCon.Alias = "Postable";
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCon.CondVal = "Y";
                oCFL.SetConditions(oCons);
            }
            catch
            { }

            oColumn = oMatrix.Columns.Item("U_BDOSPhisTx");
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("PhysicalEntityTax");
        }

        public static void chooseFromList(SAPbouiCOM.Form oForm, SAPbouiCOM.IChooseFromListEvent oCFLEvento, bool BeforeAction, int row, out string errorText)
        {
            errorText = null;
            try
            {
                string sCFL_ID = oCFLEvento.ChooseFromListUID;
                SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(sCFL_ID);

                SAPbouiCOM.DataTable oDataTable = null;
                oDataTable = oCFLEvento.SelectedObjects;

                if (BeforeAction == false)
                {
                    SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("3").Specific));

                    if (oDataTable != null)
                    {
                        if (sCFL_ID == "acc_CFL")
                        {
                            SAPbouiCOM.DataTable oDataTableSelectedObjects = oCFLEvento.SelectedObjects;
                            string val = oDataTableSelectedObjects.GetValue("AcctCode", 0);

                            SAPbouiCOM.EditText oEdit = oMatrix.Columns.Item("U_BdgtDbtAcc").Cells.Item(row).Specific;
                            oEdit.Value = val;
                        }


                    }
                }
            }
            catch
            {

            }
            finally
            {
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
            }
        }

        public static void uiApp_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD & pVal.BeforeAction == true)
                {
                    createFormItems(oForm, out errorText);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE & pVal.BeforeAction == false)
                {

                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                {
                    SAPbouiCOM.ChooseFromListEvent oCFLEvento = null;
                    oCFLEvento = ((SAPbouiCOM.ChooseFromListEvent)(pVal));
                    chooseFromList(oForm, oCFLEvento, pVal.BeforeAction, pVal.Row, out errorText);
                }
            }
        }

        public static DataTable getWtaxCodeDefinitionByDate(DateTime date, out string errorText)
        {
            errorText = null;

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string query = @"SELECT
	                         ""OWHT"".""U_BdgtDbtAcc"", 
                             ""WHT1"".""WTCode"",
                             ""WHT1"".""Rate"",
                             ""WHT1"".""EffecDate""

                        FROM (SELECT
	                         ""WTCode"",
	                         MAX(""EffecDate"") as ""EffecDate"" 
	                        FROM ( SELECT
	                         ""WTCode"", 
                             ""EffecDate""
		                        FROM ""WHT1"" 
		                        WHERE ""EffecDate""<='" + date.ToString("yyyyMMdd") + @"' ) AS ""SubTable""
	                        GROUP BY ""WTCode"") AS ""UNIQUE-WTDef-PAIRS"" 
                        INNER JOIN ""WHT1"" ON (""UNIQUE-WTDef-PAIRS"".""WTCode"" = ""WHT1"".""WTCode"") 
                        AND (""UNIQUE-WTDef-PAIRS"".""EffecDate"" = ""WHT1"".""EffecDate"" ) 
                          LEFT JOIN  ""OWHT"" ON ""OWHT"".""WTCode"" = ""WHT1"".""WTCode""";

            oRecordSet.DoQuery(query);

            DataTable WhTaxTable = new DataTable();

            WhTaxTable.Columns.Add("WTCode");
            WhTaxTable.Columns.Add("Rate");
            WhTaxTable.Columns.Add("BdgtDbtAcc");
            WhTaxTable.Columns.Add("EffecDate");

            while (!oRecordSet.EoF)
            {
                DataRow WhTaxRow = WhTaxTable.NewRow();
                WhTaxRow["WTCode"] = oRecordSet.Fields.Item("WTCode").Value;
                WhTaxRow["Rate"] = oRecordSet.Fields.Item("Rate").Value;
                WhTaxRow["BdgtDbtAcc"] = oRecordSet.Fields.Item("U_BdgtDbtAcc").Value;
                WhTaxRow["EffecDate"] = oRecordSet.Fields.Item("EffecDate").Value;

                WhTaxTable.Rows.Add(WhTaxRow);

                oRecordSet.MoveNext();
            }

            return WhTaxTable;
        }

        public static Dictionary<string, decimal> getPhysicalEntityPensionRates(DateTime DocDate, string BPWTCode, out string errorText)
        {
            errorText = "";

            Dictionary<string, decimal> PhysicalEntityPensionRates = new Dictionary<string, decimal>();
            PhysicalEntityPensionRates.Add("WTRate", 0);
            PhysicalEntityPensionRates.Add("PensionWTaxRate", 0);
            PhysicalEntityPensionRates.Add("PensionCoWTaxRate", 0);

            DataTable WTaxDefinitons = getWtaxCodeDefinitionByDate(DocDate, out errorText);
            string pensionWTCode = CommonFunctions.getOADM("U_BDOSPnPh").ToString();
            string pensionCoWTCode = CommonFunctions.getOADM("U_BDOSPnCoP").ToString();

            if (pensionWTCode == "")
            {
                errorText = BDOSResources.getTranslate("PhysicalEntityPension") + " " + BDOSResources.getTranslate("NotFilled");
            }

            if (pensionCoWTCode == "")
            {
                errorText = BDOSResources.getTranslate("CompanyPension") + " " + BDOSResources.getTranslate("NotFilled");
            }
            else
            {
                SAPbobsCOM.WithholdingTaxCodes oWHTaxCodeCo;
                oWHTaxCodeCo = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oWithholdingTaxCodes);
                oWHTaxCodeCo.GetByKey(pensionCoWTCode);

                if (oWHTaxCodeCo.Account == "")
                {
                    errorText = BDOSResources.getTranslate("CompanyPension") + " " + BDOSResources.getTranslate("Account") + " " + BDOSResources.getTranslate("NotFilled");
                }
            }

            if (BPWTCode != "")
            {
                string filter = "WTCode = '" + BPWTCode + "'";
                DataRow[] foundRows = WTaxDefinitons.Select(filter);
                if (foundRows.Count() > 0)
                {
                    PhysicalEntityPensionRates["WTRate"] = CommonFunctions.roundAmountByGeneralSettings(Convert.ToDecimal(foundRows[0]["Rate"]), "Rate");
                }
            }

            if (pensionWTCode != "")
            {
                string filter = "WTCode = '" + pensionWTCode + "'";
                DataRow[] foundRows = WTaxDefinitons.Select(filter);
                if (foundRows.Count() > 0)
                {
                    PhysicalEntityPensionRates["PensionWTaxRate"] = CommonFunctions.roundAmountByGeneralSettings(Convert.ToDecimal(foundRows[0]["Rate"]), "Rate");
                }
            }

            if (pensionCoWTCode != "")
            {
                string filter = "WTCode = '" + pensionCoWTCode + "'";
                DataRow[] foundRows = WTaxDefinitons.Select(filter);
                if (foundRows.Count() > 0)
                {
                    PhysicalEntityPensionRates["PensionCoWTaxRate"] = CommonFunctions.roundAmountByGeneralSettings(Convert.ToDecimal(foundRows[0]["Rate"]), "Rate");
                }
            }

            return PhysicalEntityPensionRates;
        }
    }
}