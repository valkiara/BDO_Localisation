using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BDO_Localisation_AddOn
{
    class Capitalization
    {
        public static void createUserFields(out string errorText)
        {
            errorText = null;

            Dictionary<string, object> fieldskeysMap;

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSDimCode");
            fieldskeysMap.Add("TableName", "ACQ1");
            fieldskeysMap.Add("Description", "Dimension");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);


            GC.Collect();
        }

        public static void createFormItems(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            Dictionary<string, object> formItems;
            string itemName = "";
            int left = oForm.Items.Item("1470000020").Left+5;
            int top = oForm.Items.Item("1470000020").Top - 50;

            formItems = new Dictionary<string, object>();
            itemName = "fillBdgFl";
            formItems.Add("Caption", BDOSResources.getTranslate("Fill"));
            formItems.Add("Size", 20);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            formItems.Add("Left", left);
            //formItems.Add("Width", 40);
            formItems.Add("Top", top);
            //formItems.Add("Height", height);
            formItems.Add("UID", itemName);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            bool multiSelection = false;
            string objectType = "62";
            string uniqueID_lf_Dist_CFL = "Dist_CFL";
            SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("1470000020").Specific;
            SAPbouiCOM.Column oColumn = oMatrix.Columns.Item("U_BDOSDimCode");
            FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_lf_Dist_CFL);
            oColumn.ChooseFromListUID = uniqueID_lf_Dist_CFL;
            oColumn.ChooseFromListAlias = "OcrCode";

        }

        public static void uiApp_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(BusinessObjectInfo.FormTypeEx, Program.currentFormCount);

            //შემოწმება
            if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD)
            {
                if (BusinessObjectInfo.ActionSuccess != BusinessObjectInfo.BeforeAction)
                {
                    //დოკუმენტის გატარების დროს გატარდეს ბუღლტრული გატარება
                    SAPbouiCOM.DBDataSource DocDBSourcePAYR = oForm.DataSources.DBDataSources.Item(0);
                    string DocEntry = DocDBSourcePAYR.GetValue("DocEntry", 0);

                    if (BusinessObjectInfo.ActionSuccess == true & BusinessObjectInfo.BeforeAction == false)
                    {
                        CommonFunctions.StartTransaction();

                        UpdateJournalEntry(DocEntry, out errorText);

                        CommonFunctions.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                    }

                }
            }



        }

        public static void UpdateJournalEntry(string DocEntry, out string errorText)
        {
            errorText = "";

            string accountStr = "";

            string ProfitCode = "";
            int Dim = Convert.ToInt32(CommonFunctions.getOADM("U_BDOSFADim"));

            if (Dim == 1)
            {
                ProfitCode = "CostingCode";
            }
            else if (Dim == 2)
            {
                ProfitCode = "CostingCode2";
            }
            else if (Dim == 3)
            {
                ProfitCode = "CostingCode3";
            }
            else if (Dim == 4)
            {
                ProfitCode = "CostingCode4";
            }
            else if (Dim == 5)
            {
                ProfitCode = "CostingCode5";
            }


            if (DocEntry != "")
            {
                SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string query = @"SELECT 
                            *  
                            FROM ""OJDT"" 
                            WHERE ""StornoToTr"" IS NULL   
                            AND ""TransType"" = '1470000049'  
                            AND ""CreatedBy"" = '" + DocEntry + "' ";

                oRecordSet.DoQuery(query);

                if (!oRecordSet.EoF)
                {
                    SAPbobsCOM.JournalEntries oJounalEntry = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
                    oJounalEntry.GetByKey(oRecordSet.Fields.Item("TransId").Value);

                    string TableQuery = @"select
                                     ""JDT1"".""Line_ID"",
                                     ""ACQ1"".""ClrAcqAct"",
	                                ""ACQ1"".""U_BDOSDimCode""
                                from ""JDT1""
                                inner join ""OJDT"" on ""JDT1"".""TransId"" = ""OJDT"".""TransId""
                                and ""OJDT"".""CreatedBy"" = "+DocEntry+ @"
                                and ""OJDT"".""StornoToTr"" IS NULL
                                and ""OJDT"".""TransType"" = '1470000049'
                                inner join(select

                                     ""OADT"".""ClrAcqAct"",
	                                ""ACQ1"".""U_BDOSDimCode""

                                    from ""ACQ1""

                                    inner join ""OITM"" on ""OITM"".""ItemCode"" = ""ACQ1"".""ItemCode""

                                    and ""ACQ1"".""DocEntry"" = " + DocEntry + @"
                                    inner join ""ACS1"" on ""ACS1"".""Code""=""OITM"".""AssetClass""
                                    inner join ""OADT"" on ""ACS1"".""AcctDtn"" = ""OADT"".""Code""

                                    group by ""OADT"".""ClrAcqAct"",
	                                ""ACQ1"".""U_BDOSDimCode"") as ""ACQ1"" on ""JDT1"".""Account"" = ""ACQ1"".""ClrAcqAct""";

                    oRecordSet.DoQuery(TableQuery);

                    while(!oRecordSet.EoF)
                    {
                        int i = oRecordSet.Fields.Item("Line_ID").Value;
                        string U_BDOSDimCode = oRecordSet.Fields.Item("U_BDOSDimCode").Value;
                        
                        oJounalEntry.Lines.SetCurrentLine(i);
                        
                            if (Dim == 1)
                            {
                                oJounalEntry.Lines.CostingCode = U_BDOSDimCode;
                            }
                            else if (Dim == 2)
                            {
                                oJounalEntry.Lines.CostingCode2 = U_BDOSDimCode;
                            }
                            else if (Dim == 3)
                            {
                                oJounalEntry.Lines.CostingCode3 = U_BDOSDimCode;
                            }
                            else if (Dim == 4)
                            {
                                oJounalEntry.Lines.CostingCode4 = U_BDOSDimCode;
                            }
                            else if (Dim == 5)
                            {
                                oJounalEntry.Lines.CostingCode5 = U_BDOSDimCode;
                            }

                        oRecordSet.MoveNext();
                    }

                    int updateCode = 0;
                    updateCode = oJounalEntry.Update();

                    if (updateCode != 0)
                    {
                        Program.oCompany.GetLastError(out updateCode, out errorText);
                    }
                }
            }
        }


        private static void FillFAAmounts(SAPbouiCOM.Form oForm)
        {
            string DateStr = oForm.Items.Item("1470000019").Specific.Value;
            SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("1470000020").Specific;
            string ProfitCode = "";
            int Dim = Convert.ToInt32(CommonFunctions.getOADM("U_BDOSFADim"));

            if (Dim == 1)
            {
                ProfitCode = "ProfitCode";
            }
            else if (Dim == 2)
            {
                ProfitCode = "OcrCode2";
            }
            else if (Dim == 3)
            {
                ProfitCode = "OcrCode3";
            }
            else if (Dim == 4)
            {
                ProfitCode = "OcrCode4";
            }
            else if (Dim == 5)
            {
                ProfitCode = "OcrCode5";
            }

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            for (int Row = 1; Row <= oMatrix.RowCount; Row++)
            {
                string ItemCode = oMatrix.GetCellSpecific("1470000003", Row).Value.ToString();
                if (ItemCode == "")
                {
                    continue;
                }
                SAPbobsCOM.Items oItems = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems);
                oItems.GetByKey(ItemCode);
                string AssetDet = FixedAsset.getAssetClassDetermination(oItems.AssetClass);
                string AcqAccount = CommonFunctions.GetAccountDetermination(AssetDet, "ClrAcqAct");

                string Query = @"select
	                         IFNULL(""OJDT"".""Debit"",0) - IFNULL(""OJDT"".""Credit"",0) as ""Balance""
                        from (select
	                         Sum(""Debit"") as ""Debit"",
	                         Sum(""Credit"") as ""Credit"" 
	                        from ""JDT1"" 
                        inner join ""OOCR"" on ""JDT1"".""" + ProfitCode + @""" = ""OOCR"".""OcrCode""
                        inner join ""OPRC"" on ""OPRC"".""PrcCode"" = ""OOCR"".""OcrCode"" and ""OPRC"".""U_BDOSFACode"" = '" + ItemCode + @"'
                            Where ""JDT1"".""TaxDate"" <= '" + DateStr + @"' and (""Account""='" + AcqAccount + @"' 
		                        or ""ContraAct""='" + AcqAccount + @"')) as ""OJDT""";

                if (Program.oCompany.DbServerType != SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                {
                    Query = Query.Replace("IFNULL", "ISNULL");
                }
                oRecordSet.DoQuery(Query);
                if (!oRecordSet.EoF)
                {
                    decimal Balance = Convert.ToDecimal(oRecordSet.Fields.Item("Balance").Value, CultureInfo.InvariantCulture);
                    if (Balance > 0)
                    {
                        oMatrix.GetCellSpecific("1470000009", Row).String = FormsB1.ConvertDecimalToStringForEditboxStrings(Balance);
                    }
                    oMatrix.GetCellSpecific("U_BDOSDimCode", Row).String = FixedAsset.getFADimension(ItemCode);
                }
            }
        }

 public static void chooseFromList(SAPbouiCOM.Form oForm, SAPbouiCOM.IChooseFromListEvent oCFLEvento, string itemUID, int RowIndex, bool beforeAction, out string errorText)
        {
            errorText = null;


            string sCFL_ID = oCFLEvento.ChooseFromListUID;

            SAPbouiCOM.DataTable oDataTable = null;
            oDataTable = oCFLEvento.SelectedObjects;

            if (oDataTable != null)
            {
                try
                {
                if (sCFL_ID == "Dist_CFL")
                {
                    SAPbouiCOM.DataTable oDataTableSelectedObjects = oCFLEvento.SelectedObjects;
                    string val = oDataTableSelectedObjects.GetValue("OcrCode", 0);
                    SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("1470000020").Specific;
                    oMatrix.GetCellSpecific("U_BDOSDimCode", RowIndex).Value = val;
                }
                }
                catch
                { }


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

                if (pVal.ItemUID == "fillBdgFl" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED & pVal.BeforeAction == false)
                {
                    FillFAAmounts(oForm);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                {
                    SAPbouiCOM.ChooseFromListEvent oCFLEvento = null;
                    oCFLEvento = ((SAPbouiCOM.ChooseFromListEvent)(pVal));
                
                    string sCFL_ID = oCFLEvento.ChooseFromListUID;

                    if (pVal.BeforeAction)
                    {
                        SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item("Dist_CFL");
                        SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();


                        SAPbouiCOM.Condition oCon = oCons.Add();
                        oCon.Alias = "DimCode";
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = CommonFunctions.getOADM("U_BDOSFADim").ToString();
                        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_NONE;

                        oCFL.SetConditions(oCons);
            }

                    else
                    {
                        chooseFromList(oForm, oCFLEvento, pVal.ItemUID, pVal.Row, pVal.BeforeAction, out errorText);
        }
    }
            }
        }
    }
}
