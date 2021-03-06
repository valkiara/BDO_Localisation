using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;

namespace BDO_Localisation_AddOn
{
    static partial class APInvoice
    {
        public static bool ProfitTaxTypeIsSharing = false;

        public static void createUserFields(out string errorText)
        {
            errorText = null;
            Dictionary<string, object> fieldskeysMap;

            //მოგების გადასახადი
            fieldskeysMap = new Dictionary<string, object>(); //ბეგრება განაწილებული მოგებით
            fieldskeysMap.Add("Name", "nonEconExp");
            fieldskeysMap.Add("TableName", "OPCH");
            fieldskeysMap.Add("Description", "Non-economic expenses");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 1);
            fieldskeysMap.Add("DefaultValue", "N");

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            //საშემოსავლო
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSWhtAmt");
            fieldskeysMap.Add("TableName", "PCH1");
            fieldskeysMap.Add("Description", "Withholding Tax");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            //დასაქმებულის საპენსიო
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSPnPhAm");
            fieldskeysMap.Add("TableName", "PCH1");
            fieldskeysMap.Add("Description", "Physical Entity Pens. Amount");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            //დამსაქმებლის საპენსიო
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSPnCoAm");
            fieldskeysMap.Add("TableName", "PCH1");
            fieldskeysMap.Add("Description", "Company Pens. Amount");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            //Use Blanket Agreement Rates
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "UseBlaAgRt");
            fieldskeysMap.Add("TableName", "OPCH");
            fieldskeysMap.Add("Description", "Use Blanket Agreement Rates");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 1);
            fieldskeysMap.Add("DefaultValue", "N");

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            //A/C Number
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSACNum");
            fieldskeysMap.Add("TableName", "OPCH");
            fieldskeysMap.Add("Description", "A/C Number");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            GC.Collect();
        }

        public static void createFormItems(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            ProfitTaxTypeIsSharing = ProfitTax.ProfitTaxTypeIsSharing();

            BDO_WBReceivedDocs.createFormItems(oForm, "OPCH", out errorText);

            Dictionary<string, object> formItems = null;

            string itemName = "";

            double height = oForm.Items.Item("86").Height;
            double top = oForm.Items.Item("86").Top + height * 1.5 + 1;
            double left_s = oForm.Items.Item("86").Left;
            double left_e = oForm.Items.Item("46").Left;
            double width_e = oForm.Items.Item("46").Width;

            //-------------------------------------------ანგარიშ-ფაქტურა----------------------------------->
            top = top + height * 1.5 + 1;

            formItems = new Dictionary<string, object>();
            itemName = "BDO_TaxTxt"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_e * 1.5);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("ChooseTaxInvoice"));
            formItems.Add("TextStyle", 4);
            formItems.Add("FontSize", 10);
            formItems.Add("Enabled", true);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            bool multiSelection = false;
            string objectType = "UDO_F_BDO_TAXR_D"; //Tax Invoice Received
            string uniqueID_TaxInvoiceReceivedCFL = "TaxInvoiceReceived_CFL";
            FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_TaxInvoiceReceivedCFL);

            formItems = new Dictionary<string, object>();
            itemName = "BDO_TaxDoc"; //10 characters
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
            formItems.Add("ChooseFromListUID", uniqueID_TaxInvoiceReceivedCFL);
            formItems.Add("ChooseFromListAlias", "DocEntry");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "BDO_TaxLB"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            formItems.Add("Left", left_e + width_e - 40 - 20);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("LinkTo", "BDO_TaxDoc");
            formItems.Add("LinkedObjectType", objectType);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "BDO_TaxCan"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            formItems.Add("Left", left_e + width_e);
            formItems.Add("Width", 20);
            formItems.Add("Top", top - 2);
            //formItems.Add("Height", height);
            formItems.Add("Image", "LINKMAP_ICON_CANCELLATION");
            formItems.Add("UID", itemName);
            formItems.Add("Description", BDOSResources.getTranslate("CancelLinkTaxInvoice"));
            formItems.Add("Visible", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 1;

            oForm.DataSources.UserDataSources.Add("BDO_TaxSer", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
            oForm.DataSources.UserDataSources.Add("BDO_TaxNum", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
            oForm.DataSources.UserDataSources.Add("BDO_TaxDat", SAPbouiCOM.BoDataType.dt_DATE, 20);
            //<-------------------------------------------ანგარიშ-ფაქტურა-----------------------------------

            //მოგების გადასახადი
            SAPbouiCOM.Item oItemS = oForm.Items.Item("230");
            SAPbouiCOM.Item oItemE = oForm.Items.Item("222");

            top = oItemS.Top + oItemS.Height;
            left_s = oItemS.Left;
            left_e = oItemE.Left;
            int width_s = oItemS.Width;
            width_e = oItemE.Width;

            formItems = new Dictionary<string, object>();
            itemName = "nonEconExp";
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OPCH");
            formItems.Add("Alias", "U_nonEconExp");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
            formItems.Add("Length", 1);
            formItems.Add("Left", left_s);
            formItems.Add("Width", 250);
            formItems.Add("Top", top);
            formItems.Add("Height", 14);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("NonEconomicExpenses"));
            formItems.Add("ValOff", "N");
            formItems.Add("ValOn", "Y");
            formItems.Add("DisplayDesc", true);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            SAPbouiCOM.Item oItem = oForm.Items.Item("90");
            top = oItem.Top;

            formItems = new Dictionary<string, object>();
            itemName = "PrBaseS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", 14);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("TaxableObject"));
            formItems.Add("LinkTo", "PrBaseE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            multiSelection = false;
            objectType = "UDO_F_BDO_PTBS_D";
            string uniqueID_lf_ProfitBaseCFL = "CFL_ProfitBase";
            FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_lf_ProfitBaseCFL);

            formItems = new Dictionary<string, object>();
            itemName = "PrBaseE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OPCH");
            formItems.Add("Alias", "U_prBase");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", 30);
            formItems.Add("Top", top);
            formItems.Add("Height", 14);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("ChooseFromListUID", uniqueID_lf_ProfitBaseCFL);
            formItems.Add("ChooseFromListAlias", "Code");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "PrBsDscr"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OPCH");
            formItems.Add("Alias", "U_PrBsDscr");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e + 32);
            formItems.Add("Width", width_e - 32);
            formItems.Add("Top", top);
            formItems.Add("Height", 14);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            //golden errow
            formItems = new Dictionary<string, object>();
            itemName = "PrBaseLB"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            formItems.Add("Left", left_e - 20);
            formItems.Add("Top", top);
            formItems.Add("Height", 14);
            formItems.Add("UID", itemName);
            formItems.Add("LinkTo", "PrBaseE");
            formItems.Add("LinkedObjectType", objectType);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            // -------------------- Use blanket agreement rates-----------------
            int pane = 7;
            int left = oForm.Items.Item("1720002167").Left;
            height = oForm.Items.Item("1720002167").Height;
            top = oForm.Items.Item("1720002167").Top + height + 5;

            formItems = new Dictionary<string, object>();
            itemName = "UsBlaAgRtS"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OPCH");
            formItems.Add("Alias", "U_UseBlaAgRt");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
            formItems.Add("Length", 1);
            formItems.Add("Left", left);
            formItems.Add("Width", 100);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("UseBlAgrRt"));
            formItems.Add("ValOff", "N");
            formItems.Add("ValOn", "Y");
            formItems.Add("DisplayDesc", true);
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("38").Specific;
            SAPbouiCOM.Column oColumn;

            oColumn = oMatrix.Columns.Item("U_BDOSWhtAmt");
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("WithholdingTax");

            oColumn = oMatrix.Columns.Item("U_BDOSPnPhAm");
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("PhysicalEntityPension");

            oColumn = oMatrix.Columns.Item("U_BDOSPnCoAm");
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("CompanyPension");

            oItem = oForm.Items.Item("70");
            top = oItem.Top + height * 2 + 1;
            left_s = oItem.Left;
            width_s = oItem.Width;
            oItem = oForm.Items.Item("4");
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
            formItems.Add("TableName", "OPCH");
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
        }

        public static void formDataLoad(SAPbouiCOM.Form oForm)
        {
            oForm.Freeze(true);
            try
            {
                setVisibleFormItems(oForm);

                var docEntry = oForm.DataSources.DBDataSources.Item("OPCH").GetValue("DocEntry", 0);
                string cardCode = oForm.DataSources.DBDataSources.Item("OPCH").GetValue("CardCode", 0).Trim();
                //-------------------------------------------ანგარიშ-ფაქტურა----------------------------------->
                string caption = BDOSResources.getTranslate("ChooseTaxInvoice");
                int taxDocEntry = 0;
                string taxID = "";
                string taxNumber = "";
                string taxSeries = "";
                string taxStatus = "";
                string taxCreateDate = "";

                if (!string.IsNullOrEmpty(docEntry))
                {
                    Dictionary<string, object> taxDocInfo = BDO_TaxInvoiceReceived.getTaxInvoiceReceivedDocumentInfo(Convert.ToInt32(docEntry), BDO_TaxInvoiceReceived.BaseDocType.oPurchaseInvoices, cardCode);
                    if (taxDocInfo != null)
                    {
                        taxDocEntry = Convert.ToInt32(taxDocInfo["docEntry"]);
                        taxID = taxDocInfo["invID"].ToString();
                        taxNumber = taxDocInfo["number"].ToString();
                        taxSeries = taxDocInfo["series"].ToString();
                        taxStatus = taxDocInfo["status"].ToString();
                        taxCreateDate = taxDocInfo["createDate"].ToString();

                        if (taxDocEntry != 0)
                        {
                            DateTime taxCreateDateDT = DateTime.ParseExact(taxCreateDate, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
                            if (taxSeries == "")
                            {
                                caption = BDOSResources.getTranslate("TaxInvoiceDate") + " " + taxCreateDateDT;
                            }
                            else
                            {
                                caption = BDOSResources.getTranslate("TaxInvoiceSeries") + " " + taxSeries + " № " + taxNumber + " " + BDOSResources.getTranslate("Data") + " " + taxCreateDateDT;
                            }
                        }
                    }
                }
                else
                {
                    taxDocEntry = 0;
                }

                oForm.DataSources.UserDataSources.Item("BDO_TaxDoc").ValueEx = taxDocEntry == 0 ? "" : taxDocEntry.ToString();
                oForm.DataSources.UserDataSources.Item("BDO_TaxSer").ValueEx = taxSeries;
                oForm.DataSources.UserDataSources.Item("BDO_TaxNum").ValueEx = taxNumber;
                oForm.DataSources.UserDataSources.Item("BDO_TaxDat").ValueEx = taxCreateDate;

                SAPbouiCOM.StaticText oStaticText = (SAPbouiCOM.StaticText)oForm.Items.Item("BDO_TaxTxt").Specific;
                oStaticText.Caption = caption;
                //<-------------------------------------------ანგარიშ-ფაქტურა-----------------------------------
            }
            catch (Exception ex)
            {
                oForm.DataSources.UserDataSources.Item("BDO_TaxDoc").ValueEx = "";
                oForm.DataSources.UserDataSources.Item("BDO_TaxSer").ValueEx = "";
                oForm.DataSources.UserDataSources.Item("BDO_TaxNum").ValueEx = "";
                oForm.DataSources.UserDataSources.Item("BDO_TaxDat").ValueEx = "";

                SAPbouiCOM.StaticText oStaticText = (SAPbouiCOM.StaticText)oForm.Items.Item("BDO_TaxTxt").Specific;
                oStaticText.Caption = BDOSResources.getTranslate("ChooseTaxInvoice");
                oForm.Items.Item("BDO_TaxCan").Visible = false;
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        public static void getAmount(int docEntry, out double gTotal, out double lineVat, out string wblNumber)
        {
            gTotal = 0.0;
            lineVat = 0.0;
            wblNumber = string.Empty;

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string query = @"SELECT 
            ""PCH1"".""DocEntry"" AS ""docEntry"",
            ""OPCH"".""U_BDO_WBNo"" AS ""WblNumber"",
            SUM(""PCH1"".""GTotal"") AS ""GTotal"",
            SUM(""PCH1"".""LineVat"") AS ""LineVat""
            FROM ""PCH1"" AS ""PCH1""
            INNER JOIN ""OPCH"" ON ""PCH1"".""DocEntry"" = ""OPCH"".""DocEntry""
            WHERE ""PCH1"".""DocEntry"" = '" + docEntry + @"'
            GROUP BY ""PCH1"".""DocEntry"", 
                     ""OPCH"".""U_BDO_WBNo"" ";
            try
            {
                oRecordSet.DoQuery(query);

                if (!oRecordSet.EoF)
                {
                    gTotal = oRecordSet.Fields.Item("GTotal").Value;
                    lineVat = oRecordSet.Fields.Item("LineVat").Value;
                    wblNumber = oRecordSet.Fields.Item("WblNumber").Value;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
            }
        }

        public static void chooseFromList(SAPbouiCOM.Form oForm, SAPbouiCOM.IChooseFromListEvent oCFLEvento, string itemUID, bool beforeAction)
        {
            try
            {
                string sCFL_ID = oCFLEvento.ChooseFromListUID;
                SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(sCFL_ID);

                if (beforeAction)
                {
                    if (sCFL_ID == "TaxInvoiceReceived_CFL")
                    {
                        string wbNumber = oForm.DataSources.DBDataSources.Item("OPCH").GetValue("U_BDO_WBNo", 0).Trim();
                        string cardCode = oForm.DataSources.DBDataSources.Item("OPCH").GetValue("CardCode", 0).Trim();
                        DateTime docDate = DateTime.ParseExact(oForm.DataSources.DBDataSources.Item("OPCH").GetValue("DocDate", 0), "yyyyMMdd", CultureInfo.InvariantCulture);

                        List<string> taxInvoiceDocList = BDO_TaxInvoiceReceived.getListTaxInvoiceReceived(cardCode, wbNumber, BDO_TaxInvoiceReceived.BaseDocType.oPurchaseInvoices, docDate);

                        int docCount = taxInvoiceDocList.Count;
                        SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();
                        if (docCount == 0)
                        {
                            SAPbouiCOM.Condition oCon = oCons.Add();
                            oCon.Alias = "DocEntry";
                            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                            oCon.CondVal = "";
                        }
                        else
                        {
                            for (int i = 0; i < docCount; i++)
                            {
                                SAPbouiCOM.Condition oCon = oCons.Add();
                                oCon.Alias = "DocEntry";
                                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                                oCon.CondVal = taxInvoiceDocList[i];
                                oCon.Relationship = (i == docCount - 1) ? SAPbouiCOM.BoConditionRelationship.cr_NONE : SAPbouiCOM.BoConditionRelationship.cr_OR;
                            }
                        }
                        oCFL.SetConditions(oCons);
                    }
                }
                else
                {
                    SAPbouiCOM.DataTable oDataTable = oCFLEvento.SelectedObjects;

                    if (oDataTable != null)
                    {
                        if (sCFL_ID == "TaxInvoiceReceived_CFL")
                        {
                            string taxDocEntryStr = oDataTable.GetValue("DocEntry", 0).ToString();
                            BDO_TaxInvoiceReceived.chooseFromListForBaseDocs(oForm, taxDocEntryStr);
                        }

                        if (sCFL_ID == "CFL_ProfitBase")
                        {
                            string ProfitBaseCode = Convert.ToString(oDataTable.GetValue("Code", 0));
                            string ProfitBaseName = Convert.ToString(oDataTable.GetValue("Name", 0));

                            try
                            {
                                SAPbouiCOM.EditText oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("PrBaseE").Specific;
                                oEditText.Value = ProfitBaseCode;
                            }
                            catch { }

                            try
                            {
                                SAPbouiCOM.EditText oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("PrBsDscr").Specific;
                                oEditText.Value = ProfitBaseName;
                            }
                            catch { }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static void setVisibleFormItems(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.Item oItem;
            oForm.Freeze(true);

            try
            {

                string NonEconExp = oForm.DataSources.DBDataSources.Item("OPCH").GetValue("U_nonEconExp", 0);
                string docEntrySTR = oForm.DataSources.DBDataSources.Item("OPCH").GetValue("DocEntry", 0).Trim();
                bool docEntryIsEmpty = string.IsNullOrEmpty(docEntrySTR);

                oForm.Items.Item("16").Click(); //focus on remark item

                if (ProfitTaxTypeIsSharing)
                {
                    oForm.Items.Item("nonEconExp").Enabled = (docEntryIsEmpty);
                    oForm.Items.Item("PrBaseE").Enabled = (docEntryIsEmpty && NonEconExp == "Y");

                    string uniqueID_lf_ProfitBaseCFL = "CFL_ProfitBase";
                    oForm.Items.Item("PrBaseE").Specific.ChooseFromListUID = uniqueID_lf_ProfitBaseCFL;
                    oForm.Items.Item("PrBaseE").Specific.ChooseFromListAlias = "Code";
                }
                else
                {
                    oForm.Items.Item("nonEconExp").Visible = false;
                    oForm.Items.Item("PrBaseS").Visible = false;
                    oForm.Items.Item("PrBaseE").Visible = false;
                    oForm.Items.Item("PrBsDscr").Visible = false;
                    oForm.Items.Item("PrBaseLB").Visible = false;
                }

                if (oForm.DataSources.UserDataSources.Item("BDO_TaxDoc").ValueEx != "")
                {
                    oItem = oForm.Items.Item("BDO_TaxCan");
                    oItem.Visible = true;
                    oItem = oForm.Items.Item("BDO_TaxDoc");
                    oItem.Enabled = false;
                }
                else if (string.IsNullOrEmpty(docEntrySTR))
                {
                    oItem = oForm.Items.Item("BDO_TaxDoc");
                    oItem.Enabled = false;
                }
                else
                {
                    oItem = oForm.Items.Item("BDO_TaxCan");
                    oItem.Visible = false;
                    oItem = oForm.Items.Item("BDO_TaxDoc");
                    oItem.Enabled = true;

                    SAPbouiCOM.EditText oEditText = (SAPbouiCOM.EditText)oItem.Specific;
                    oEditText.ChooseFromListUID = "TaxInvoiceReceived_CFL";
                    oEditText.ChooseFromListAlias = "DocEntry";
                }
            }
            catch
            {

            }
            finally
            {
                oForm.Freeze(false);
                oForm.Update();
            }

            FormsB1.WB_TAX_AuthorizationsItems(oForm);
        }

        public static List<int> getDocListAPCreditMemo(string docEntryAPInvoice)
        {
            List<int> docListAPCreditMemo = new List<int>();
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string query = "SELECT DISTINCT " +
            "\"RPC1\".\"DocEntry\" AS \"DocEntry\", " +
            "\"RPC1\".\"BaseType\" AS \"BaseType\", " +
            "\"RPC1\".\"BaseEntry\" AS \"BaseEntry\" " +
            "FROM \"RPC1\" " +
            "WHERE \"RPC1\".\"BaseEntry\" IN (" + docEntryAPInvoice + ") AND \"RPC1\".\"BaseType\" = '18'";

            oRecordSet.DoQuery(query);
            while (!oRecordSet.EoF)
            {
                docListAPCreditMemo.Add(oRecordSet.Fields.Item("DocEntry").Value);
                oRecordSet.MoveNext();
            }
            return docListAPCreditMemo;
        }

        public static void formDataAddUpdate(SAPbouiCOM.Form oForm)
        {
            try
            {
                string taxDocEntry = oForm.DataSources.UserDataSources.Item("BDO_TaxDoc").ValueEx;
                string docEntryStr = oForm.DataSources.DBDataSources.Item("OPCH").GetValue("DocEntry", 0);
                if (!string.IsNullOrEmpty(taxDocEntry) && !string.IsNullOrEmpty(docEntryStr))
                {
                    int docEntry = Convert.ToInt32(docEntryStr);
                    getAmount(docEntry, out var baseDocGTotal, out var baseDocLineVat, out var baseDocWblNmber);
                    BDO_TaxInvoiceReceived.addBaseDoc(Convert.ToInt32(taxDocEntry), docEntry, BDO_TaxInvoiceReceived.BaseDocType.oPurchaseInvoices, baseDocWblNmber, baseDocGTotal, baseDocLineVat);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static void uiApp_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText;

            SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(BusinessObjectInfo.FormTypeEx, Program.currentFormCount);

            if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD)
            {
                if (BusinessObjectInfo.BeforeAction)
                {
                    SAPbouiCOM.DBDataSource DocDBSource = oForm.DataSources.DBDataSources.Item(0);
                    if (DocDBSource.GetValue("CANCELED", 0) == "N")
                    {
                        // მოგების გადასახადი
                        if (ProfitTaxTypeIsSharing)
                        {
                            if (oForm.Items.Item("204").Specific.Value != "")
                            {
                                bool TaxAccountsIsEmpty = ProfitTax.TaxAccountsIsEmpty();
                                if (TaxAccountsIsEmpty)
                                {
                                    Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("TaxAccounts") + " " + BDOSResources.getTranslate("YouCantLeaveEmpty"));
                                    Program.uiApp.MessageBox(BDOSResources.getTranslate("OperationUnsuccesfullSeeLog"));
                                    BubbleEvent = false;
                                }
                            }
                            if (oForm.DataSources.DBDataSources.Item("OPCH").GetValue("U_nonEconExp", 0) == "Y")
                            {
                                if (oForm.DataSources.DBDataSources.Item("OPCH").GetValue("U_prBase", 0) == "")
                                {
                                    Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("TaxableObject") + " " + BDOSResources.getTranslate("YouCantLeaveEmpty"));
                                    Program.uiApp.MessageBox(BDOSResources.getTranslate("OperationUnsuccesfullSeeLog"));
                                    BubbleEvent = false;
                                }
                            }
                        }
                        //დღგს გატარების შემოწმება
                        //string VatStatus = oForm.DataSources.DBDataSources.Item("OCRD").GetValue("VatStatus", 0);

                        //if (VatStatus.Trim() == "E")
                        //{
                        //    WithholdingTax.JrnEntryAPInvoiceCredidNoteCheck( oForm, BusinessObjectInfo.Type, out errorText);

                        //    if (errorText != null)
                        //    {
                        //        Program.uiApp.MessageBox(errorText);
                        //        BubbleEvent = false;
                        //        return;
                        //    }
                        //}

                        //დღგს თარიღის შევსება
                        oForm.Freeze(true);
                        int panelLevel = oForm.PaneLevel;
                        string sdocDate = oForm.Items.Item("10").Specific.Value;
                        oForm.PaneLevel = 7;
                        oForm.Items.Item("1000").Specific.Value = sdocDate;
                        oForm.PaneLevel = panelLevel;
                        oForm.Freeze(false);

                        if (oForm.DataSources.DBDataSources.Item("PCH1").GetValue("BaseType", 0) == "20")
                        {
                            //SAPbouiCOM.EditText oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("BDO_WBNo").Specific;
                            oForm.Items.Item("BDO_WBNo").Specific.Value = "1";
                            oForm.Items.Item("BDO_WBNo").Specific.Value = " ";

                            //oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("BDO_WBID").Specific;
                            oForm.Items.Item("BDO_WBID").Specific.Value = "1";
                            oForm.Items.Item("BDO_WBID").Specific.Value = " ";

                            //oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("actDate").Specific;
                            oForm.Items.Item("actDate").Specific.Value = "00010101";

                            //SAPbouiCOM.ComboBox oCombobox = (SAPbouiCOM.ComboBox)oForm.Items.Item("BDO_WBSt").Specific;
                            oForm.Items.Item("BDO_WBSt").Specific.Select("-1");

                            //SAPbouiCOM.CheckBox oCheckBox = (SAPbouiCOM.CheckBox)oForm.Items.Item("WBrec").Specific;
                            oForm.Items.Item("WBrec").Specific.Checked = false;
                        }
                        else
                        {
                            //დოკუმენტი არ დაემატოს ზედნადების გარეშე, თუ მომწოდებელს ჩართული აქვს
                            string CardCode = DocDBSource.GetValue("CardCode", 0);

                            SAPbobsCOM.BusinessPartners oBP = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);
                            oBP.GetByKey(CardCode);

                            string RSControlType = oBP.UserFields.Fields.Item("U_BDO_MapCnt").Value;
                            string NeedWB = oBP.UserFields.Fields.Item("U_BDO_NeedWB").Value;
                            RSControlType = RSControlType.Trim();
                            NeedWB = NeedWB.Trim();

                            string DocType = DocDBSource.GetValue("DocType", 0);

                            if ((RSControlType == "2" || RSControlType == "3") && (DocType == "I"))
                            {
                                SAPbouiCOM.EditText oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("BDO_WBID").Specific;
                                string WBID = oEditText.Value;

                                if (WBID == "" && NeedWB == "Y")
                                {
                                    bool isStock = false;

                                    SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("38").Specific;

                                    for (int row = 1; row <= oMatrix.RowCount; row++)
                                    {
                                        // SAPbouiCOM.EditText Edtfieldtxt = oMatrix.Columns.Item("ItemCode").Cells.Item(row).Specific;
                                        string formItemCode = oMatrix.GetCellSpecific("1", row).Value;

                                        if (Items.isStockItem(formItemCode))
                                        {
                                            isStock = true;
                                            break;
                                        }
                                    }

                                    if (isStock)
                                    {
                                        Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("BPControledOnRSLinkWaybillDocument"));
                                        BubbleEvent = !(BubbleEvent);
                                    }
                                }
                                else
                                {
                                    string Doctype = "";

                                    if (BusinessObjectInfo.Type == "18")
                                    {
                                        Doctype = "APInvoice";
                                    }
                                    else if (BusinessObjectInfo.Type == "19")
                                    {
                                        Doctype = "CredMemo";
                                    }
                                    try
                                    {
                                        bool continuePosting = BDO_WBReceivedDocs.waybillsCompare(WBID, oForm, RSControlType, Doctype, out errorText);

                                        if (continuePosting == false)
                                        {
                                            Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("GoodsTableNotMatchedESTable"));
                                            Program.uiApp.MessageBox(BDOSResources.getTranslate("OperationUnsuccesfullSeeLog"));
                                            BubbleEvent = !(BubbleEvent);
                                        }
                                    }
                                    catch { }
                                }
                            }
                        }
                    }
                }

                if (BusinessObjectInfo.ActionSuccess != BusinessObjectInfo.BeforeAction)
                {
                    //დოკუმენტის გატარების დროს გატარდეს ბუღლტრული გატარება
                    SAPbouiCOM.DBDataSource DocDBSource = oForm.DataSources.DBDataSources.Item(0);

                    if (DocDBSource.GetValue("CANCELED", 0) == "N")
                    {
                        string DocEntry = DocDBSource.GetValue("DocEntry", 0);
                        string DocCurrency = DocDBSource.GetValue("DocCur", 0);
                        decimal DocRate = FormsB1.cleanStringOfNonDigits(DocDBSource.GetValue("DocRate", 0));
                        string DocNum = DocDBSource.GetValue("DocNum", 0);
                        DateTime DocDate = DateTime.ParseExact(DocDBSource.GetValue("DocDate", 0), "yyyyMMdd", CultureInfo.InvariantCulture);

                        CommonFunctions.StartTransaction();

                        Program.JrnLinesGlobal = new DataTable();
                        DataTable JrnLinesDT = createAdditionalEntries(oForm, null, null, DocCurrency, out var reLines, DocRate);

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

                        if (Program.oCompany.InTransaction)
                        {
                            //თუ დოკუმენტი გატარდა, მერე ვაკეთებს ბუღალტრულ გატარებას
                            if (BusinessObjectInfo.ActionSuccess && BusinessObjectInfo.BeforeAction == false)
                            {
                                CommonFunctions.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                            }
                            else
                            {
                                CommonFunctions.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                            }
                        }
                        else
                        {
                            Program.uiApp.MessageBox("ტრანზაქციის დასრულებს შეცდომა");
                            BubbleEvent = false;
                        }
                    }
                }
            }

            //A/C Number and Use Rate Ranges Update
            if ((BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD || BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE)
                            && BusinessObjectInfo.ActionSuccess && BusinessObjectInfo.BeforeAction == false)
            {
                CommonFunctions.StartTransaction();

                SAPbouiCOM.DBDataSource DocDBSource = oForm.DataSources.DBDataSources.Item(0);
                string DocEntry = DocDBSource.GetValue("DocEntry", 0);
                string ObjType = DocDBSource.GetValue("ObjType", 0);
                string ACNumber = DocDBSource.GetValue("U_BDOSACNum", 0);

                string UseRateRanges = DocDBSource.GetValue("U_UseBlaAgRt", 0);

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

            if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE)
            {
                if (BusinessObjectInfo.BeforeAction)
                {
                    try
                    {
                        formDataAddUpdate(oForm);
                    }
                    catch (Exception ex)
                    {
                        Program.uiApp.MessageBox(ex.Message);
                        BubbleEvent = false;
                    }
                    CheckIfDepreciationIsAccrued(oForm, out string deprError);
                    if (!string.IsNullOrEmpty(deprError))
                    {
                        Program.uiApp.MessageBox(deprError);
                        BubbleEvent = false;
                    }
                }
                else if (BusinessObjectInfo.BeforeAction == false && BusinessObjectInfo.ActionSuccess)
                {
                    setVisibleFormItems(oForm);
                }
            }

            if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD && !BusinessObjectInfo.BeforeAction)
            {
                //when "Keep Visible" is not selected Program.uiApp.Forms.ActiveForm.Type = 10017, so we need check
                if (Program.uiApp.Forms.ActiveForm.Type == 141) // Keep Visible Case
                    oForm = Program.uiApp.Forms.ActiveForm;

                formDataLoad(oForm);
                setVisibleFormItems(oForm);
                BDO_WBReceivedDocs.setwaybillText(oForm);
            }

            if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD && BusinessObjectInfo.BeforeAction == false && BusinessObjectInfo.ActionSuccess)
            {
                if (Program.canceledDocEntry != 0)
                {
                    cancellation(oForm, Program.canceledDocEntry);
                    Program.canceledDocEntry = 0;
                }
            }
        }

        public static void uiApp_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText;

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (pVal.ItemUID == "PrBaseE" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                {
                    if (!pVal.BeforeAction)
                    {
                        SAPbouiCOM.ChooseFromListEvent oCFLEvento = (SAPbouiCOM.ChooseFromListEvent)pVal;
                        chooseFromList(oForm, oCFLEvento, pVal.ItemUID, pVal.BeforeAction);
                    }
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
                {
                    if (pVal.ItemUID == "1" && pVal.BeforeAction)
                    {
                        CommonFunctions.fillDocRate(oForm, "OPCH");
                    }

                    if (pVal.ItemUID == "UsBlaAgRtS")
                    {
                        SAPbouiCOM.EditText oBlankAgr = (SAPbouiCOM.EditText)oForm.Items.Item("1980002192").Specific;

                        if (string.IsNullOrEmpty(oBlankAgr.Value))
                        {
                            Program.uiApp.SetStatusBarMessage(errorText = BDOSResources.getTranslate("EmptyBlaAgrError"), SAPbouiCOM.BoMessageTime.bmt_Short);
                            SAPbouiCOM.CheckBox oUsBlaAgRtCB = (SAPbouiCOM.CheckBox)oForm.Items.Item("UsBlaAgRtS").Specific;
                            oUsBlaAgRtCB.Checked = false;
                            oForm.Items.Item("1980002192").Click();
                        }
                    }

                    if (pVal.ItemUID == "nonEconExp" && !pVal.BeforeAction)
                    {
                        nonEconExp_OnClick(oForm);
                    }
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_RESIZE && !pVal.BeforeAction)
                {
                    resizeForm(oForm);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD && pVal.BeforeAction)
                {
                    createFormItems(oForm, out errorText);
                    formDataLoad(oForm);
                    setVisibleFormItems(oForm);

                    Program.currentFormCount = pVal.FormTypeCount;
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_COMBO_SELECT && pVal.ItemUID == "WBOper" && !pVal.BeforeAction)
                {
                    Program.oIncWaybDocFormAPInv = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                    oForm.Freeze(true);
                    BDO_WBReceivedDocs.comboSelect(oForm, Program.oIncWaybDocFormAPInv, "Invoice", out errorText);
                    oForm.Freeze(false);
                }

                if (pVal.ItemUID == "BDO_TaxDoc" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST) // || pVal.ItemUID == "4")
                {
                    SAPbouiCOM.IChooseFromListEvent oCFLEvento = (SAPbouiCOM.IChooseFromListEvent)pVal;
                    chooseFromList(oForm, oCFLEvento, pVal.ItemUID, pVal.BeforeAction);
                }

                if (pVal.ItemUID == "16" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_GOT_FOCUS)
                {
                    setVisibleFormItems(oForm);
                }

                if (pVal.ItemUID == "BDO_TaxDoc" && !pVal.BeforeAction && pVal.EventType == SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK)
                {
                    formDataLoad(oForm);
                }

                if (pVal.ItemUID == "BDO_TaxCan" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.BeforeAction == false)
                {
                    FormsB1.WB_TAX_AuthorizationsOperations("UDO_FT_UDO_F_BDO_TAXS_D", SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    int taxDocEntry = Convert.ToInt32(oForm.DataSources.UserDataSources.Item("BDO_TaxDoc").ValueEx);
                    int docEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("OPCH").GetValue("DocEntry", 0));
                    if (taxDocEntry != 0)
                    {
                        int answer = Program.uiApp.MessageBox(BDOSResources.getTranslate("DocumentLinkedToTaxInvoiceCancel"), 1, BDOSResources.getTranslate("Yes"), BDOSResources.getTranslate("No"), "");
                        if (answer == 1)
                        {
                            BDO_TaxInvoiceReceived.removeBaseDoc(taxDocEntry, docEntry, BDO_TaxInvoiceReceived.BaseDocType.oPurchaseInvoices);
                            formDataLoad(oForm);
                            setVisibleFormItems(oForm);
                        }
                    }
                }
            }
        }

        public static void uiApp_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent, out string errorText)
        {
            errorText = null;
            BubbleEvent = true;

            //----------------------------->Cancel <-----------------------------
            SAPbouiCOM.Form oForm = Program.uiApp.Forms.ActiveForm;
            if (pVal.BeforeAction && pVal.MenuUID == "1284")
            {
                CheckIfDepreciationIsAccrued(oForm, out string deprError);
                if (!string.IsNullOrEmpty(deprError))
                {
                    Program.uiApp.MessageBox(deprError);
                    BubbleEvent = false;
                }
            }
        }

        public static void cancellation(SAPbouiCOM.Form oForm, int docEntry)
        {
            try
            {
                JournalEntry.cancellation(oForm, docEntry, "18", out var errorText);
                if (!string.IsNullOrEmpty(errorText))
                {
                    throw new Exception(errorText);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static DataTable createAdditionalEntries(SAPbouiCOM.Form oForm, SAPbobsCOM.GeneralData oGeneralData, DataTable DTSource, string DocCurrency, out DataTable reLines, decimal DocRate)
        {
            DataTable jeLines = JournalEntry.JournalEntryTable();
            DocCurrency = DocCurrency == CommonFunctions.getLocalCurrency() ? "" : DocCurrency;
            DataTable AccountTable = CommonFunctions.GetOACTTable();
            reLines = ProfitTax.ProfitTaxTable();
            DataRow reLinesRow = null;

            SAPbouiCOM.DBDataSources docDBSources = oForm.DataSources.DBDataSources;
            SAPbouiCOM.DBDataSource DBDataSourceTable = null;
            int JEcount = 0;

            if (oForm == null)
            {
                JEcount = DTSource.Rows.Count;
            }
            else
            {
                DBDataSourceTable = docDBSources.Item("PCH11");
                JEcount = DBDataSourceTable.Size;
            }

            SAPbouiCOM.DBDataSource BPDataSourceTable = docDBSources.Item("OCRD");

            string CardCode = BPDataSourceTable.GetValue("CardCode", 0).Trim();
            string vatCode = BPDataSourceTable.GetValue("ECVatGroup", 0).Trim();
            string TaxType = BPDataSourceTable.GetValue("U_BDO_TaxTyp", 0).Trim();
            string U_BDOSPnAcc = CommonFunctions.getOADM("U_BDOSPnAcc").ToString();
            //მოგების გადასახადის გატარება
            ProfitTaxTypeIsSharing = ProfitTax.ProfitTaxTypeIsSharing();

            if (ProfitTaxTypeIsSharing)
            {
                string U_nonEconExp = docDBSources.Item("OPCH").GetValue("U_nonEconExp", 0).Trim();
                decimal DpmAmnt = FormsB1.cleanStringOfNonDigits(docDBSources.Item("OPCH").GetValue("DpmAmnt", 0));


                decimal U_BDO_PrTxRt = Convert.ToDecimal(CommonFunctions.getOADM("U_BDO_PrTxRt").ToString());

                if (U_nonEconExp == "Y" && DpmAmnt > 0)
                {
                    string DebitAccount = CommonFunctions.getOADM("U_BDO_CapAcc").ToString();
                    string CreditAccount = CommonFunctions.getOADM("U_BDO_TaxAcc").ToString();

                    string prBase = docDBSources.Item("OPCH").GetValue("U_prBase", 0).Trim();
                    decimal TaxAmount = DpmAmnt * U_BDO_PrTxRt / (100 - U_BDO_PrTxRt);
                    decimal TaxAmountFC = DocCurrency == "" ? 0 : TaxAmount / DocRate;

                    JournalEntry.AddJournalEntryRow(AccountTable, jeLines, "Full", DebitAccount, CreditAccount, TaxAmount, TaxAmountFC, DocCurrency,
                                                    "", "", "", "", "", "", "", "");

                    reLinesRow = reLines.Rows.Add();

                    reLinesRow["debitAccount"] = DebitAccount;
                    reLinesRow["creditAccount"] = CreditAccount;
                    reLinesRow["prBase"] = prBase;
                    reLinesRow["txType"] = "Accrual";
                    reLinesRow["amtTx"] = DpmAmnt;
                    reLinesRow["amtPrTx"] = TaxAmount;

                }

                for (int i = 0; i < JEcount; i++)
                {
                    string BaseAbs = CommonFunctions.getChildOrDbDataSourceValue(DBDataSourceTable, null, DTSource, "BaseAbs", i).ToString();
                    string BaseType = CommonFunctions.getChildOrDbDataSourceValue(DBDataSourceTable, null, DTSource, "BaseType", i).ToString();
                    decimal BaseGross = FormsB1.cleanStringOfNonDigits(CommonFunctions.getChildOrDbDataSourceValue(DBDataSourceTable, null, DTSource, "BaseGross", i).ToString());
                    string CreditAccount = CommonFunctions.getOADM("U_BDO_CapAcc").ToString();
                    string DebitAccount = CommonFunctions.getOADM("U_BDO_TaxAcc").ToString();

                    if (BaseType == "204" && BaseGross > 0)
                    {
                        SAPbobsCOM.Documents oInvoice;
                        oInvoice = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDownPayments);
                        oInvoice.GetByKey(Convert.ToInt32(BaseAbs));
                        string U_liablePrTx = oInvoice.UserFields.Fields.Item("U_liablePrTx").Value;
                        string prBase = oInvoice.UserFields.Fields.Item("U_prBase").Value;

                        if (U_liablePrTx == "Y")
                        {
                            decimal TaxAmount = BaseGross * U_BDO_PrTxRt / (100 - U_BDO_PrTxRt);
                            decimal TaxAmountFC = DocCurrency == "" ? 0 : TaxAmount / DocRate;

                            JournalEntry.AddJournalEntryRow(AccountTable, jeLines, "Full", DebitAccount, CreditAccount, TaxAmount, TaxAmountFC, DocCurrency,
                                                    "", "", "", "", "", "", "", "");

                            reLinesRow = reLines.Rows.Add();

                            reLinesRow["debitAccount"] = DebitAccount;
                            reLinesRow["creditAccount"] = CreditAccount;
                            reLinesRow["prBase"] = prBase;
                            reLinesRow["txType"] = "Uncrediting";
                            reLinesRow["amtTx"] = BaseGross;
                            reLinesRow["amtPrTx"] = TaxAmount;

                        }
                    }
                }
            }

            //დამსაქმებლის საპენსიო გატარება
            string wtCode = BPDataSourceTable.GetValue("WTCode", 0);
            bool physicalEntityTax = (BPDataSourceTable.GetValue("WTLiable", 0) == "Y" &&
                                       docDBSources.Item("OWHT").GetValue("U_BDOSPhisTx", 0) == "Y");
            if (physicalEntityTax)
            {
                string pensionCoWTCode = CommonFunctions.getOADM("U_BDOSPnCoP").ToString();
                string pensionPhWTCode = CommonFunctions.getOADM("U_BDOSPnPh").ToString();
                string CreditAccount = CommonFunctions.getValue("OWHT", "Account", "WTCode", pensionCoWTCode).ToString();
                bool wt_InvoiceType = CommonFunctions.getValue("OWHT", "Category", "WTCode", wtCode).ToString() == "I";
                string expAcct = CommonFunctions.getValue("OWHT", "U_BDOSExpAcc", "WTCode", pensionCoWTCode).ToString();

                //    string expQuery = "select \"U_BDOSExpAcc\" from OWHT where \"WTCode\"='" + pensionCoWTCode + "'";
                //    SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                //    oRecordSet.DoQuery(expQuery);

                //    if (!oRecordSet.EoF) expAcct = oRecordSet.Fields.Item("U_BDOSExpAcc").Value;
                decimal CompanyPensionAmount;
                decimal CompanyPensionAmountFC;
                decimal PhysPensionAmount;
                decimal PhysPensionAmountFC;
                decimal WhtAmount;
                decimal WhtAmountFC;
                string DebitAccount;
                string Project;
                string DistrRule1;
                string DistrRule2;
                string DistrRule3;
                string DistrRule4;
                string DistrRule5;

                DBDataSourceTable = docDBSources.Item("PCH1");
                JEcount = DBDataSourceTable.Size;

                var dpmCompanyPensionAmt = decimal.Zero;
                var companyPensionAmtSum = decimal.Zero;
                if (oForm.DataSources.DBDataSources.Item("PCH9").Size > 0) //Company Pension for AP Down Payment Request)
                {
                    dpmCompanyPensionAmt = GetAPDownPaymentDocsCompanyPensionAmt(oForm);
                    for (int i = 0; i < JEcount; i++)
                        companyPensionAmtSum += Convert.ToDecimal(CommonFunctions.getChildOrDbDataSourceValue(DBDataSourceTable, null, DTSource, "U_BDOSPnCoAm", i), CultureInfo.InvariantCulture);
                }
                for (int i = 0; i < JEcount; i++)
                {
                    CompanyPensionAmount = Convert.ToDecimal(CommonFunctions.getChildOrDbDataSourceValue(DBDataSourceTable, null, DTSource, "U_BDOSPnCoAm", i), CultureInfo.InvariantCulture);
                    PhysPensionAmount = Convert.ToDecimal(CommonFunctions.getChildOrDbDataSourceValue(DBDataSourceTable, null, DTSource, "U_BDOSPnPhAm", i), CultureInfo.InvariantCulture);
                    WhtAmount = Convert.ToDecimal(CommonFunctions.getChildOrDbDataSourceValue(DBDataSourceTable, null, DTSource, "U_BDOSWhtAmt", i), CultureInfo.InvariantCulture);

                    Project = CommonFunctions.getChildOrDbDataSourceValue(DBDataSourceTable, null, DTSource, "Project", i).ToString();
                    DistrRule1 = CommonFunctions.getChildOrDbDataSourceValue(DBDataSourceTable, null, DTSource, "OcrCode", i).ToString();
                    DistrRule2 = CommonFunctions.getChildOrDbDataSourceValue(DBDataSourceTable, null, DTSource, "OcrCode2", i).ToString();
                    DistrRule3 = CommonFunctions.getChildOrDbDataSourceValue(DBDataSourceTable, null, DTSource, "OcrCode3", i).ToString();
                    DistrRule4 = CommonFunctions.getChildOrDbDataSourceValue(DBDataSourceTable, null, DTSource, "OcrCode4", i).ToString();
                    DistrRule5 = CommonFunctions.getChildOrDbDataSourceValue(DBDataSourceTable, null, DTSource, "OcrCode5", i).ToString();

                    CompanyPensionAmount += companyPensionAmtSum == decimal.Zero ? decimal.Zero : CompanyPensionAmount * dpmCompanyPensionAmt / companyPensionAmtSum;

                    if (CompanyPensionAmount != 0)
                    {
                        CompanyPensionAmountFC = DocCurrency == "" ? 0 : CompanyPensionAmount / DocRate;
                        if (!wt_InvoiceType)
                        {
                            DebitAccount = i == 0 && expAcct != "" && U_BDOSPnAcc == "Y" ? expAcct : CommonFunctions.getChildOrDbDataSourceValue(DBDataSourceTable, null, DTSource, "AcctCode", i).ToString();
                            JournalEntry.AddJournalEntryRow(AccountTable, jeLines, "Full", DebitAccount, CreditAccount, CompanyPensionAmount, CompanyPensionAmountFC, DocCurrency, DistrRule1, DistrRule2, DistrRule3, DistrRule4, DistrRule5, Project, "", "");
                        }
                        //Invoice შემთხვევაში
                        else
                        {
                            DebitAccount = i == 0 && expAcct != "" && U_BDOSPnAcc == "Y" ? expAcct : CommonFunctions.getChildOrDbDataSourceValue(DBDataSourceTable, null, DTSource, "AcctCode", i).ToString();
                            CreditAccount = CommonFunctions.getValue("OWHT", "U_BdgtDbtAcc", "WTCode", pensionCoWTCode).ToString(); // დამსაქმებლის საპენსიოს ვალდებულების ანგარიში
                            JournalEntry.AddJournalEntryRow(AccountTable, jeLines, "Full", DebitAccount, CreditAccount, CompanyPensionAmount, CompanyPensionAmountFC, DocCurrency, DistrRule1, DistrRule2, DistrRule3, DistrRule4, DistrRule5, Project, "", "");
                        }
                    }

                    // Invoice ტიპის შემთხვევაში
                    if (wt_InvoiceType && WhtAmount != 0 && PhysPensionAmount != 0)
                    {
                        PhysPensionAmountFC = DocCurrency == "" ? 0 : PhysPensionAmount / DocRate;
                        WhtAmountFC = DocCurrency == "" ? 0 : WhtAmount / DocRate;
                        //DebitAccount = i == 0 && expAcct != "" ? expAcct : CommonFunctions.getChildOrDbDataSourceValue(DBDataSourceTable, null, DTSource, "AcctCode", i).ToString();//BP-ს ძირითადი WTCode-ს ანგარიში
                        DebitAccount = CommonFunctions.getValue("OWHT", "Account", "WTCode", wtCode).ToString();
                        JournalEntry.AddJournalEntryRow(AccountTable, jeLines, "OnlyDebit", DebitAccount, "", (WhtAmount + PhysPensionAmount), (WhtAmountFC + PhysPensionAmountFC), DocCurrency,
                                                            DistrRule1, DistrRule2, DistrRule3, DistrRule4, DistrRule5, Project, "", "");

                        CreditAccount = CommonFunctions.getValue("OWHT", "U_BdgtDbtAcc", "WTCode", wtCode).ToString(); //BP-ს ძირითადი WTCode-ს ვალდებულების ანგარიში
                        JournalEntry.AddJournalEntryRow(AccountTable, jeLines, "OnlyCredit", "", CreditAccount, WhtAmount, WhtAmountFC, DocCurrency,
                                                            DistrRule1, DistrRule2, DistrRule3, DistrRule4, DistrRule5, Project, "", "");

                        CreditAccount = CommonFunctions.getValue("OWHT", "Account", "WTCode", pensionPhWTCode).ToString(); //U_BdgtDbtAcc დასაქმებულის საპენსიოს ვალდებულების ანგარიში
                        JournalEntry.AddJournalEntryRow(AccountTable, jeLines, "OnlyCredit", "", CreditAccount, PhysPensionAmount, PhysPensionAmountFC, DocCurrency,
                                        DistrRule1, DistrRule2, DistrRule3, DistrRule4, DistrRule5, Project, "", "");
                    }
                }
            }
            return jeLines;
        }

        public static void JrnEntry(string DocEntry, string DocNum, DateTime DocDate, DataTable JrnLinesDT, DataTable reLines, out string errorText)
        {
            errorText = null;

            try
            {
                JournalEntry.JrnEntry(DocEntry, "18", "AP Invoice: " + DocNum, DocDate, JrnLinesDT, out errorText);

                if (errorText != null)
                {
                    return;
                }

                for (int i = 0; i < reLines.Rows.Count; i++)
                {
                    reLines.Rows[i]["DocEntry"] = DocEntry == "" ? 0 : Convert.ToInt32(DocEntry);
                    reLines.Rows[i]["DocNum"] = DocNum.ToString();
                    reLines.Rows[i]["docDate"] = DocDate;
                }

                ProfitTax.AddRecord(reLines, "18", "AP Invoice: " + DocNum, out errorText);
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
        }

        private static decimal GetAPDownPaymentDocsCompanyPensionAmt(SAPbouiCOM.Form oForm)
        {
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                SAPbouiCOM.DBDataSource oDBDataSourceDrawnDpm = oForm.DataSources.DBDataSources.Item("PCH9");

                StringBuilder dummyTable = new StringBuilder();
                var dpmDatasourceSize = oDBDataSourceDrawnDpm.Size;
                for (int i = 0; i < dpmDatasourceSize; i++)
                {
                    dummyTable.AppendLine($"SELECT {oDBDataSourceDrawnDpm.GetValue("BaseAbs", i)} AS \"BaseAbs\", {oDBDataSourceDrawnDpm.GetValue("Gross", i)} AS \"Gross\" ");
                    if (Program.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                        dummyTable.Append(" FROM DUMMY ");
                    if (i != (dpmDatasourceSize - 1))
                        dummyTable.AppendLine(" UNION ALL ");
                }

                StringBuilder query = new StringBuilder();
                query.Append("SELECT SUM(T0.\"Gross\" * T0.\"U_BDOSPnPhAm\" / T0.\"DocTotal\") AS \"DownPaymentPencCo\" \n");
                query.Append("FROM   (SELECT T1.\"BaseAbs\", \n");
                query.Append("               T1.\"Gross\", \n");
                query.Append("               T2.\"DocTotal\", \n");
                query.Append("               SUM(T3.\"U_BDOSPnPhAm\") AS \"U_BDOSPnPhAm\" \n");
                query.Append($"        FROM  ({dummyTable}) T1 \n");
                query.Append("               INNER JOIN \"ODPO\" T2 \n");
                query.Append("                       ON T1.\"BaseAbs\" = T2.\"DocEntry\" \n");
                query.Append("               INNER JOIN \"DPO1\" T3 \n");
                query.Append("                       ON T1.\"BaseAbs\" = T3.\"DocEntry\" \n");
                query.Append("        GROUP  BY T1.\"BaseAbs\", \n");
                query.Append("                  T1.\"Gross\", \n");
                query.Append("                  T2.\"DocTotal\") T0");

                oRecordSet.DoQuery(query.ToString());

                if (!oRecordSet.EoF)
                    return Convert.ToDecimal(oRecordSet.Fields.Item("DownPaymentPencCo").Value);
                else
                    return decimal.Zero;
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

        public static void nonEconExp_OnClick(SAPbouiCOM.Form oForm)
        {
            try
            {
                oForm.Freeze(true);

                string NonEconExp = oForm.DataSources.DBDataSources.Item("OPCH").GetValue("U_nonEconExp", 0).Trim();

                if (NonEconExp != "Y")
                {
                    SAPbouiCOM.EditText PrBaseEdit = oForm.Items.Item("PrBaseE").Specific;
                    PrBaseEdit.Value = "";

                    PrBaseEdit = oForm.Items.Item("PrBsDscr").Specific;
                    PrBaseEdit.Value = "";
                }
            }
            catch { }

            finally
            {
                oForm.Freeze(false);
            }

            setVisibleFormItems(oForm);
        }

        public static void resizeForm(SAPbouiCOM.Form oForm)
        {
            reArrangeFormItems(oForm);
        }

        public static void reArrangeFormItems(SAPbouiCOM.Form oForm)
        {
            oForm.Freeze(true);
            try
            {
                SAPbouiCOM.Item oItem = oForm.Items.Item("90");

                int top = oItem.Top;

                //მოგების გადასახადი
                oItem = oForm.Items.Item("PrBaseS");
                oItem.Top = top;

                oItem = oForm.Items.Item("PrBaseE");
                oItem.Top = top;

                oItem = oForm.Items.Item("PrBsDscr");
                oItem.Top = top;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        public static double GetInvoiceBalanceFC(int docEntry)
        {
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string query = $"SELECT \"DocTotalFC\" - \"PaidFC\" As \"BalanceFC\" FROM \"OPCH\" WHERE \"DocEntry\" = {docEntry}";
            oRecordSet.DoQuery(query);

            if (!oRecordSet.EoF)
                return oRecordSet.Fields.Item("BalanceFC").Value;

            return 0;

        }
        private static void CheckIfDepreciationIsAccrued(SAPbouiCOM.Form oForm, out string deprError)
        {
            string DocDate = oForm.DataSources.DBDataSources.Item("OPCH").GetValue("DocDate", 0);
            string DocNum = oForm.DataSources.DBDataSources.Item("OPCH").GetValue("DocNum", 0);

            StringBuilder query = new StringBuilder();
            query.Append("select distinct \"U_AccrMnth\",\"BatchNum\",\"ItemCode\" \n");
            query.Append("from \"@BDOSDEPAC1\" as \"child\" \n");
            query.Append("inner join \"OIBT\" as \"bt\" \n");
            query.Append("on \"bt\".\"ItemCode\" =\"child\".\"U_ItemCode\" \n");
            query.Append("inner join \"@BDOSDEPACR\" as \"mom\" \n");
            query.Append("on \"mom\".\"DocEntry\"=\"child\".\"DocEntry\" \n");
            query.Append($"where \"BaseNum\" = '{DocNum}' and \"U_AccrMnth\" >= '{DocDate}'");

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRecordSet.DoQuery(query.ToString());
            List<string> recordsetValues = new List<string>();
            if (!oRecordSet.EoF)
            {
                string docItem = oRecordSet.Fields.Item("ItemCode").Value;
                string batchNum = oRecordSet.Fields.Item("BatchNum").Value;
                deprError = BDOSResources.getTranslate("DepreciationIsAccrued") + " Item : " + docItem + " Batch : " + batchNum;
            }
            else
            {
                deprError = string.Empty;
            }
            Marshal.ReleaseComObject(oRecordSet);
        }
    }
}