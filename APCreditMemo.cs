using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Runtime.InteropServices;

namespace BDO_Localisation_AddOn
{
    static partial class APCreditMemo
    {
        public static void JrnEntry(string DocEntry, string DocNum, DateTime DocDate, Decimal rate, string currency, out string errorText)
        {
            errorText = null;

            try
            {
                DataTable jeLines = JournalEntry.JournalEntryTable();
                SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);


            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
        }

        public static void createFormItems(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            BDO_WBReceivedDocs.createFormItems(oForm, "ORPC", out errorText);

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

            oForm.DataSources.UserDataSources.Add("BDO_TaxSer", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
            oForm.DataSources.UserDataSources.Add("BDO_TaxNum", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
            oForm.DataSources.UserDataSources.Add("BDO_TaxDat", SAPbouiCOM.BoDataType.dt_DATE, 20);
            //<-------------------------------------------ანგარიშ-ფაქტურა-----------------------------------

            SAPbouiCOM.Item oItem = oForm.Items.Item("70");
            top = oItem.Top + height * 2 + 1;
            left_s = oItem.Left;
            int width_s = oItem.Width;
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
            formItems.Add("Enabled", true);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top += height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "BDO_TpSt";
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("OperationType"));

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            List<string> listValidValues = new List<string>();
            listValidValues.Add(BDOSResources.getTranslate("Correction")); //0 //კორექტირება
            listValidValues.Add(BDOSResources.getTranslate("Return")); //1 //დაბრუნება

            formItems = new Dictionary<string, object>();
            itemName = "BDO_CNTp";
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "ORPC");
            formItems.Add("Alias", "U_BDO_CNTp");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
            formItems.Add("DisplayDesc", true);
            formItems.Add("ValidValues", listValidValues);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            GC.Collect();
        }

        public static void createUserFields(out string errorText)
        {
            BDO_WBReceivedDocs.createUserFields("ORPC", out errorText);

            List<string> listValidValues;
            Dictionary<string, object> fieldskeysMap;
            listValidValues = new List<string>();
            listValidValues.Add("Correction"); //0 //კორექტირება
            listValidValues.Add("Return"); //1 //დაბრუნება

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDO_CNTp");
            fieldskeysMap.Add("TableName", "ORPC");
            fieldskeysMap.Add("Description", "CreditNote Type");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);
            fieldskeysMap.Add("ValidValues", listValidValues);

            UDO.addUserTableFields(fieldskeysMap, out errorText);
        }

        public static void formDataLoad(SAPbouiCOM.Form oForm)
        {
            oForm.Freeze(true);
            try
            {
                var docEntry = oForm.DataSources.DBDataSources.Item("ORPC").GetValue("DocEntry", 0);
                string cardCode = oForm.DataSources.DBDataSources.Item("ORPC").GetValue("CardCode", 0).Trim();
                string caption = BDOSResources.getTranslate("ChooseTaxInvoice");
                int taxDocEntry = 0;
                string taxID = "";
                string taxNumber = "";
                string taxSeries = "";
                string taxStatus = "";
                string taxCreateDate = "";

                if (!string.IsNullOrEmpty(docEntry))
                {
                    Dictionary<string, object> taxDocInfo = BDO_TaxInvoiceReceived.getTaxInvoiceReceivedDocumentInfo(Convert.ToInt32(docEntry), BDO_TaxInvoiceReceived.BaseDocType.oPurchaseCreditNotes, cardCode);
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
            }
            catch
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
            ""RPC1"".""DocEntry"" AS ""docEntry"",
            ""ORPC"".""U_BDO_WBNo"" AS ""WblNumber"",
            SUM(""RPC1"".""GTotal"") AS ""GTotal"",
            SUM(""RPC1"".""LineVat"") AS ""LineVat""
            FROM ""RPC1"" AS ""RPC1""
            INNER JOIN ""ORPC"" ON ""RPC1"".""DocEntry"" = ""ORPC"".""DocEntry""
            WHERE ""RPC1"".""DocEntry"" = '" + docEntry + @"'
            GROUP BY ""RPC1"".""DocEntry"", 
                     ""ORPC"".""U_BDO_WBNo"" ";

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

        public static void setValues(SAPbouiCOM.Form oForm)
        {
            try
            {
                string docEntry = oForm.DataSources.DBDataSources.Item("ORPC").GetValue("DocEntry", 0).Trim();

                if (!string.IsNullOrEmpty(docEntry))
                {
                    return;
                }
                SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)oForm.Items.Item("BDO_CNTp").Specific;
                oCombo.Select("0", SAPbouiCOM.BoSearchKey.psk_ByValue);
            }
            catch (Exception ex)
            {
                throw ex;
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
                        string wbNumber = oForm.DataSources.DBDataSources.Item("ORPC").GetValue("U_BDO_WBNo", 0).Trim();
                        string cardCode = oForm.DataSources.DBDataSources.Item("ORPC").GetValue("CardCode", 0).Trim();
                        DateTime docDate = DateTime.ParseExact(oForm.DataSources.DBDataSources.Item("ORPC").GetValue("DocDate", 0), "yyyyMMdd", CultureInfo.InvariantCulture);
                        //string baseType = oForm.DataSources.DBDataSources.Item("RPC1").GetValue("BaseType", 0).Trim();

                        List<string> taxInvoiceDocList = BDO_TaxInvoiceReceived.getListTaxInvoiceReceived(cardCode, wbNumber, BDO_TaxInvoiceReceived.BaseDocType.oPurchaseCreditNotes, docDate);

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
                string docEntrySTR = oForm.DataSources.DBDataSources.Item("ORPC").GetValue("DocEntry", 0);

                oForm.Items.Item("16").Click(); //focus on remark item

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
                //errorText = ex.Message;
            }
            finally
            {
                oForm.Freeze(false);
                oForm.Update();
            }

            FormsB1.WB_TAX_AuthorizationsItems(oForm);
        }

        public static void formDataAddUpdate(SAPbouiCOM.Form oForm)
        {
            try
            {
                string taxDocEntry = oForm.DataSources.UserDataSources.Item("BDO_TaxDoc").ValueEx;
                string docEntryStr = oForm.DataSources.DBDataSources.Item("ORPC").GetValue("DocEntry", 0);
                if (!string.IsNullOrEmpty(taxDocEntry) && !string.IsNullOrEmpty(docEntryStr))
                {
                    int docEntry = Convert.ToInt32(docEntryStr);
                    getAmount(docEntry, out var baseDocGTotal, out var baseDocLineVat, out var baseDocWblNmber);
                    BDO_TaxInvoiceReceived.addBaseDoc(Convert.ToInt32(taxDocEntry), docEntry, BDO_TaxInvoiceReceived.BaseDocType.oPurchaseCreditNotes, baseDocWblNmber, baseDocGTotal, baseDocLineVat);
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
            string errorText = null;

            SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(BusinessObjectInfo.FormTypeEx, Program.currentFormCount);

            if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD)
            {
                if (BusinessObjectInfo.BeforeAction)
                {
                    SAPbouiCOM.DBDataSource DocDBSource = oForm.DataSources.DBDataSources.Item(0);
                    if (DocDBSource.GetValue("CANCELED", 0) == "N")
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

                        SAPbouiCOM.ComboBox opType = (SAPbouiCOM.ComboBox)oForm.Items.Item("BDO_CNTp").Specific;
                        if (opType.Value == "1")
                        {
                            BDO_WBReceivedDocs.ClearWaybillItemsValues(oForm);
                        }

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
                        DataTable JrnLinesDT = createAdditionalEntries(oForm, null, null, DocCurrency, DocRate);

                        JrnEntry(DocEntry, DocNum, DocDate, JrnLinesDT, out errorText);
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

            //A/C Number Update
            if ((BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD || BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE)
                && BusinessObjectInfo.ActionSuccess && BusinessObjectInfo.BeforeAction == false)
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
                }
                else if (BusinessObjectInfo.BeforeAction == false && BusinessObjectInfo.ActionSuccess)
                {
                    setVisibleFormItems(oForm);
                }
            }

            if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD && BusinessObjectInfo.BeforeAction == false)
            {
                //when "Keep Visible" is not selected Program.uiApp.Forms.ActiveForm.Type = 10018, so we need check
                if (Program.uiApp.Forms.ActiveForm.Type == 181) // Keep Visible Case
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

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD && pVal.BeforeAction)
                {
                    createFormItems(oForm, out errorText);
                    formDataLoad(oForm);
                    setVisibleFormItems(oForm);
                }
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD && !pVal.BeforeAction)
                {
                    setValues(oForm);
                }
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_COMBO_SELECT && pVal.ItemUID == "WBOper" && pVal.BeforeAction == false)
                {
                    Program.oIncWaybDocFormCrMemo = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                    oForm.Freeze(true);
                    BDO_WBReceivedDocs.comboSelect(oForm, Program.oIncWaybDocFormCrMemo, "CreditMemo", out errorText);
                    oForm.Freeze(false);
                }

                if ((pVal.ItemUID == "BDO_TaxDoc") && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                {
                    SAPbouiCOM.IChooseFromListEvent oCFLEvento = (SAPbouiCOM.IChooseFromListEvent)pVal;
                    chooseFromList(oForm, oCFLEvento, pVal.ItemUID, pVal.BeforeAction);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_GOT_FOCUS && pVal.ItemUID == "BDO_TaxDoc" && !pVal.BeforeAction)
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

                    int taxDocEntry = Convert.ToInt32(oForm.DataSources.UserDataSources.Item("BDO_TaxDoc").ValueEx.Trim());
                    int docEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("ORPC").GetValue("DocEntry", 0));
                    if (taxDocEntry != 0)
                    {
                        int answer = Program.uiApp.MessageBox(BDOSResources.getTranslate("DocumentLinkedToTaxInvoiceCancel"), 1, BDOSResources.getTranslate("Yes"), BDOSResources.getTranslate("No"), "");
                        if (answer == 1)
                        {
                            BDO_TaxInvoiceReceived.removeBaseDoc(taxDocEntry, docEntry, BDO_TaxInvoiceReceived.BaseDocType.oPurchaseCreditNotes);
                            formDataLoad(oForm);
                            setVisibleFormItems(oForm);
                        }
                    }
                }
            }
        }

        public static DataTable createAdditionalEntries(SAPbouiCOM.Form oForm, SAPbobsCOM.GeneralData oGeneralData, DataTable DTSource, string DocCurrency, decimal DocRate)
        {
            DataTable jeLines = JournalEntry.JournalEntryTable();
            DocCurrency = DocCurrency == CommonFunctions.getLocalCurrency() ? "" : DocCurrency;
            DataTable AccountTable = CommonFunctions.GetOACTTable();
            SAPbouiCOM.DBDataSources docDBSources = oForm.DataSources.DBDataSources;
            SAPbouiCOM.DBDataSource DBDataSourceTable = null;
            int JEcount = 0;

            if (oForm == null)
            {
                JEcount = DTSource.Rows.Count;
            }
            else
            {
                DBDataSourceTable = docDBSources.Item("RPC1");
                JEcount = DBDataSourceTable.Size;
            }

            SAPbouiCOM.DBDataSource BPDataSourceTable = docDBSources.Item("OCRD");

            string CardCode = BPDataSourceTable.GetValue("CardCode", 0).Trim();
            string vatCode = BPDataSourceTable.GetValue("ECVatGroup", 0).Trim();
            string TaxType = BPDataSourceTable.GetValue("U_BDO_TaxTyp", 0).Trim();


            //დამსაქმებლის საპენსიო გატარება
            string wtCode = BPDataSourceTable.GetValue("WTCode", 0);
            bool physicalEntityTax = (BPDataSourceTable.GetValue("WTLiable", 0) == "Y" &&
                                        docDBSources.Item("OWHT").GetValue("U_BDOSPhisTx", 0) == "Y");
            if (physicalEntityTax)
            {
                string pensionCoWTCode = CommonFunctions.getOADM("U_BDOSPnCoP").ToString();
                SAPbobsCOM.WithholdingTaxCodes oWHTaxCodeCo;
                oWHTaxCodeCo = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oWithholdingTaxCodes);
                oWHTaxCodeCo.GetByKey(pensionCoWTCode);

                decimal CompanyPensionAmount;
                decimal CompanyPensionAmountFC;
                string DebitAccount;
                string Project;
                string DistrRule1;
                string DistrRule2;
                string DistrRule3;
                string DistrRule4;
                string DistrRule5;

                for (int i = 0; i < JEcount; i++)
                {
                    CompanyPensionAmount = Convert.ToDecimal(CommonFunctions.getChildOrDbDataSourceValue(DBDataSourceTable, null, DTSource, "U_BDOSPnCoAm", i), CultureInfo.InvariantCulture);
                    if (CompanyPensionAmount > 0)
                    {
                        CompanyPensionAmountFC = DocCurrency == "" ? 0 : CompanyPensionAmount / DocRate;

                        DebitAccount = CommonFunctions.getChildOrDbDataSourceValue(DBDataSourceTable, null, DTSource, "AcctCode", i).ToString();

                        Project = CommonFunctions.getChildOrDbDataSourceValue(DBDataSourceTable, null, DTSource, "Project", i).ToString();
                        DistrRule1 = CommonFunctions.getChildOrDbDataSourceValue(DBDataSourceTable, null, DTSource, "OcrCode", i).ToString();
                        DistrRule2 = CommonFunctions.getChildOrDbDataSourceValue(DBDataSourceTable, null, DTSource, "OcrCode2", i).ToString();
                        DistrRule3 = CommonFunctions.getChildOrDbDataSourceValue(DBDataSourceTable, null, DTSource, "OcrCode3", i).ToString();
                        DistrRule4 = CommonFunctions.getChildOrDbDataSourceValue(DBDataSourceTable, null, DTSource, "OcrCode4", i).ToString();
                        DistrRule5 = CommonFunctions.getChildOrDbDataSourceValue(DBDataSourceTable, null, DTSource, "OcrCode5", i).ToString();


                        JournalEntry.AddJournalEntryRow(AccountTable, jeLines, "Full", DebitAccount, oWHTaxCodeCo.Account, CompanyPensionAmount * (-1), CompanyPensionAmountFC * (-1), DocCurrency, DistrRule1, DistrRule2, DistrRule3, DistrRule4, DistrRule5, Project, "", "");
                    }
                }
            }

            return jeLines;

        }

        public static void JrnEntry(string DocEntry, string DocNum, DateTime DocDate, DataTable JrnLinesDT, out string errorText)
        {
            try
            {
                JournalEntry.JrnEntry(DocEntry, "19", "AP Credit Note: " + DocNum, DocDate, JrnLinesDT, out errorText);

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

        public static void cancellation(SAPbouiCOM.Form oForm, int docEntry)
        {
            try
            {
                JournalEntry.cancellation(oForm, docEntry, "19", out var errorText);
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

        public static List<int> getAllConnectedDoc(List<int> docEntry, string baseType)
        {
            List<int> connectedDocList = new List<int>();

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string query = @"SELECT
            	 ""RPC1"".""DocEntry""
            FROM ""RPC1""
            WHERE ""RPC1"".""BaseEntry"" IN (" + string.Join(",", docEntry) + @")
            AND ""RPC1"".""BaseType"" = '" + baseType + @"'
            GROUP BY ""RPC1"".""DocEntry""";

            try
            {
                oRecordSet.DoQuery(query);
                while (!oRecordSet.EoF)
                {
                    connectedDocList.Add(Convert.ToInt32(oRecordSet.Fields.Item("DocEntry").Value));
                    oRecordSet.MoveNext();
                }
                return connectedDocList;
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
    }
}
