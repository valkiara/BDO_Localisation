using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using SAPbouiCOM;
using DataTable = System.Data.DataTable;

namespace BDO_Localisation_AddOn
{
    static partial class ARInvoice
    {
        private static Dictionary<int,decimal> InitialItemGrossPrices = new Dictionary<int, decimal>();
        public static void createFormItems(Form oForm, out string errorText)
        {
            errorText = null;
            Dictionary<string, object> formItems = null;

            string itemName = "";

            //<-------------------------------------------სასაქონლო ზედნადები----------------------------------->
            double height = oForm.Items.Item("86").Height;
            double top = oForm.Items.Item("86").Top + height * 1.5 + 1;
            double left_s = oForm.Items.Item("86").Left;
            double left_e = oForm.Items.Item("46").Left;
            double width_e = oForm.Items.Item("46").Width;

            string caption = BDOSResources.getTranslate("CreateWaybill");
            formItems = new Dictionary<string, object>();
            itemName = "BDO_WblTxt"; //10 characters
            formItems.Add("Type", BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_e * 1.5);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", caption);
            formItems.Add("TextStyle", 4);
            formItems.Add("FontSize", 10);
            formItems.Add("Enabled", true);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            bool multiSelection = false;
            string objectType = "UDO_F_BDO_WBLD_D"; //Waybill document
            string uniqueID_WaybillCFL = "Waybill_CFL";
            FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_WaybillCFL);

            formItems = new Dictionary<string, object>();
            itemName = "BDO_WblDoc"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "UserDataSources");
            formItems.Add("DataType", BoDataType.dt_SHORT_TEXT);
            formItems.Add("Length", 11);
            formItems.Add("TableName", "");
            formItems.Add("Alias", itemName);
            formItems.Add("Bound", true);
            formItems.Add("Type", BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e + width_e - 40);
            formItems.Add("Width", 40);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("AffectsFormMode", false);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", false);
            formItems.Add("ChooseFromListUID", uniqueID_WaybillCFL);
            formItems.Add("ChooseFromListAlias", "DocEntry");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "BDO_WblLB"; //10 characters
            formItems.Add("Type", BoFormItemTypes.it_LINKED_BUTTON);
            formItems.Add("Left", left_e + width_e - 40 - 20);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("LinkTo", "BDO_WblDoc");
            formItems.Add("LinkedObjectType", objectType);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            oForm.DataSources.UserDataSources.Add("BDO_WblID", BoDataType.dt_SHORT_TEXT, 20);
            oForm.DataSources.UserDataSources.Add("BDO_WblNum", BoDataType.dt_SHORT_TEXT, 20);
            oForm.DataSources.UserDataSources.Add("BDO_WblSts", BoDataType.dt_SHORT_TEXT, 50);
            //<-------------------------------------------სასაქონლო ზედნადები-----------------------------------

            //-------------------------------------------ანგარიშ-ფაქტურა----------------------------------->
            top = top + height * 1.5 + 1;

            formItems = new Dictionary<string, object>();
            itemName = "BDO_TaxTxt"; //10 characters
            formItems.Add("Type", BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_e * 1.5);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("CreateTaxInvoice"));
            formItems.Add("TextStyle", 4);
            formItems.Add("FontSize", 10);
            formItems.Add("Enabled", true);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            multiSelection = false;
            objectType = "UDO_F_BDO_TAXS_D"; //Tax invoice sent document
            string uniqueID_TaxInvoiceSentCFL = "TaxInvoiceSent_CFL";
            FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_TaxInvoiceSentCFL);

            formItems = new Dictionary<string, object>();
            itemName = "BDO_TaxDoc"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "UserDataSources");
            formItems.Add("DataType", BoDataType.dt_SHORT_TEXT);
            formItems.Add("Length", 11);
            formItems.Add("TableName", "");
            formItems.Add("Alias", itemName);
            formItems.Add("Bound", true);
            formItems.Add("Type", BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e + width_e - 40);
            formItems.Add("Width", 40);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("AffectsFormMode", false);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", false);
            formItems.Add("ChooseFromListUID", uniqueID_TaxInvoiceSentCFL);
            formItems.Add("ChooseFromListAlias", "DocEntry");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "BDO_TaxLB"; //10 characters
            formItems.Add("Type", BoFormItemTypes.it_LINKED_BUTTON);
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

            top = top + height + 1;

            oForm.DataSources.UserDataSources.Add("BDO_TaxSer", BoDataType.dt_SHORT_TEXT, 20);
            oForm.DataSources.UserDataSources.Add("BDO_TaxNum", BoDataType.dt_SHORT_TEXT, 20);
            oForm.DataSources.UserDataSources.Add("BDO_TaxDat", BoDataType.dt_DATE, 20);
            //<-------------------------------------------ანგარიშ-ფაქტურა-----------------------------------


            // -------------------- Use blanket agreement rates-----------------
            int pane = 7;
            int left = oForm.Items.Item("1720002167").Left;
            height = oForm.Items.Item("1720002167").Height;
            top = oForm.Items.Item("1720002167").Top + height + 5;

            formItems = new Dictionary<string, object>();
            itemName = "UsBlaAgRtS"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OINV");
            formItems.Add("Alias", "U_UseBlaAgRt");
            formItems.Add("Bound", true);
            formItems.Add("Type", BoFormItemTypes.it_CHECK_BOX);
            formItems.Add("DataType", BoDataType.dt_SHORT_TEXT);
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
            formItems.Add("SetAutoManaged", true);
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            #region Discount field

            height = oForm.Items.Item("42").Height;
            top = oForm.Items.Item("42").Top;
            left_e = oForm.Items.Item("42").Left;
            width_e = oForm.Items.Item("42").Width;

            formItems = new Dictionary<string, object>();
            itemName = "DiscountE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OINV");
            formItems.Add("Alias", "U_Discount");
            formItems.Add("Bound", true);
            formItems.Add("Type", BoFormItemTypes.it_EDIT);
            formItems.Add("DataType", BoDataType.dt_SUM);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Discount"));
            formItems.Add("DisplayDesc", true);
            formItems.Add("SetAutoManaged", true);
            formItems.Add("Visible", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            #endregion

            GC.Collect();
        }

        public static void createUserFields(out string errorText)
        {
            errorText = null;

            Dictionary<string, object> fieldskeysMap;

            #region UseBlaAgRt

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "UseBlaAgRt");
            fieldskeysMap.Add("TableName", "OINV");
            fieldskeysMap.Add("Description", "Use Blanket Agreement Rates");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 1);
            fieldskeysMap.Add("DefaultValue", "N");

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            #endregion

            #region Discount

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "Discount");
            fieldskeysMap.Add("TableName", "OINV");
            fieldskeysMap.Add("Description", "Discount Amount");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            #endregion

            GC.Collect();
        }

        public static void formDataLoad(Form oForm, out string errorText)
        {
            errorText = null;

            StaticText oStaticText = null;
            oForm.Freeze(true);
            try
            {
                int docEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("OINV").GetValue("DocEntry", 0));

                //-------------------------------------------სასაქონლო ზედნადები----------------------------------->
                string caption = BDOSResources.getTranslate("CreateWaybill");
                int wblDocEntry = 0;
                string wblID = "";
                string wblNum = "";
                string wblSts = "";
                string objType = "13";

                if (docEntry != 0)
                {
                    Dictionary<string, string> wblDocInfo = BDO_Waybills.getWaybillDocumentInfo(docEntry, objType, out errorText);
                    wblDocEntry = Convert.ToInt32(wblDocInfo["DocEntry"]);
                    wblID = wblDocInfo["wblID"];
                    wblNum = wblDocInfo["number"];
                    wblSts = wblDocInfo["status"];

                    if (wblDocEntry != 0)
                    {
                        caption = BDOSResources.getTranslate("Wb") + ": " + wblSts + " " + wblID + (wblNum != "" ? " № " + wblNum : "");
                    }
                }
                else
                {
                    caption = BDOSResources.getTranslate("CreateWaybill");
                    wblDocEntry = 0;
                }

                oForm.DataSources.UserDataSources.Item("BDO_WblDoc").ValueEx = wblDocEntry == 0 ? "" : wblDocEntry.ToString();
                oForm.DataSources.UserDataSources.Item("BDO_WblID").ValueEx = wblID;
                oForm.DataSources.UserDataSources.Item("BDO_WblNum").ValueEx = wblNum;
                oForm.DataSources.UserDataSources.Item("BDO_WblSts").ValueEx = wblSts;

                oStaticText = (StaticText)oForm.Items.Item("BDO_WblTxt").Specific;
                oStaticText.Caption = caption;
                //<-------------------------------------------სასაქონლო ზედნადები-----------------------------------

                //-------------------------------------------ანგარიშ-ფაქტურა----------------------------------->
                string cardCode = oForm.DataSources.DBDataSources.Item("OINV").GetValue("CardCode", 0).Trim();
                caption = BDOSResources.getTranslate("CreateTaxInvoice");
                int taxDocEntry = 0;
                string taxID = "";
                string taxNumber = "";
                string taxSeries = "";
                string taxStatus = "";
                string taxCreateDate = "";

                if (docEntry != 0)
                {
                    Dictionary<string, object> taxDocInfo = BDO_TaxInvoiceSent.getTaxInvoiceSentDocumentInfo(docEntry, "ARInvoice", cardCode);
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
                            DateTime taxCreateDateDT = DateTime.ParseExact(taxCreateDate, "yyyyMMdd", CultureInfo.InvariantCulture);
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

                oForm.Items.Item("BDO_TaxDoc").Enabled = false;
                oForm.Items.Item("BDO_WblDoc").Enabled = false;



                oStaticText = (StaticText)oForm.Items.Item("BDO_TaxTxt").Specific;
                oStaticText.Caption = caption;
                //<-------------------------------------------ანგარიშ-ფაქტურა-----------------------------------
            }
            catch (Exception ex)
            {
                oForm.DataSources.UserDataSources.Item("BDO_WblDoc").ValueEx = "";
                oForm.DataSources.UserDataSources.Item("BDO_WblID").ValueEx = "";
                oForm.DataSources.UserDataSources.Item("BDO_WblNum").ValueEx = "";
                oForm.DataSources.UserDataSources.Item("BDO_WblSts").ValueEx = "";

                oStaticText = (StaticText)oForm.Items.Item("BDO_WblTxt").Specific;
                oStaticText.Caption = BDOSResources.getTranslate("CreateWaybill");

                oForm.DataSources.UserDataSources.Item("BDO_TaxDoc").ValueEx = "";
                oForm.DataSources.UserDataSources.Item("BDO_TaxSer").ValueEx = "";
                oForm.DataSources.UserDataSources.Item("BDO_TaxNum").ValueEx = "";
                oForm.DataSources.UserDataSources.Item("BDO_TaxDat").ValueEx = "";

                oStaticText = (StaticText)oForm.Items.Item("BDO_TaxTxt").Specific;
                oStaticText.Caption = BDOSResources.getTranslate("CreateTaxInvoice");

                errorText = ex.Message;
            }
            finally
            {
                oForm.Freeze(false);
                GC.Collect();
            }
        }

        public static void cancellation(Form oForm, int docEntry, out string errorText)
        {
            errorText = null;

            try
            {
                Dictionary<string, string> wblDocInfo = BDO_Waybills.getWaybillDocumentInfo(docEntry, "13", out errorText);
                int wblDocEntry = Convert.ToInt32(wblDocInfo["DocEntry"]);

                if (wblDocEntry != 0)
                {
                    int answer = Program.uiApp.MessageBox(BDOSResources.getTranslate("DocumentLinkedToWaybillCancel"), 1, BDOSResources.getTranslate("Yes"), BDOSResources.getTranslate("No"), "");
                    string operation = answer == 1 ? "Update" : "Cancel";
                    BDO_Waybills.cancellation(wblDocEntry, operation, out errorText);
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }

            try
            {
                JournalEntry.cancellation(oForm, docEntry, "13", out errorText);
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

        public static void itemPressed(Form oForm, ItemEvent pVal, out int newDocEntry, out string bstrUDOObjectType)
        {
            string errorText = null;
            newDocEntry = 0;
            bstrUDOObjectType = null;

            try
            {
                oForm.Freeze(true);

                string docEntrySTR = oForm.DataSources.DBDataSources.Item("OINV").GetValue("DocEntry", 0);
                docEntrySTR = string.IsNullOrEmpty(docEntrySTR) ? "0" : docEntrySTR;
                int docEntry = Convert.ToInt32(docEntrySTR);
                string cancelled = oForm.DataSources.DBDataSources.Item("OINV").GetValue("CANCELED", 0).Trim();
                string docType = oForm.DataSources.DBDataSources.Item("OINV").GetValue("DocType", 0).Trim();
                string objectType = "13";

                if (pVal.ItemUID == "BDO_WblTxt")
                {
                    string wblDoc = oForm.DataSources.UserDataSources.Item("BDO_WblDoc").ValueEx;
                    bstrUDOObjectType = "UDO_F_BDO_WBLD_D";

                    if (docEntry != 0 && (oForm.Mode == BoFormMode.fm_OK_MODE || oForm.Mode == BoFormMode.fm_VIEW_MODE))
                    {
                        if (wblDoc == "" && cancelled == "N" && docType == "I")
                        {
                            BDO_Waybills.createDocument(objectType, docEntry, null, null, null, null, out newDocEntry, out errorText);
                            if (errorText == null && newDocEntry != 0)
                            {
                                oForm.DataSources.UserDataSources.Item("BDO_WblDoc").ValueEx = newDocEntry.ToString();
                                formDataLoad(oForm, out errorText);
                                return;
                            }
                        }
                        else if (cancelled != "N")
                            throw new Exception(BDOSResources.getTranslate("DocumentMustNotBeCancelledOrCancellation"));
                        else if (docType != "I")
                            throw new Exception(BDOSResources.getTranslate("DocumentTypeMustBeItem"));
                    }
                    else
                        throw new Exception(BDOSResources.getTranslate("ToCreateWaybillWriteDocument"));
                }

                else if (pVal.ItemUID == "BDO_TaxTxt")
                {
                    string taxDoc = oForm.DataSources.UserDataSources.Item("BDO_TaxDoc").ValueEx;
                    bstrUDOObjectType = "UDO_F_BDO_TAXS_D";

                    if (docEntry != 0 && (oForm.Mode == BoFormMode.fm_OK_MODE || oForm.Mode == BoFormMode.fm_VIEW_MODE))
                    {
                        if (taxDoc == "" && cancelled == "N")
                        {
                            BDO_TaxInvoiceSent.createDocument(objectType, docEntry, "", true, 0, null, false, null, null, out newDocEntry, out errorText);
                            if (string.IsNullOrEmpty(errorText) && newDocEntry != 0)
                            {
                                oForm.DataSources.UserDataSources.Item("BDO_TaxDoc").ValueEx = newDocEntry.ToString();
                                formDataLoad(oForm, out errorText);
                                return;
                            }
                        }
                        else if (cancelled != "N")
                            throw new Exception(BDOSResources.getTranslate("DocumentMustNotBeCancelledOrCancellation"));
                    }
                    else
                        throw new Exception(BDOSResources.getTranslate("ToCreateTaxInvoiceWriteDocument"));
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

        public static void getAmount(int docEntry, out double gTotal, out double lineVat, out string errorText)
        {
            errorText = null;
            gTotal = 0;
            lineVat = 0;

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string query = @"SELECT 
            ""INV1"".""DocEntry"" AS ""docEntry"", 
            SUM(""INV1"".""GTotal"") AS ""GTotal"", 
            SUM(""INV1"".""LineVat"") AS ""LineVat"" 
            FROM ""INV1"" AS ""INV1"" 
            WHERE ""INV1"".""DocEntry"" = '" + docEntry + @"' 
            GROUP BY ""INV1"".""DocEntry""";

            try
            {
                oRecordSet.DoQuery(query);

                while (!oRecordSet.EoF)
                {
                    gTotal = oRecordSet.Fields.Item("GTotal").Value;
                    lineVat = oRecordSet.Fields.Item("LineVat").Value;

                    oRecordSet.MoveNext();
                    break;
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                oRecordSet = null;
                GC.Collect();
            }
        }

        public static void uiApp_FormDataEvent(ref BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            Form oForm = Program.uiApp.Forms.GetForm(BusinessObjectInfo.FormTypeEx, Program.currentFormCount);

            if (oForm.TypeEx == "133")
            {
                if (BusinessObjectInfo.EventType == BoEventTypes.et_FORM_DATA_LOAD && BusinessObjectInfo.BeforeAction == false)
                {
                    formDataLoad(oForm, out errorText);
                }

                if (BusinessObjectInfo.EventType == BoEventTypes.et_FORM_DATA_ADD && BusinessObjectInfo.BeforeAction == false && BusinessObjectInfo.ActionSuccess)
                {
                    if (Program.canceledDocEntry != 0)
                    {
                        cancellation(oForm, Program.canceledDocEntry, out errorText);
                        Program.canceledDocEntry = 0;
                    }
                }

                if (BusinessObjectInfo.EventType == BoEventTypes.et_FORM_DATA_ADD)
                {
                    if (BusinessObjectInfo.BeforeAction)
                    {
                        DBDataSource DocDBSource = oForm.DataSources.DBDataSources.Item(0);
                        bool rejection = false;
                        if (DocDBSource.GetValue("CANCELED", 0) == "N")
                        {
                            //უარყოფითი ნაშთების კონტროლი დოკ.თარიღით
                            CommonFunctions.blockNegativeStockByDocDate(oForm, "OINV", "INV1", "WhsCode", out rejection);
                            if (rejection)
                            {
                                Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("DocumentCannotBeAdded"));
                                BubbleEvent = false;
                            }
                        }

                        //ძირითადი საშუალებების შემოწმება
                        if (BatchNumberSelection.SelectedBatches != null)
                        {
                            bool rejectionAsset = false;
                            CommonFunctions.blockAssetInvoice(oForm, "OINV", out rejectionAsset);
                            if (rejectionAsset)
                            {
                                Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("DocumentCannotBeAdded"));
                                BubbleEvent = false;
                            }
                        }
                    }

                    if (BusinessObjectInfo.ActionSuccess != BusinessObjectInfo.BeforeAction)
                    {
                        //დოკუმენტის გატარების დროს გატარდეს ბუღლტრული გატარება
                        DBDataSource DocDBSource = oForm.DataSources.DBDataSources.Item(0);

                        if (DocDBSource.GetValue("CANCELED", 0) == "N")
                        {
                            string DocEntry = DocDBSource.GetValue("DocEntry", 0);
                            DocEntry = DocEntry == "" ? "0" : DocEntry;

                            string DocCurrency = DocDBSource.GetValue("DocCur", 0);
                            decimal DocRate = FormsB1.cleanStringOfNonDigits(DocDBSource.GetValue("DocRate", 0));
                            string DocNum = DocDBSource.GetValue("DocNum", 0);
                            DateTime DocDate = DateTime.ParseExact(DocDBSource.GetValue("DocDate", 0), "yyyyMMdd", CultureInfo.InvariantCulture);

                            CommonFunctions.StartTransaction();

                            Program.JrnLinesGlobal = new DataTable();
                            DataTable JrnLinesDT = createAdditionalEntries(oForm, null, null, DocCurrency, DocEntry, DocRate);

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

                    //Use Rate Ranges Update
                    if ((BusinessObjectInfo.EventType == BoEventTypes.et_FORM_DATA_ADD || BusinessObjectInfo.EventType == BoEventTypes.et_FORM_DATA_UPDATE)
                                    && BusinessObjectInfo.ActionSuccess && BusinessObjectInfo.BeforeAction == false)
                    {
                        CommonFunctions.StartTransaction();

                        DBDataSource DocDBSource = oForm.DataSources.DBDataSources.Item(0);
                        string DocEntry = DocDBSource.GetValue("DocEntry", 0);
                        string ObjType = DocDBSource.GetValue("ObjType", 0);
                        string UseRateRanges = DocDBSource.GetValue("U_UseBlaAgRt", 0);

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
        }

        public static DataTable createAdditionalEntries(Form oForm, SAPbobsCOM.GeneralData oGeneralData, DataTable DTSource, string DocCurrency, string DocEntry, decimal DocRate)
        {
            DataTable jeLines = JournalEntry.JournalEntryTable();
            DocCurrency = DocCurrency == CommonFunctions.getLocalCurrency() ? "" : DocCurrency;
            DataTable AccountTable = CommonFunctions.GetOACTTable();


            //            string DocEntry = DocDBSource.GetValue("DocEntry", 0);
            //            string DocNum = DocDBSource.GetValue("DocNum", 0);
            //            string DocCurr = DocDBSource.GetValue("DocCur", 0);
            //            decimal DocRate = FormsB1.cleanStringOfNonDigits( DocDBSource.GetValue("DocRate", 0));
            //            DateTime DocDate = DateTime.ParseExact(DocDBSource.GetValue("DocDate", 0), "yyyyMMdd", CultureInfo.InvariantCulture);
            //            JrnEntry( DocEntry, DocNum, DocDate, DocRate, DocCurr, out errorText);
            //            if (errorText != null)
            //            {
            //                Program.uiApp.MessageBox(errorText);
            //                BubbleEvent = false;

            //            }
            //        }
            //    }
            //}
            //}

            return jeLines;

        }

        public static void uiApp_ItemEvent(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            if (pVal.EventType != BoEventTypes.et_FORM_UNLOAD)
            {
                Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (pVal.EventType == BoEventTypes.et_FORM_LOAD && pVal.BeforeAction)
                {
                    createFormItems(oForm, out errorText);
                    formDataLoad(oForm, out errorText);
                    SetVisibility(oForm);
                    oForm.Items.Item("4").Click();
                }

                else if (pVal.EventType == BoEventTypes.et_ITEM_PRESSED)
                {
                    if (pVal.ItemUID == "1" && pVal.BeforeAction)
                    {
                        CommonFunctions.fillDocRate(oForm, "OINV");
                    }

                    if (!pVal.BeforeAction)
                    {
                        if (pVal.ItemUID == "UsBlaAgRtS")
                        {
                            EditText oBlankAgr = (EditText)oForm.Items.Item("1980002192").Specific;

                            if (string.IsNullOrEmpty(oBlankAgr.Value))
                            {
                                Program.uiApp.SetStatusBarMessage(errorText = BDOSResources.getTranslate("EmptyBlaAgrError"), BoMessageTime.bmt_Short);
                                CheckBox oUsBlaAgRtCB = (CheckBox)oForm.Items.Item("UsBlaAgRtS").Specific;
                                oUsBlaAgRtCB.Checked = false;
                                oForm.Items.Item("1980002192").Click();
                            }
                        }
                        else if (pVal.ItemUID == "BDO_WblTxt" || pVal.ItemUID == "BDO_TaxTxt")
                        {
                            int newDocEntry = 0;
                            string bstrUDOObjectType = null;

                            itemPressed(oForm, pVal, out newDocEntry, out bstrUDOObjectType);
                            oForm.Update();

                            if (newDocEntry != 0 && bstrUDOObjectType != null)
                            {
                                Program.uiApp.OpenForm(BoFormObjectEnum.fo_UserDefinedObject, bstrUDOObjectType, newDocEntry.ToString());
                            }
                        }
                    }
                }

                else if (pVal.EventType == BoEventTypes.et_VALIDATE && !pVal.BeforeAction && pVal.ItemChanged)
                {
                    if (oForm.Items.Item("DiscountE").Visible)
                    {
                        if (pVal.ItemUID == "38" && (pVal.ColUID == "14" || (pVal.ColUID == "15" && !pVal.InnerEvent)))
                        {
                            SetInitialItemGrossPrices(oForm, pVal.ColUID, pVal.Row);
                            ApplyDiscount(oForm);
                        }

                        else if (((pVal.ItemUID == "38" && pVal.ColUID == "11") || pVal.ItemUID == "DiscountE") && !pVal.InnerEvent)
                        {
                            ApplyDiscount(oForm);
                        }
                    }
                }
            }
        }

        public static void JrnEntry(string DocEntry, string DocNum, DateTime DocDate, DataTable JrnLinesDT, out string errorText)
        {
            errorText = null;

            try
            {
                JournalEntry.JrnEntry(DocEntry, "13", "AR Invoice: " + DocNum, DocDate, JrnLinesDT, out errorText);

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

        public static List<int> getAllConnectedDoc(List<int> docEntry, string baseType, DateTime docDate, DateTime docDateCorr, int docTime, out string errorText)
        {
            errorText = null;
            List<int> connectedDocList = new List<int>();

            DateTime firstDayOfMonth = new DateTime(docDate.Year, docDate.Month, 1);
            DateTime lastDayOfMonth = firstDayOfMonth.AddMonths(1).AddDays(-1);

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string query = @"SELECT
	             DISTINCT ""RIN1"".""DocEntry"",
	             ""RIN1"".""DocDate"",
	             ""RIN1"".""BaseType"",
	             ""RIN1"".""BaseEntry"",
	             ""ORIN"".""U_BDO_CNTp"",
                 ""ORIN"".""DocTime"",
	             ""ORIN"".""ObjType"" 
            FROM ""RIN1"" 
            INNER JOIN ""ORIN"" AS ""ORIN"" ON ""RIN1"".""DocEntry"" = ""ORIN"".""DocEntry"" 
            WHERE ""RIN1"".""BaseEntry"" IN (" + string.Join(",", docEntry) + @")            
            AND ""ORIN"".""CANCELED"" = 'N' 
            AND ""RIN1"".""BaseType"" = '" + baseType + @"'";

            if (docDate != new DateTime())
            {
                query = query + @" AND ""RIN1"".""DocDate"" >= '" + firstDayOfMonth.ToString("yyyyMMdd") + @"'  
                                   AND ""RIN1"".""DocDate"" <= '" + lastDayOfMonth.ToString("yyyyMMdd") + @"'";
            }
            if (docDateCorr != new DateTime())
            {
                query = query + @" AND ""RIN1"".""DocDate"" <= '" + docDateCorr.ToString("yyyyMMdd") + @"'";
                query = query + @" AND ""ORIN"".""DocTime"" <= '" + docTime + @"'";
            }

            query = query + @" ORDER BY ""RIN1"".""DocDate"", ""ORIN"".""DocTime"", ""RIN1"".""DocEntry"" ASC";

            try
            {
                oRecordSet.DoQuery(query);
                if (oRecordSet.RecordCount > 0)
                {
                    while (!oRecordSet.EoF)
                    {
                        connectedDocList.Add(Convert.ToInt32(oRecordSet.Fields.Item("DocEntry").Value));
                        oRecordSet.MoveNext();
                    }
                }
                return connectedDocList;
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
                return connectedDocList;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                oRecordSet = null;
                GC.Collect();
            }
        }

        public static List<int> getAllConnectedARCorrectionDoc(List<int> docEntry, string baseType, DateTime docDate, DateTime docDateCorr, int docTime, out string errorText)
        {
            errorText = null;
            List<int> connectedDocList = new List<int>();

            DateTime firstDayOfMonth = new DateTime(docDate.Year, docDate.Month, 1);
            DateTime lastDayOfMonth = firstDayOfMonth.AddMonths(1).AddDays(-1);

            SAPbobsCOM.Recordset oRecordSet = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string query = @"SELECT
	             DISTINCT ""CSI1"".""DocEntry"",
	             ""CSI1"".""DocDate"",
	             ""CSI1"".""BaseType"",
	             ""CSI1"".""BaseEntry"",
	             ""OCSI"".""U_BDOSCITp"",
                 ""OCSI"".""DocTime"",
	             ""OCSI"".""ObjType"" 
            FROM ""CSI1"" 
            INNER JOIN ""OCSI"" AS ""OCSI"" ON ""CSI1"".""DocEntry"" = ""OCSI"".""DocEntry"" 
            WHERE ""CSI1"".""BaseEntry"" IN (" + string.Join(",", docEntry) + @")            
            AND ""OCSI"".""CANCELED"" = 'N' 
            AND ""CSI1"".""BaseType"" = '" + baseType + @"'";

            if (docDate != new DateTime())
            {
                query = query + @" AND ""CSI1"".""DocDate"" >= '" + firstDayOfMonth.ToString("yyyyMMdd") + @"'  
                                   AND ""CSI1"".""DocDate"" <= '" + lastDayOfMonth.ToString("yyyyMMdd") + @"'";
            }
            if (docDateCorr != new DateTime())
            {
                query = query + @" AND ""CSI1"".""DocDate"" <= '" + docDateCorr.ToString("yyyyMMdd") + @"'";
                query = query + @" AND ""OCSI"".""DocTime"" <= '" + docTime + @"'";
            }

            query = query + @" ORDER BY ""CSI1"".""DocDate"", ""OCSI"".""DocTime"", ""CSI1"".""DocEntry"" ASC";

            try
            {
                oRecordSet.DoQuery(query);
                if (oRecordSet.RecordCount > 0)
                {
                    while (!oRecordSet.EoF)
                    {
                        connectedDocList.Add(Convert.ToInt32(oRecordSet.Fields.Item("DocEntry").Value));
                        oRecordSet.MoveNext();
                    }
                }
                return connectedDocList;
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
                return connectedDocList;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
            }
        }

        private static void SetVisibility(Form oForm)
        {
            var isDiscountUsed = CompanyDetails.IsDiscountUsed();
            oForm.Items.Item("24").Visible = !isDiscountUsed;
            oForm.Items.Item("283").Visible = !isDiscountUsed;
            oForm.Items.Item("42").Visible = !isDiscountUsed;
            oForm.Items.Item("DiscountE").Visible = isDiscountUsed;
        }

        private static void SetInitialItemGrossPrices(Form oForm, string column, int row)
        {
            try
            {
                oForm.Freeze(true);

                Matrix oMatrix = oForm.Items.Item("38").Specific;

                if (column == "14")
                {
                    oMatrix.GetCellSpecific("15", row).Value = 0;
                }

                var initialItemGrossPrice =
                    Convert.ToDecimal(FormsB1.cleanStringOfNonDigits(oMatrix.GetCellSpecific("20", row).Value));

                if (initialItemGrossPrice == 0) return;
                InitialItemGrossPrices[row] = initialItemGrossPrice;
            }
            catch (Exception ex)
            {
                Program.uiApp.StatusBar.SetSystemMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        private static void ApplyDiscount(Form oForm)
        {
            try
            {
                oForm.Freeze(true);
                Matrix oMatrix = oForm.Items.Item("38").Specific;
                EditText oEditText = oForm.Items.Item("DiscountE").Specific;

                var discountTotal = string.IsNullOrEmpty(oEditText.Value) ? 0 : Convert.ToDecimal(oEditText.Value);
                
                var grossTotal = 0;

                for (var row = 1; row < oMatrix.RowCount; row++)
                {
                    var itemPrice = oMatrix.GetCellSpecific("14", row).Value;
                    if (!string.IsNullOrEmpty(itemPrice))
                    {
                        var itemQuantity = Convert.ToDecimal(oMatrix.GetCellSpecific("11", row).Value);

                        grossTotal += itemQuantity * InitialItemGrossPrices[row];
                    }
                    else
                    {
                        //Program.uiApp.StatusBar.SetSystemMessage("Fill Item Prices", BoMessageTime.bmt_Short);
                        oEditText.Value = string.Empty;
                        return;
                    }
                }

                for (var row = 1; row < oMatrix.RowCount; row++)
                {
                    var grossItemAmt = InitialItemGrossPrices[row];

                    var discount = discountTotal / grossTotal * grossItemAmt;

                    var grossAfterDiscount = Math.Round(grossItemAmt - discount, 4);

                    oMatrix.GetCellSpecific("20", row).Value =
                        FormsB1.ConvertDecimalToStringForEditboxStrings(grossAfterDiscount);
                }
            }
            catch (Exception ex)
            {
                Program.uiApp.StatusBar.SetSystemMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short);
            }
            finally
            {
                oForm.Freeze(false);
            }

        }
    }
}