using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Data;

namespace BDO_Localisation_AddOn
{
    static partial class ARInvoice
    {

        public static void createFormItems(  SAPbouiCOM.Form oForm, out string errorText)
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
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
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
            FormsB1.addChooseFromList( oForm, multiSelection, objectType, uniqueID_WaybillCFL);

            formItems = new Dictionary<string, object>();
            itemName = "BDO_WblDoc"; //10 characters
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
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
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

            oForm.DataSources.UserDataSources.Add("BDO_WblID", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
            oForm.DataSources.UserDataSources.Add("BDO_WblNum", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
            oForm.DataSources.UserDataSources.Add("BDO_WblSts", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
            //<-------------------------------------------სასაქონლო ზედნადები-----------------------------------

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
            FormsB1.addChooseFromList( oForm, multiSelection, objectType, uniqueID_TaxInvoiceSentCFL);

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

            top = top + height + 1;

            oForm.DataSources.UserDataSources.Add("BDO_TaxSer", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
            oForm.DataSources.UserDataSources.Add("BDO_TaxNum", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
            oForm.DataSources.UserDataSources.Add("BDO_TaxDat", SAPbouiCOM.BoDataType.dt_DATE, 20);
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
            formItems.Add("SetAutoManaged", true);
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);
            formItems.Add("Enabled", false);
           

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
            fieldskeysMap.Add("Name", "UseBlaAgRt");
            fieldskeysMap.Add("TableName", "OINV");
            fieldskeysMap.Add("Description", "Use Blanket Agreement Rates");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 1);
            fieldskeysMap.Add("DefaultValue", "N");

            UDO.addUserTableFields(fieldskeysMap, out errorText);
            GC.Collect();
        }
        
        public static void formDataLoad( SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            SAPbouiCOM.StaticText oStaticText = null;
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
                    Dictionary<string, string> wblDocInfo = BDO_Waybills.getWaybillDocumentInfo( docEntry, objType, out errorText);
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
                
                oStaticText = (SAPbouiCOM.StaticText)oForm.Items.Item("BDO_WblTxt").Specific;
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
                    Dictionary<string, object> taxDocInfo = BDO_TaxInvoiceSent.getTaxInvoiceSentDocumentInfo( docEntry, "ARInvoice", cardCode, out errorText);
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



                oStaticText = (SAPbouiCOM.StaticText)oForm.Items.Item("BDO_TaxTxt").Specific;
                oStaticText.Caption = caption;
                //<-------------------------------------------ანგარიშ-ფაქტურა-----------------------------------
            }
            catch (Exception ex)
            {
                oForm.DataSources.UserDataSources.Item("BDO_WblDoc").ValueEx = "";
                oForm.DataSources.UserDataSources.Item("BDO_WblID").ValueEx = "";
                oForm.DataSources.UserDataSources.Item("BDO_WblNum").ValueEx = "";
                oForm.DataSources.UserDataSources.Item("BDO_WblSts").ValueEx = "";
                
                oStaticText = (SAPbouiCOM.StaticText)oForm.Items.Item("BDO_WblTxt").Specific;
                oStaticText.Caption = BDOSResources.getTranslate("CreateWaybill");

                oForm.DataSources.UserDataSources.Item("BDO_TaxDoc").ValueEx = "";
                oForm.DataSources.UserDataSources.Item("BDO_TaxSer").ValueEx = "";
                oForm.DataSources.UserDataSources.Item("BDO_TaxNum").ValueEx = "";
                oForm.DataSources.UserDataSources.Item("BDO_TaxDat").ValueEx = "";
                
                oStaticText = (SAPbouiCOM.StaticText)oForm.Items.Item("BDO_TaxTxt").Specific;
                oStaticText.Caption = BDOSResources.getTranslate("CreateTaxInvoice");

                errorText = ex.Message;
            }
            finally
            {
                oForm.Freeze(false);
                GC.Collect();
            }
        }

        public static void cancellation(  SAPbouiCOM.Form oForm, int docEntry, out string errorText)
        {
            errorText = null;

            try
            {
                Dictionary<string, string> wblDocInfo = BDO_Waybills.getWaybillDocumentInfo( docEntry, "13", out errorText);
                int wblDocEntry = Convert.ToInt32(wblDocInfo["DocEntry"]);

                if (wblDocEntry != 0)
                {
                    int answer = Program.uiApp.MessageBox(BDOSResources.getTranslate("DocumentLinkedToWaybillCancel"), 1, BDOSResources.getTranslate("Yes"), BDOSResources.getTranslate("No"), "");
                    string operation = answer == 1 ? "Update" : "Cancel";
                    BDO_Waybills.cancellation( wblDocEntry, operation, out errorText);
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }

            try
            {
                JournalEntry.cancellation( oForm, docEntry, "13", out errorText);
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

        public static void itemPressed(  SAPbouiCOM.Form oForm, SAPbouiCOM.ItemEvent pVal, out int newDocEntry, out string bstrUDOObjectType, out string errorText)
        {
            errorText = null;
            newDocEntry = 0;
            bstrUDOObjectType = null;

            string docEntrySTR = oForm.DataSources.DBDataSources.Item("OINV").GetValue("DocEntry", 0);
            docEntrySTR = string.IsNullOrEmpty(docEntrySTR) == true ? "0" : docEntrySTR;
            int docEntry = Convert.ToInt32(docEntrySTR);
            string cancelled = oForm.DataSources.DBDataSources.Item("OINV").GetValue("CANCELED", 0).Trim();
            string docType = oForm.DataSources.DBDataSources.Item("OINV").GetValue("DocType", 0).Trim();
            string objectType = "13";

            if (pVal.ItemUID == "BDO_WblTxt")
            {
                string wblDoc = oForm.DataSources.UserDataSources.Item("BDO_WblDoc").ValueEx;
                bstrUDOObjectType = "UDO_F_BDO_WBLD_D";

                if (docEntry != 0 & (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_VIEW_MODE))
                {
                    if (wblDoc == "" && cancelled == "N" && docType == "I")
                    {
                        BDO_Waybills.createDocument( objectType, docEntry, null, null, null, null, out newDocEntry, out errorText);
                        if (errorText == null & newDocEntry != 0)
                        {
                            oForm.DataSources.UserDataSources.Item("BDO_WblDoc").ValueEx = newDocEntry.ToString();
                            formDataLoad( oForm, out errorText);
                            return;
                        }
                    }
                    else if (cancelled != "N")
                    {
                        errorText = BDOSResources.getTranslate("DocumentMustNotBeCancelledOrCancellation");
                    }
                    else if (docType != "I")
                    {
                        errorText = BDOSResources.getTranslate("DocumentTypeMustBeItem");
                    }
                }
                else
                {
                    errorText = BDOSResources.getTranslate("ToCreateWaybillWriteDocument");
                }
            }

            else if (pVal.ItemUID == "BDO_TaxTxt")
            {
                string taxDoc = oForm.DataSources.UserDataSources.Item("BDO_TaxDoc").ValueEx;
                bstrUDOObjectType = "UDO_F_BDO_TAXS_D";

                if (docEntry != 0 && (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_VIEW_MODE))
                {
                    if (taxDoc == "" && cancelled == "N")
                    {
                        BDO_TaxInvoiceSent.createDocument( objectType, docEntry, "", true, 0, null, false, null, out newDocEntry, out errorText);
                        if (string.IsNullOrEmpty(errorText) && newDocEntry != 0)
                        {
                            oForm.DataSources.UserDataSources.Item("BDO_TaxDoc").ValueEx = newDocEntry.ToString();
                            formDataLoad( oForm, out errorText);
                            return;
                        }
                    }
                    else if (cancelled != "N")
                    {
                        errorText = BDOSResources.getTranslate("DocumentMustNotBeCancelledOrCancellation");
                    }
                }
                else
                {
                    errorText = BDOSResources.getTranslate("ToCreateTaxInvoiceWriteDocument");
                }
            }
        }

        public static void getAmount( int docEntry, out double gTotal, out double lineVat, out string errorText)
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

        public static void uiApp_FormDataEvent(  ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(BusinessObjectInfo.FormTypeEx, Program.currentFormCount);

            if (oForm.TypeEx == "133")
            {
                if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD & BusinessObjectInfo.BeforeAction == false)
                {
                    formDataLoad( oForm, out errorText);
                }

                if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD & BusinessObjectInfo.BeforeAction == false & BusinessObjectInfo.ActionSuccess == true)
                {
                    if (Program.canceledDocEntry != 0)
                    {
                        cancellation( oForm, Program.canceledDocEntry, out errorText);
                        Program.canceledDocEntry = 0;
                    }
                }

                if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD)
                {
                    if (BusinessObjectInfo.BeforeAction == true)
                    {
                        SAPbouiCOM.DBDataSource DocDBSource = oForm.DataSources.DBDataSources.Item(0);
                        if (DocDBSource.GetValue("CANCELED", 0) == "N")
                        {
                            //უარყოფითი ნაშთების კონტროლი დოკ.თარიღით
                            bool rejection = false;
                            CommonFunctions.blockNegativeStockByDocDate(oForm, "OINV", "INV1", "WhsCode", out rejection);
                            if (rejection)
                            {
                                Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("DocumentCannotBeAdded"));                                
                                BubbleEvent = false;
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
                                if (BusinessObjectInfo.ActionSuccess == true & BusinessObjectInfo.BeforeAction == false)
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
            }
        }

        public static DataTable createAdditionalEntries(SAPbouiCOM.Form oForm, SAPbobsCOM.GeneralData oGeneralData, DataTable DTSource, string DocCurrency, string DocEntry, decimal DocRate)
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
               // DBDataSourceTable = docDBSources.Item("PCH11");
                //JEcount = DBDataSourceTable.Size;
            }

            SAPbouiCOM.DBDataSource BPDataSourceTable = docDBSources.Item("OCRD");

                   


            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string query = @"select
	                        ""OBVL"".""ItemCode"",
	                        ""OBVL"".""DistNumber"",
                            ""OBVL"".""Quantity"",
	                        ""OITB"".""SaleCostAc"",
		                    ""OITB"".""U_BDOSAccDep"",
                            ""INV1"".""Project"",	                    
                            ""INV1"".""OcrCode"",
                            ""INV1"".""OcrCode2"",
                            ""INV1"".""OcrCode3"",
                            ""INV1"".""OcrCode4"",
                            ""INV1"".""OcrCode5""
                        from ""OBVL"" 
                        inner join ""OITM"" on ""OBVL"".""ItemCode"" = ""OITM"".""ItemCode""
                        inner join ""OITB"" on  ""OITB"".""ItmsGrpCod"" = ""OITM"".""ItmsGrpCod"" and ""OITB"".""U_BDOSFxAs""='Y'
                        inner join ""INV1"" on  ""OBVL"".""DocEntry"" = ""INV1"".""DocEntry""                       
                        where""OBVL"".""DocEntry"" = " + DocEntry + @" 
                        and ""OBVL"".""DocType"" = 13";


            oRecordSet.DoQuery(query);
            if (oRecordSet.RecordCount > 0)
            {
                while (!oRecordSet.EoF)
                {
                    string ItemCode = oRecordSet.Fields.Item("ItemCode").Value;
                    string DistNumber = oRecordSet.Fields.Item("DistNumber").Value;
                    string SaleCostAc = oRecordSet.Fields.Item("SaleCostAc").Value;
                    string U_BDOSAccDep = oRecordSet.Fields.Item("U_BDOSAccDep").Value;
                    string DistrRule1 = oRecordSet.Fields.Item( "OcrCode").Value;
                    string DistrRule2 = oRecordSet.Fields.Item( "OcrCode2").Value;
                    string DistrRule3 = oRecordSet.Fields.Item( "OcrCode3").Value;
                    string DistrRule4 = oRecordSet.Fields.Item( "OcrCode4").Value;
                    string DistrRule5 = oRecordSet.Fields.Item( "OcrCode5").Value;
                    string Project    = oRecordSet.Fields.Item("Project").Value;
                    decimal DeprPrice = BDOSDepreciationAccrualDocument.getDepreciationPriceDistNumber(ItemCode, DistNumber);
                    decimal DeprAmount = DeprPrice * Convert.ToDecimal(oRecordSet.Fields.Item("Quantity").Value,CultureInfo.InvariantCulture);
                    JournalEntry.AddJournalEntryRow(AccountTable, jeLines, "Full", U_BDOSAccDep, SaleCostAc, DeprAmount, 0, DocCurrency, "", DistrRule2, DistrRule3, DistrRule4, DistrRule5, Project, "", "");

                    oRecordSet.MoveNext();
            }
        }

            if (jeLines.Rows.Count > 0)
            {
                jeLines = jeLines.AsEnumerable()
                      .GroupBy(row => new
                      {
                          AccountCode = row.Field<string>("AccountCode"),
                          ShortName = row.Field<string>("ShortName"),
                          ContraAccount = row.Field<string>("ContraAccount"),
                          FCCurrency = row.Field<string>("FCCurrency"),
                          CostingCode = row.Field<string>("CostingCode"),
                          CostingCode2 = row.Field<string>("CostingCode2"),
                          CostingCode3 = row.Field<string>("CostingCode3"),
                          CostingCode4 = row.Field<string>("CostingCode4"),
                          CostingCode5 = row.Field<string>("CostingCode5"),
                          ProjectCode = row.Field<string>("ProjectCode"),
                          VatGroupCode = row.Field<string>("VatGroup"),
                          U_BDOSEmpID = row.Field<string>("U_BDOSEmpID")
                      })
                      .Select(g =>
                      {
                          var row = jeLines.NewRow();
                          row["AccountCode"] = g.Key.AccountCode;
                          row["ShortName"] = g.Key.ShortName;
                          row["ContraAccount"] = g.Key.ContraAccount;
                          row["FCCurrency"] = g.Key.FCCurrency;
                          row["CostingCode"] = g.Key.CostingCode;
                          row["CostingCode2"] = g.Key.CostingCode2;
                          row["CostingCode3"] = g.Key.CostingCode3;
                          row["CostingCode4"] = g.Key.CostingCode4;
                          row["CostingCode5"] = g.Key.CostingCode5;
                          row["ProjectCode"] = g.Key.ProjectCode;
                          row["VatGroup"] = g.Key.VatGroupCode;
                          row["U_BDOSEmpID"] = g.Key.U_BDOSEmpID;


                          row["Credit"] = g.Sum(r => r.Field<double>("Credit"));
                          row["Debit"] = g.Sum(r => r.Field<double>("Debit"));
                          row["FCCredit"] = g.Sum(r => r.Field<double>("FCCredit"));
                          row["FCDebit"] = g.Sum(r => r.Field<double>("FCDebit"));
                          return row;
                      }).CopyToDataTable();
            }




            return jeLines;

        }

        public static void uiApp_ItemEvent(  string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);


                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_VALIDATE & pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "1980002192")
                    {
                        setVisibleFormItems(oForm, out errorText);
                    }
                }

                
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD & pVal.BeforeAction == true)
                {
                    createFormItems( oForm, out errorText);
                    formDataLoad( oForm, out errorText);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
                {
                    if (pVal.ItemUID == "1" && pVal.BeforeAction == true)
                    {
                        CommonFunctions.fillDocRate( oForm, "OINV", "INV11");
                    }
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED & pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "BDO_WblTxt" || pVal.ItemUID == "BDO_TaxTxt")
                    {
                        oForm.Freeze(true);
                        int newDocEntry = 0;
                        string bstrUDOObjectType = null;

                        itemPressed(oForm, pVal, out newDocEntry, out bstrUDOObjectType, out errorText);

                        if (errorText != null)
                        {
                            Program.uiApp.MessageBox(errorText);
                        }

                        oForm.Freeze(false);
                        oForm.Update();

                        if (newDocEntry != 0 && bstrUDOObjectType != null)
                        {
                            Program.uiApp.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_UserDefinedObject, bstrUDOObjectType, newDocEntry.ToString());
                        }
                    }
                }

                
            }
        }

        public static void setVisibleFormItems(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            SAPbouiCOM.Item oItem = null;
            oForm.Freeze(true);

            try
            {
                oItem = oForm.Items.Item("1980002192");
                SAPbouiCOM.EditText oEdit = oItem.Specific;
                oItem = oForm.Items.Item("UsBlaAgRtS");
                if (oEdit.Value != "")
                {
                    oItem.Enabled = true;
                }
                else oItem.Enabled = false;
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
            finally
            {
                GC.Collect();
                oForm.Freeze(false);
                oForm.Update();
            }

        }

        public static void JrnEntry(string DocEntry, string DocNum, DateTime DocDate, DataTable JrnLinesDT,  out string errorText)
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

        public static List<int> getAllConnectedDoc( List<int> docEntry, string baseType, DateTime docDate, DateTime docDateCorr, int docTime, out string errorText)
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
    }
}
