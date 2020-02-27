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
    static partial class ARCreditNote
    {
        public static bool ReserveInvoiceAsService = false;

        public static void createUserFields(out string errorText)
        {
            errorText = null;
            List<string> listValidValues;
            Dictionary<string, object> fieldskeysMap;
            listValidValues = new List<string>();
            listValidValues.Add("Correction"); //0 //კორექტირება
            listValidValues.Add("Return"); //1 //დაბრუნება

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDO_CNTp");
            fieldskeysMap.Add("TableName", "ORIN");
            fieldskeysMap.Add("Description", "CreditNote Type");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);
            fieldskeysMap.Add("ValidValues", listValidValues);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            //მომსახურების აღწერა
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSSrvDsc");
            fieldskeysMap.Add("TableName", "ORIN");
            fieldskeysMap.Add("Description", "Service Description");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 150);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            GC.Collect();
        }

        public static void createFormItems(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            Dictionary<string, object> formItems = null;

            string itemName = "";

            SAPbouiCOM.Item oItem = oForm.Items.Item("15");
            double height = oItem.Height;
            double top = oForm.Items.Item("70").Top + height * 2 + 1;
            double left_s = oItem.Left;
            double left_e = oForm.Items.Item("14").Left;
            double width = oItem.Width;

            formItems = new Dictionary<string, object>();
            itemName = "BDO_TpSt";
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width);
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
            formItems.Add("TableName", "ORIN");
            formItems.Add("Alias", "U_BDO_CNTp");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width);
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

            //<-------------------------------------------სასაქონლო ზედნადები----------------------------------->
            height = oForm.Items.Item("86").Height;
            top = oForm.Items.Item("86").Top + height * 1.5 + 1;
            left_s = oForm.Items.Item("86").Left;
            left_e = oForm.Items.Item("46").Left;
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
            FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_WaybillCFL);

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
            FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_TaxInvoiceSentCFL);

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

            //top = top + height + 1;

            oForm.DataSources.UserDataSources.Add("BDO_TaxSer", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
            oForm.DataSources.UserDataSources.Add("BDO_TaxNum", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
            oForm.DataSources.UserDataSources.Add("BDO_TaxDat", SAPbouiCOM.BoDataType.dt_DATE, 20);
            //<-------------------------------------------ანგარიშ-ფაქტურა-----------------------------------

            ReserveInvoiceAsService = (CommonFunctions.getOADM("U_BDOSResSrv").ToString() == "Y");

            if (ReserveInvoiceAsService)
            {

                top = top + height * 1.5 + 1;

                formItems = new Dictionary<string, object>();
                itemName = "SrvDscSt"; //10 characters
                formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                formItems.Add("Left", left_s);
                formItems.Add("Width", width_e * 1.5);
                formItems.Add("Top", top);
                formItems.Add("Height", height);
                formItems.Add("UID", itemName);
                formItems.Add("Caption", BDOSResources.getTranslate("Description"));
                formItems.Add("TextStyle", 4);
                formItems.Add("FontSize", 10);
                formItems.Add("Enabled", false);

                FormsB1.createFormItem(oForm, formItems, out errorText);
                if (errorText != null)
                {
                    return;
                }

                formItems = new Dictionary<string, object>();
                itemName = "BDOSSrvDsc"; //10 characters

                formItems.Add("isDataSource", true);
                formItems.Add("DataSource", "DBDataSources");
                formItems.Add("TableName", "ORIN");
                formItems.Add("Alias", "U_BDOSSrvDsc");
                formItems.Add("Bound", true);
                formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                formItems.Add("Left", left_e);
                formItems.Add("Width", width_e);
                formItems.Add("Top", top);
                formItems.Add("Height", height);
                formItems.Add("UID", itemName);
                formItems.Add("DisplayDesc", true);
                formItems.Add("Enabled", false);

                FormsB1.createFormItem(oForm, formItems, out errorText);
                if (errorText != null)
                {
                    return;
                }
            }

            GC.Collect();
        }

        public static void setVisibleFormItems(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            SAPbouiCOM.Item oItem = null;

            try
            {
                oForm.Items.Item("BDO_TaxDoc").Enabled = false;
                oForm.Items.Item("BDO_WblDoc").Enabled = false;

                oForm.Items.Item("BDOSSrvDsc").Enabled = false;

                string baseType = oForm.DataSources.DBDataSources.Item("RIN1").GetValue("BaseType", 0).Trim();

                if (baseType == "203") //A/R Down Payment Invoice-ის საფუძველზეა გაფორმებული
                {
                    oItem = oForm.Items.Item("BDO_WblTxt");
                    oItem.Visible = false;
                    oItem = oForm.Items.Item("BDO_WblDoc");
                    oItem.Visible = false;
                    oItem = oForm.Items.Item("BDO_TaxTxt");
                    oItem.Visible = false;
                    oItem = oForm.Items.Item("BDO_TaxDoc");
                    oItem.Visible = false;
                }
                else
                {
                    oItem = oForm.Items.Item("BDO_WblTxt");
                    oItem.Visible = true;
                    oItem = oForm.Items.Item("BDO_WblDoc");
                    oItem.Visible = true;
                    oItem = oForm.Items.Item("BDO_TaxTxt");
                    oItem.Visible = true;
                    oItem = oForm.Items.Item("BDO_TaxDoc");
                    oItem.Visible = true;
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

        public static void formDataLoad(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            SAPbouiCOM.StaticText oStaticText = null;
            oForm.Freeze(true);
            try
            {
                string baseType = oForm.DataSources.DBDataSources.Item("RIN1").GetValue("BaseType", 0).Trim();
                if (baseType == "203") //A/R Down Payment Invoice-ის საფუძველზეა გაფორმებული
                {
                    return;
                }

                int docEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("ORIN").GetValue("DocEntry", 0).Trim());

                int oBaseDocEntry = 0;

                //-------------------------------------------სასაქონლო ზედნადები----------------------------------->
                string caption = BDOSResources.getTranslate("CreateWaybill");
                int wblDocEntry = 0;
                string wblID = "";
                string wblNum = "";
                string wblSts = "";
                string objType = "14";

                string oCNTp = oForm.DataSources.DBDataSources.Item("ORIN").GetValue("U_BDO_CNTp", 0).Trim();

                if (oCNTp == "0")
                {
                    getBaseDoc(docEntry, "13", out oBaseDocEntry);
                    if (oBaseDocEntry == 0)
                    {
                        return;
                    }
                    docEntry = oBaseDocEntry;
                    objType = "13";
                }
                else
                {
                    docEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("ORIN").GetValue("DocEntry", 0));
                    objType = "14";
                }

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

                oStaticText = (SAPbouiCOM.StaticText)oForm.Items.Item("BDO_WblTxt").Specific;
                oStaticText.Caption = caption;
                //<-------------------------------------------სასაქონლო ზედნადები-----------------------------------

                //-------------------------------------------ანგარიშ-ფაქტურა----------------------------------->
                string cardCode = oForm.DataSources.DBDataSources.Item("ORIN").GetValue("CardCode", 0).Trim();
                docEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("ORIN").GetValue("DocEntry", 0));
                caption = BDOSResources.getTranslate("CreateTaxInvoice");
                int taxDocEntry = 0;
                string taxID = "";
                string taxNumber = "";
                string taxSeries = "";
                string taxStatus = "";
                string taxCreateDate = "";

                if (docEntry != 0)
                {
                    Dictionary<string, object> taxDocInfo = BDO_TaxInvoiceSent.getTaxInvoiceSentDocumentInfo(docEntry, "ARCreditNote", cardCode);
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

        public static void getBaseDoc(int docEntry, string baseType, out int baseEntry)
        {
            baseEntry = 0;

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                string query = "SELECT DISTINCT " +
                "\"RIN1\".\"BaseEntry\" AS \"BaseEntry\" " +
                "FROM \"RIN1\" " +
                "WHERE \"RIN1\".\"DocEntry\" = '" + docEntry + "' AND \"RIN1\".\"BaseType\" = '13'";
                oRecordSet.DoQuery(query);

                if (oRecordSet.RecordCount > 1)
                {
                    return;
                }
                while (!oRecordSet.EoF)
                {
                    baseEntry = oRecordSet.Fields.Item("BaseEntry").Value;

                    oRecordSet.MoveNext();
                    break;
                }

                return;
            }
            catch { }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                GC.Collect();
            }
        }

        public static void cancellation(SAPbouiCOM.Form oForm, int docEntry, out string errorText)
        {
            errorText = null;

            try
            {

                SAPbobsCOM.Documents oCreditNotes = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes);
                if (oCreditNotes.GetByKey(docEntry) && oCreditNotes.UserFields.Fields.Item("U_BDO_CNTp").Value == "1")
                {
                    Dictionary<string, string> wblDocInfo = BDO_Waybills.getWaybillDocumentInfo(docEntry, "14", out errorText);
                    int wblDocEntry = Convert.ToInt32(wblDocInfo["DocEntry"]);

                    if (wblDocEntry != 0)
                    {
                        int answer = Program.uiApp.MessageBox(BDOSResources.getTranslate("DocumentLinkedToWaybillCancel"), 1, BDOSResources.getTranslate("Yes"), BDOSResources.getTranslate("No"), "");
                        string operation = answer == 1 ? "Update" : "Cancel";
                        BDO_Waybills.cancellation(wblDocEntry, operation, out errorText);
                    }
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }

            try
            {
                JournalEntry.cancellation(oForm, docEntry, "13", out errorText); //საკითხავია
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

        public static void itemPressed(SAPbouiCOM.Form oForm, SAPbouiCOM.ItemEvent pVal, out int newDocEntry, out string bstrUDOObjectType, out string errorText)
        {
            errorText = null;
            newDocEntry = 0;
            bstrUDOObjectType = null;

            string docEntrySTR = oForm.DataSources.DBDataSources.Item("ORIN").GetValue("DocEntry", 0);
            docEntrySTR = string.IsNullOrEmpty(docEntrySTR) == true ? "0" : docEntrySTR;
            int docEntry = Convert.ToInt32(docEntrySTR);
            string cNTp = oForm.DataSources.DBDataSources.Item("ORIN").GetValue("U_BDO_CNTp", 0).Trim();
            string cancelled = oForm.DataSources.DBDataSources.Item("ORIN").GetValue("CANCELED", 0).Trim();
            string docType = oForm.DataSources.DBDataSources.Item("ORIN").GetValue("DocType", 0).Trim();
            string objectType = "14";

            if (pVal.ItemUID == "BDO_WblTxt")
            {
                string wblDoc = oForm.DataSources.UserDataSources.Item("BDO_WblDoc").ValueEx;
                bstrUDOObjectType = "UDO_F_BDO_WBLD_D";

                if (docEntry != 0 & (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_VIEW_MODE))
                {
                    if (wblDoc == "" && cancelled == "N" && docType == "I" && cNTp == "1")
                    {
                        BDO_Waybills.createDocument(objectType, docEntry, null, null, null, null, out newDocEntry, out errorText);
                        if (errorText == null & newDocEntry != 0)
                        {
                            oForm.DataSources.UserDataSources.Item("BDO_WblDoc").ValueEx = newDocEntry.ToString();
                            formDataLoad(oForm, out errorText);
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
                    else if (cNTp != "1")
                    {
                        errorText = BDOSResources.getTranslate("CreateWaybillAllowedOnlyForReturnType");
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
                    if (taxDoc == "" && cancelled == "N") // && cNTp == "1"
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

        public static string getGoodsQueryForTaxInvoiceSent(int docEntry)
        {
            string ID = "0";
            string query = "SELECT " +
            "'" + ID + "'" + " AS \"ID\", " +
            "\"MNTB\".\"LineNum\" AS \"LineNum\", " +
            "\"MNTB\".\"DocEntry\" AS \"DocEntry\", " +
            "\"MNTB\".\"ItemCode\" AS \"ItemCode\", " +
            "\"OITM\".\"CodeBars\" AS \"CodeBars\", " +
            "\"OITM\".\"SWW\" AS \"AdditionalIdentifier\", " +
            "\"MNTB\".\"Dscription\" AS \"W_NAME\", " +
            "\"BDO_RSUOM\".\"U_RSCode\" AS \"UNIT_ID\", " +
            "\"OITM\".\"InvntryUom\" AS \"UNIT_TXT\", " +
            "\"MNTB\".\"VatPrcnt\" AS \"VAT_TYPE\", " +
            "\"MNTB\".\"VatGroup\"AS \"VatGroup\", " +
            "'0' AS \"A_ID\", " +
            "SUM(\"MNTB\".\"Quantity\") AS \"QUANTITY\", " +
            "SUM(\"MNTB\".\"GTotal\") AS \"AMOUNT\", " +
            "CASE WHEN SUM(\"MNTB\".\"Quantity\") = 0 THEN 0 ELSE SUM(\"MNTB\".\"GTotal\")/SUM(\"MNTB\".\"Quantity\") END AS \"PRICE\", " +
            "SUM(\"MNTB\".\"LineVat\") AS \"LineVat\" " +

            "FROM " +

            "(SELECT " +
            "\"RIN1\".\"DocEntry\", " +
            "\"RIN1\".\"LineNum\", " +
            "\"RIN1\".\"ItemCode\", " +
            "\"RIN1\".\"Dscription\", " +
            "\"RIN1\".\"Quantity\" * (CASE WHEN \"RIN1\".\"NoInvtryMv\" = 'Y' THEN 0 ELSE 1 END) * \"RIN1\".\"NumPerMsr\" AS \"Quantity\", " +
            "\"RIN1\".\"GTotal\" , " +
            "\"RIN1\".\"VatPrcnt\", " +
            "\"RIN1\".\"VatGroup\", " +
            "\"RIN1\".\"LineVat\" " +

            "FROM \"RIN1\" " +

            "INNER JOIN \"ORIN\" " +
            "ON \"ORIN\".\"DocEntry\" = \"RIN1\".\"DocEntry\" " +

            "LEFT JOIN \"OITM\" AS \"OITM\" " +
            "ON \"RIN1\".\"ItemCode\" = \"OITM\".\"ItemCode\" " +

            "WHERE \"RIN1\".\"DocEntry\" = '" + docEntry + "' AND \"RIN1\".\"TargetType\" < 0  AND \"ORIN\".\"U_BDO_CNTp\" = 1) AS \"MNTB\" " +

            "LEFT JOIN \"OITM\" AS \"OITM\" " +
            "ON \"MNTB\".\"ItemCode\" = \"OITM\".\"ItemCode\" " +

            "LEFT JOIN \"OUOM\" AS \"OUOM\" " +
            "ON \"OITM\".\"InvntryUom\" = \"OUOM\".\"UomName\" " +

            "LEFT JOIN \"@BDO_RSUOM\" AS \"BDO_RSUOM\" " +
            "ON \"OUOM\".\"UomEntry\" = \"BDO_RSUOM\".\"U_UomEntry\" " +

            "GROUP BY " +
            "\"MNTB\".\"DocEntry\", " +
            "\"MNTB\".\"LineNum\", " +
            "\"MNTB\".\"ItemCode\", " +
            "\"MNTB\".\"Dscription\", " +
            "\"OITM\".\"CodeBars\", " +
            "\"OITM\".\"SWW\", " +
            "\"BDO_RSUOM\".\"U_RSCode\", " +
            "\"OITM\".\"InvntryUom\", " +
            "\"MNTB\".\"VatPrcnt\", " +
            "\"MNTB\".\"VatGroup\" " +
            "HAVING SUM(\"MNTB\".\"Quantity\") > 0 ";

            return query;
        }

        public static void getAmount(int docEntry, out double gTotal, out double lineVat, out string errorText)
        {
            errorText = null;
            gTotal = 0;
            lineVat = 0;

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string query = @"SELECT 
            ""RIN1"".""DocEntry"" AS ""docEntry"", 
            SUM(""RIN1"".""GTotal"") AS ""GTotal"", 
            SUM(""RIN1"".""LineVat"") AS ""LineVat"" 
            FROM ""RIN1"" AS ""RIN1"" 
            WHERE ""RIN1"".""DocEntry"" = '" + docEntry + @"' 
            GROUP BY ""RIN1"".""DocEntry""";

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

        public static void setValues(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            try
            {
                string docEntry = oForm.DataSources.DBDataSources.Item("ORIN").GetValue("DocEntry", 0).Trim();

                if (string.IsNullOrEmpty(docEntry) == false)
                {
                    return;
                }
                SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)oForm.Items.Item("BDO_CNTp").Specific;
                oCombo.Select("0", SAPbouiCOM.BoSearchKey.psk_ByValue);
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
        }

        public static void canUpdateDocument(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            try
            {
                string oCNTp = oForm.DataSources.DBDataSources.Item("ORIN").GetValue("U_BDO_CNTp", 0).Trim();
                bool WoQ = false;

                if (oCNTp == "1")
                {
                    SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("38").Specific));

                    for (int row = 1; row <= oMatrix.RowCount; row++)
                    {
                        WoQ = oMatrix.GetCellSpecific("1250002121", row).Checked;

                        if (WoQ)
                        {
                            errorText = BDOSResources.getTranslate("PostReturnWithoutQuantNotAllowed");
                            break;
                        }
                    }
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

        public static void uiApp_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(BusinessObjectInfo.FormTypeEx, Program.currentFormCount);

            if (oForm.TypeEx == "179")
            {
                if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD & BusinessObjectInfo.BeforeAction == false)
                {
                    formDataLoad(oForm, out errorText);
                    setVisibleFormItems(oForm, out errorText);
                }

                if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD)
                {
                    if (BusinessObjectInfo.BeforeAction == true)
                    {
                        canUpdateDocument(oForm, out errorText);

                        if (string.IsNullOrEmpty(errorText) == false)
                        {
                            Program.uiApp.MessageBox(errorText);
                            BubbleEvent = false;
                        }
                    }
                    else if (BusinessObjectInfo.BeforeAction == false & BusinessObjectInfo.ActionSuccess == true)
                    {
                        if (Program.canceledDocEntry != 0)
                        {
                            cancellation(oForm, Program.canceledDocEntry, out errorText);
                            Program.canceledDocEntry = 0;
                        }
                    }
                }

                if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE & BusinessObjectInfo.BeforeAction == true)
                {
                    canUpdateDocument(oForm, out errorText);

                    if (errorText != null)
                    {
                        Program.uiApp.MessageBox(errorText);
                        BubbleEvent = false;
                    }
                }

                //if (BusinessObjectInfo.ActionSuccess & BusinessObjectInfo.BeforeAction == false)
                //{
                //    //დოკუმენტის გატარების დროს გატარდეს ბუღლტრული გატარება
                //    SAPbouiCOM.DBDataSource DocDBSource = oForm.DataSources.DBDataSources.Item(0);

                //    string DocEntry = DocDBSource.GetValue("DocEntry", 0);
                //    string DocNum = DocDBSource.GetValue("DocNum", 0);
                //    string DocCurr = DocDBSource.GetValue("DocCur", 0);
                //    decimal DocRate = FormsB1.cleanStringOfNonDigits( DocDBSource.GetValue("DocRate", 0));
                //    DateTime DocDate = DateTime.ParseExact(DocDBSource.GetValue("DocDate", 0), "yyyyMMdd", CultureInfo.InvariantCulture);
                //    JrnEntry( DocEntry, DocNum, DocDate, DocRate, DocCurr, out errorText);
                //    if (errorText != null)
                //    {
                //        Program.uiApp.MessageBox(errorText);
                //        BubbleEvent = false;

                //    }
                //}
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
                    formDataLoad(oForm, out errorText);
                    setVisibleFormItems(oForm, out errorText);
                }
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD & pVal.BeforeAction == false)
                {
                    setValues(oForm, out errorText);
                }
                //if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
                //{
                //    if (pVal.ItemUID == "1" && pVal.BeforeAction == true)
                //    {
                //        CommonFunctions.fillDocRate( oForm, "ORIN");
                //    }
                //}

                if (pVal.ItemUID == "BDO_CNTp" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_COMBO_SELECT & pVal.BeforeAction == false)
                {
                    formDataLoad(oForm, out errorText);
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

        public static void JrnEntry(string DocEntry, string DocNum, DateTime DocDate, Decimal rate, string currency, out string errorText)
        {
            errorText = null;

            try
            {
                DataTable jeLines = JournalEntry.JournalEntryTable();

            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
        }

        public static List<int> getAllConnectedDoc(List<int> docEntry, string baseType)
        {
            List<int> connectedDocList = new List<int>();

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string query = @"SELECT
            	 ""RIN1"".""DocEntry"" 
            FROM ""RIN1"" 
            WHERE ""RIN1"".""BaseEntry"" IN (" + string.Join(",", docEntry) + @") 
            AND ""RIN1"".""BaseType"" = '" + baseType + @"'
            GROUP BY ""RIN1"".""DocEntry""";

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
                throw new Exception(ex.Message);
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
