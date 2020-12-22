using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.InteropServices;
using SAPbobsCOM;
using SAPbouiCOM;
using static BDO_Localisation_AddOn.BDOSResources;
using static BDO_Localisation_AddOn.FormsB1;
using static BDO_Localisation_AddOn.Program;

namespace BDO_Localisation_AddOn
{
    static class ArCorrectionInvoice
    {
        public static void CreateUserFields(out string errorText)
        {
            #region Correction Invoice Type

            var listValidValues = new List<string> { "Correction", "Return" };
            //0 //კორექტირება
            //1 //დაბრუნება

            var fieldsKeysMap = new Dictionary<string, object>
            {
                {"Name", "BDOSCITp"},
                {"TableName", "OCSI"},
                {"Description", "Correction Invoice Type"},
                {"Type", BoFieldTypes.db_Alpha},
                {"EditSize", 50},
                {"ValidValues", listValidValues}
            };

            UDO.addUserTableFields(fieldsKeysMap, out errorText);

            #endregion
        }

        private static void CreateFormItems(Form oForm, out string errorText)
        {

            #region Waybill

            //<-------------------------------------------სასაქონლო ზედნადები----------------------------------->

            double height = oForm.Items.Item("86").Height;
            double top = oForm.Items.Item("86").Top + height * 1.5 + 1;
            double leftS = oForm.Items.Item("86").Left;
            double leftE = oForm.Items.Item("46").Left;
            double widthE = oForm.Items.Item("46").Width;

            string caption = getTranslate("CreateWaybill");
            var formItems = new Dictionary<string, object>();
            var itemName = "BDO_WblTxt";
            formItems.Add("Type", BoFormItemTypes.it_STATIC);
            formItems.Add("Left", leftS);
            formItems.Add("Width", widthE * 1.5);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", caption);
            formItems.Add("TextStyle", 4);
            formItems.Add("FontSize", 10);
            formItems.Add("Enabled", true);

            createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            string objectType = "UDO_F_BDO_WBLD_D"; //Waybill document
            string uniqueID_WaybillCFL = "Waybill_CFL";
            addChooseFromList(oForm, false, objectType, uniqueID_WaybillCFL);

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
            formItems.Add("Left", leftE + widthE - 40);
            formItems.Add("Width", 40);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("AffectsFormMode", false);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", false);
            formItems.Add("ChooseFromListUID", uniqueID_WaybillCFL);
            formItems.Add("ChooseFromListAlias", "DocEntry");

            createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "BDO_WblLB"; //10 characters
            formItems.Add("Type", BoFormItemTypes.it_LINKED_BUTTON);
            formItems.Add("Left", leftE + widthE - 40 - 20);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("LinkTo", "BDO_WblDoc");
            formItems.Add("LinkedObjectType", objectType);

            createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            oForm.DataSources.UserDataSources.Add("BDO_WblID", BoDataType.dt_SHORT_TEXT, 20);
            oForm.DataSources.UserDataSources.Add("BDO_WblNum", BoDataType.dt_SHORT_TEXT, 20);
            oForm.DataSources.UserDataSources.Add("BDO_WblSts", BoDataType.dt_SHORT_TEXT, 50);

            //<-------------------------------------------სასაქონლო ზედნადები-----------------------------------

            #endregion

            #region TaxInvoice

            //-------------------------------------------ანგარიშ-ფაქტურა----------------------------------->
            top = top + height * 1.5 + 1;

            formItems = new Dictionary<string, object>();
            itemName = "BDO_TaxTxt"; //10 characters
            formItems.Add("Type", BoFormItemTypes.it_STATIC);
            formItems.Add("Left", leftS);
            formItems.Add("Width", widthE * 1.5);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", getTranslate("CreateTaxInvoice"));
            formItems.Add("TextStyle", 4);
            formItems.Add("FontSize", 10);
            formItems.Add("Enabled", true);

            createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            objectType = "UDO_F_BDO_TAXS_D"; //Tax invoice sent document
            const string uniqueID_TaxInvoiceSentCFL = "TaxInvoiceSent_CFL";
            addChooseFromList(oForm, false, objectType, uniqueID_TaxInvoiceSentCFL);

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
            formItems.Add("Left", leftE + widthE - 40);
            formItems.Add("Width", 40);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("AffectsFormMode", false);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", false);
            formItems.Add("ChooseFromListUID", uniqueID_TaxInvoiceSentCFL);
            formItems.Add("ChooseFromListAlias", "DocEntry");

            createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "BDO_TaxLB"; //10 characters
            formItems.Add("Type", BoFormItemTypes.it_LINKED_BUTTON);
            formItems.Add("Left", leftE + widthE - 40 - 20);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("LinkTo", "BDO_TaxDoc");
            formItems.Add("LinkedObjectType", objectType);

            createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            oForm.DataSources.UserDataSources.Add("BDO_TaxSer", BoDataType.dt_SHORT_TEXT, 20);
            oForm.DataSources.UserDataSources.Add("BDO_TaxNum", BoDataType.dt_SHORT_TEXT, 20);
            oForm.DataSources.UserDataSources.Add("BDO_TaxDat", BoDataType.dt_DATE, 20);

            //<-------------------------------------------ანგარიშ-ფაქტურა-----------------------------------

            #endregion

            #region OpType

            //--------------------------------------------ოპერაციის ტიპი-----------------------------------------

            top = oForm.Items.Item("10001018").Top + height + 1;
            leftS = oForm.Items.Item("10001018").Left;
            leftE = oForm.Items.Item("10001019").Left;
            int widthS = oForm.Items.Item("10001018").Width;
            widthE = oForm.Items.Item("10001019").Width;

            formItems = new Dictionary<string, object>();
            itemName = "BDOSTpSt";
            formItems.Add("Type", BoFormItemTypes.it_STATIC);
            formItems.Add("Left", leftS);
            formItems.Add("Width", widthS);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", getTranslate("OperationType"));

            createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            List<string> listValidValues = new List<string> { getTranslate("Correction"), getTranslate("Return") };
            //0 //კორექტირება
            //1 //დაბრუნება

            formItems = new Dictionary<string, object>();
            itemName = "BDOSCITp";
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OCSI");
            formItems.Add("Alias", "U_BDOSCITp");
            formItems.Add("Bound", true);
            formItems.Add("Type", BoFormItemTypes.it_COMBO_BOX);
            formItems.Add("Left", leftE);
            formItems.Add("Width", widthE);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("ExpandType", BoExpandType.et_DescriptionOnly);
            formItems.Add("DisplayDesc", true);
            formItems.Add("ValidValues", listValidValues);

            createFormItem(oForm, formItems, out errorText);

            //--------------------------------------------ოპერაციის ტიპი-----------------------------------------

            #endregion
        }

        public static void UiApp_FormDataEvent(ref BusinessObjectInfo businessObjectInfo, out bool bubbleEvent)
        {
            bubbleEvent = true;

            Form oForm = uiApp.Forms.GetForm(businessObjectInfo.FormTypeEx, currentFormCount);

            if (oForm.TypeEx != "70008") return;

            if (businessObjectInfo.EventType == BoEventTypes.et_FORM_DATA_LOAD &
                !businessObjectInfo.BeforeAction)
            {
                FormDataLoad(oForm, out _);
            }

            else if (businessObjectInfo.EventType == BoEventTypes.et_FORM_DATA_ADD &
                     !businessObjectInfo.BeforeAction & businessObjectInfo.ActionSuccess)
            {
                if (canceledDocEntry == 0) return;
                Cancellation(oForm, canceledDocEntry, out _);
                canceledDocEntry = 0;
            }
        }

        public static void UiApp_ItemEvent(ref ItemEvent pVal, out bool bubbleEvent)
        {
            bubbleEvent = true;

            if (pVal.EventType == BoEventTypes.et_FORM_UNLOAD) return;

            Form oForm = uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

            if (pVal.EventType == BoEventTypes.et_FORM_LOAD)
            {
                if (pVal.BeforeAction)
                {
                    CreateFormItems(oForm, out _);
                    FormDataLoad(oForm, out _);
                    FormsB1.WB_TAX_AuthorizationsItems(oForm);
                }
                else
                {
                    SetValues(oForm, out _);
                }
            }

            else if (pVal.EventType == BoEventTypes.et_ITEM_PRESSED &&
                     (pVal.ItemUID == "BDO_WblTxt" || pVal.ItemUID == "BDO_TaxTxt") & !pVal.BeforeAction)
            {
                oForm.Freeze(true);

                ItemPressed(oForm, pVal, out var newDocEntry, out var bstrUdoObjectType, out var errorText);

                if (errorText != null)
                {
                    uiApp.MessageBox(errorText);
                }

                oForm.Freeze(false);
                oForm.Update();

                if (newDocEntry != 0 && bstrUdoObjectType != null)
                {
                    uiApp.OpenForm(BoFormObjectEnum.fo_UserDefinedObject, bstrUdoObjectType, newDocEntry.ToString());
                }
            }

            if (pVal.EventType == BoEventTypes.et_COMBO_SELECT & pVal.ItemUID == "BDOSCITp" & !pVal.BeforeAction)
            {
                FormDataLoad(oForm, out _);
            }
        }

        public static void FormDataLoad(Form oForm, out string errorText)
        {
            errorText = null;

            StaticText oStaticText;
            oForm.Freeze(true);
            try
            {
                int docEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("OCSI").GetValue("DocEntry", 0));

                #region Waybill

                string caption = getTranslate("CreateWaybill");
                int wblDocEntry;
                string wblId = "";
                string wblNum = "";
                string wblSts = "";
                string objType;

                string oCITp = oForm.DataSources.DBDataSources.Item("OCSI").GetValue("U_BDOSCITp", 0).Trim();

                if (oCITp == "0")
                {
                    GetBaseDoc(docEntry, out int oBaseDocEntry);
                    if (oBaseDocEntry == 0)
                    {
                        return;
                    }

                    docEntry = oBaseDocEntry;
                    objType = "13";
                }

                else
                {
                    objType = "165";
                }

                if (docEntry != 0)
                {
                    Dictionary<string, string> wblDocInfo =
                        BDO_Waybills.getWaybillDocumentInfo(docEntry, objType, out errorText);
                    wblDocEntry = Convert.ToInt32(wblDocInfo["DocEntry"]);
                    wblId = wblDocInfo["wblID"];
                    wblNum = wblDocInfo["number"];
                    wblSts = wblDocInfo["status"];

                    if (wblDocEntry != 0)
                    {
                        caption = getTranslate("Wb") + ": " + wblSts + " " + wblId +
                                  (wblNum != "" ? " № " + wblNum : "");
                    }
                }
                else
                {
                    caption = getTranslate("CreateWaybill");
                    wblDocEntry = 0;
                }

                oForm.DataSources.UserDataSources.Item("BDO_WblDoc").ValueEx =
                    wblDocEntry == 0 ? "" : wblDocEntry.ToString();
                oForm.DataSources.UserDataSources.Item("BDO_WblID").ValueEx = wblId;
                oForm.DataSources.UserDataSources.Item("BDO_WblNum").ValueEx = wblNum;
                oForm.DataSources.UserDataSources.Item("BDO_WblSts").ValueEx = wblSts;

                oStaticText = oForm.Items.Item("BDO_WblTxt").Specific;
                oStaticText.Caption = caption;

                #endregion

                #region Tax Invoice

                //-------------------------------------------ანგარიშ-ფაქტურა----------------------------------->
                string cardCode = oForm.DataSources.DBDataSources.Item("OCSI").GetValue("CardCode", 0).Trim();
                docEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("OCSI").GetValue("DocEntry", 0));
                caption = getTranslate("CreateTaxInvoice");
                int taxDocEntry = 0;
                string taxNumber = "";
                string taxSeries = "";
                string taxCreateDate = "";

                if (docEntry != 0)
                {
                    Dictionary<string, object> taxDocInfo = BDO_TaxInvoiceSent.getTaxInvoiceSentDocumentInfo(docEntry, "ARCorrectionInvoice", cardCode);
                    if (taxDocInfo != null)
                    {
                        taxDocEntry = Convert.ToInt32(taxDocInfo["docEntry"]);
                        taxNumber = taxDocInfo["number"].ToString();
                        taxSeries = taxDocInfo["series"].ToString();
                        taxCreateDate = taxDocInfo["createDate"].ToString();

                        if (taxDocEntry != 0)
                        {
                            DateTime taxCreateDateDt = DateTime.ParseExact(taxCreateDate, "yyyyMMdd", CultureInfo.InvariantCulture);
                            if (taxSeries == "")
                            {
                                caption = getTranslate("TaxInvoiceDate") + " " + taxCreateDateDt;
                            }
                            else
                            {
                                caption = getTranslate("TaxInvoiceSeries") + " " + taxSeries + " № " + taxNumber + " " + getTranslate("Data") + " " + taxCreateDateDt;
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

                oStaticText = (StaticText)oForm.Items.Item("BDO_TaxTxt").Specific;
                oStaticText.Caption = caption;
                //<-------------------------------------------ანგარიშ-ფაქტურა-----------------------------------

                #endregion

            }

            catch (Exception ex)
            {
                oForm.DataSources.UserDataSources.Item("BDO_WblDoc").ValueEx = "";
                oForm.DataSources.UserDataSources.Item("BDO_WblID").ValueEx = "";
                oForm.DataSources.UserDataSources.Item("BDO_WblNum").ValueEx = "";
                oForm.DataSources.UserDataSources.Item("BDO_WblSts").ValueEx = "";

                oStaticText = oForm.Items.Item("BDO_WblTxt").Specific;
                oStaticText.Caption = getTranslate("CreateWaybill");

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
            }
        }

        public static void GetBaseDoc(int docEntry, out int baseEntry)
        {
            baseEntry = 0;

            var oRecordSet = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            try
            {
                string query = "SELECT DISTINCT " +
                               "\"CSI1\".\"BaseEntry\" AS \"BaseEntry\" " +
                               "FROM \"CSI1\" " +
                               "WHERE \"CSI1\".\"DocEntry\" = '" + docEntry + "' AND \"CSI1\".\"BaseType\" = '13'";
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
            }
            catch
            {
                // ignored
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
            }
        }

        public static void GetAmount(int docEntry, out double gTotal, out double lineVat, out string errorText)
        {
            errorText = null;
            gTotal = 0;
            lineVat = 0;

            Recordset oRecordSet = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            string query = @"SELECT 
            ""CSI1"".""DocEntry"" AS ""docEntry"", 
            SUM(""CSI1"".""GTotal"") AS ""GTotal"", 
            SUM(""CSI1"".""LineVat"") AS ""LineVat"" 
            FROM ""CSI1"" AS ""CSI1"" 
            WHERE ""CSI1"".""DocEntry"" = '" + docEntry + @"' 
            GROUP BY ""CSI1"".""DocEntry""";

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
            }
        }

        private static void Cancellation(Form oForm, int docEntry, out string errorText)
        {
            errorText = null;

            try
            {
                Documents oCorrectionInvoice = oCompany.GetBusinessObject(BoObjectTypes.oCorrectionInvoice);

                if (oCorrectionInvoice.GetByKey(docEntry) &&
                    oCorrectionInvoice.UserFields.Fields.Item("U_BDOSCITp").Value == "1")
                {
                    Dictionary<string, string> wblDocInfo =
                        BDO_Waybills.getWaybillDocumentInfo(docEntry, "165", out errorText);
                    int wblDocEntry = Convert.ToInt32(wblDocInfo["DocEntry"]);

                    if (wblDocEntry != 0)
                    {
                        int answer = uiApp.MessageBox(getTranslate("DocumentLinkedToWaybillCancel"), 1,
                            getTranslate("Yes"), getTranslate("No"));
                        string operation = answer == 1 ? "Update" : "Cancel";
                        BDO_Waybills.cancellation(wblDocEntry, operation, out errorText);
                    }
                }

                JournalEntry.cancellation(oForm, docEntry, "13", out errorText);
            }

            catch (Exception ex)
            {
                errorText = ex.Message;
            }
        }

        private static void ItemPressed(Form oForm, ItemEvent pVal, out int newDocEntry, out string bstrUdoObjectType,
            out string errorText)
        {
            errorText = null;
            newDocEntry = 0;
            bstrUdoObjectType = null;

            string docEntryStr = oForm.DataSources.DBDataSources.Item("OCSI").GetValue("DocEntry", 0);
            docEntryStr = string.IsNullOrEmpty(docEntryStr) ? "0" : docEntryStr;
            int docEntry = Convert.ToInt32(docEntryStr);
            string cancelled = oForm.DataSources.DBDataSources.Item("OCSI").GetValue("CANCELED", 0).Trim();
            string docType = oForm.DataSources.DBDataSources.Item("OCSI").GetValue("DocType", 0).Trim();
            string cNTp = oForm.DataSources.DBDataSources.Item("OCSI").GetValue("U_BDOSCITp", 0).Trim();

            switch (pVal.ItemUID)
            {
                case "BDO_WblTxt":
                    {
                        if (docEntry != 0 & (oForm.Mode == BoFormMode.fm_OK_MODE || oForm.Mode == BoFormMode.fm_VIEW_MODE))
                        {
                            string wblDoc = oForm.DataSources.UserDataSources.Item("BDO_WblDoc").ValueEx;
                            bstrUdoObjectType = "UDO_F_BDO_WBLD_D";

                            if (wblDoc == "" && cancelled == "N" && docType == "I" && cNTp == "1")
                            {
                                BDO_Waybills.createDocument("165", docEntry, null, null, null, null, out newDocEntry,
                                    out errorText);

                                if (!(errorText == null & newDocEntry != 0)) return;

                                oForm.DataSources.UserDataSources.Item("BDO_WblDoc").ValueEx = newDocEntry.ToString();
                                FormDataLoad(oForm, out errorText);
                            }

                            else if (cancelled != "N")
                            {
                                errorText = getTranslate("DocumentMustNotBeCancelledOrCancellation");
                            }

                            else if (docType != "I")
                            {
                                errorText = getTranslate("DocumentTypeMustBeItem");
                            }

                            else if (cNTp != "1")
                            {
                                errorText = getTranslate("CreateWaybillAllowedOnlyForReturnType");
                            }
                            else
                            {
                                errorText = getTranslate("ToCreateWaybillWriteDocument");
                            }
                        }

                        break;
                    }
                case "BDO_TaxTxt":
                    {
                        string taxDoc = oForm.DataSources.UserDataSources.Item("BDO_TaxDoc").ValueEx;
                        bstrUdoObjectType = "UDO_F_BDO_TAXS_D";

                        if (docEntry != 0 && (oForm.Mode == BoFormMode.fm_OK_MODE ||
                                              oForm.Mode == BoFormMode.fm_VIEW_MODE))
                        {
                            if (taxDoc == "" && cancelled == "N")
                            {
                                BDO_TaxInvoiceSent.createDocument("165", docEntry, "", true, 0, null, false, null, null,
                                    out newDocEntry, out errorText);

                                if (!string.IsNullOrEmpty(errorText) || newDocEntry == 0) return;

                                oForm.DataSources.UserDataSources.Item("BDO_TaxDoc").ValueEx = newDocEntry.ToString();
                                FormDataLoad(oForm, out errorText);
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

                        break;
                    }
            }
        }

        private static void SetValues(Form oForm, out string errorText)
        {
            errorText = null;
            try
            {
                string docEntry = oForm.DataSources.DBDataSources.Item("OCSI").GetValue("DocEntry", 0).Trim();

                if (!string.IsNullOrEmpty(docEntry))
                {
                    return;
                }

                ComboBox oCombo = (ComboBox)oForm.Items.Item("BDOSCITp").Specific;
                oCombo.Select("0");
            }

            catch (Exception ex)
            {
                errorText = ex.Message;
            }
        }

        public static List<int> getAllConnectedDoc(List<int> docEntry, string baseType)
        {
            List<int> connectedDocList = new List<int>();

            Recordset oRecordSet = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            string query = @"SELECT
            	 ""CSI1"".""DocEntry"" 
            FROM ""CSI1"" 
            WHERE ""CSI1"".""BaseEntry"" IN (" + string.Join(",", docEntry) + @") 
            AND ""CSI1"".""BaseType"" = '" + baseType + @"'
            GROUP BY ""CSI1"".""DocEntry""";

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
            }
        }
    }
}
