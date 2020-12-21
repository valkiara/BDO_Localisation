using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Data;
using SAPbouiCOM;

namespace BDO_Localisation_AddOn
{
    static partial class ARReserveInvoice
    {
        public static bool ReserveInvoiceAsService = false;
        public static string WBAUT = null;
        public static string TXAUT = null;

        private static Dictionary<int, decimal> InitialLineNetTotals = new Dictionary<int, decimal>();

        public static void createUserFields( out string errorText)
        {
            errorText = null;
            Dictionary<string, object> fieldskeysMap;

            //მომსახურების აღწერა
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSSrvDsc");
            fieldskeysMap.Add("TableName", "OINV");
            fieldskeysMap.Add("Description", "Service Description");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 150);

            UDO.addUserTableFields( fieldskeysMap, out errorText);
        }
        

        public static void createFormItems(  SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            Dictionary<string, object> formItems = null;

            string errorTextWB = null;
            Dictionary<string, string> rsSettings = CompanyDetails.getRSSettings(out errorTextWB);
            if (errorTextWB != null)
            {
                WBAUT = "2";
                TXAUT = "2";
            }
            else
            {
                WBAUT = rsSettings["WBAUT"];
                TXAUT = rsSettings["TXAUT"];
            }

            string itemName = "";

            double height = oForm.Items.Item("86").Height;
            double top = oForm.Items.Item("86").Top + height * 1.5 + 1;
            double left_s = oForm.Items.Item("86").Left;
            double left_e = oForm.Items.Item("46").Left;
            double width_e = oForm.Items.Item("46").Width;

            bool multiSelection = false;
            string objectType = "";
                       
            //-------------------------------------------ანგარიშ-ფაქტურა----------------------------------->
            
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

            //top = top + height + 1;

            oForm.DataSources.UserDataSources.Add("BDO_TaxSer", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
            oForm.DataSources.UserDataSources.Add("BDO_TaxNum", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
            oForm.DataSources.UserDataSources.Add("BDO_TaxDat", SAPbouiCOM.BoDataType.dt_DATE, 20);
            //<-------------------------------------------ანგარიშ-ფაქტურა-----------------------------------

            ReserveInvoiceAsService = (CommonFunctions.getOADM( "U_BDOSResSrv").ToString() == "Y");

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
                formItems.Add("TableName", "OINV");
                formItems.Add("Alias", "U_BDOSSrvDsc");
                formItems.Add("Bound", true);
                formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                formItems.Add("Left", left_e);
                formItems.Add("Width", width_e);
                formItems.Add("Top", top);
                formItems.Add("Height", height);
                formItems.Add("UID", itemName);
                formItems.Add("DisplayDesc", true);

                FormsB1.createFormItem(oForm, formItems, out errorText);
                if (errorText != null)
                {
                    return;           
                }

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
            formItems.Add("DataType", BoDataType.dt_PRICE);
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
        
        public static void formDataLoad( SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            SAPbouiCOM.StaticText oStaticText = null;
            oForm.Freeze(true);
            try
            {
                int docEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("OINV").GetValue("DocEntry", 0));
                            
                //-------------------------------------------ანგარიშ-ფაქტურა----------------------------------->
                string cardCode = oForm.DataSources.DBDataSources.Item("OINV").GetValue("CardCode", 0).Trim();
                string caption = BDOSResources.getTranslate("CreateTaxInvoice");
                int taxDocEntry = 0;
                string taxID = "";
                string taxNumber = "";
                string taxSeries = "";
                string taxStatus = "";
                string taxCreateDate = "";

                if (docEntry != 0)
                {
                    Dictionary<string, object> taxDocInfo = BDO_TaxInvoiceSent.getTaxInvoiceSentDocumentInfo( docEntry, "ARInvoice", cardCode);
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

                oStaticText = (SAPbouiCOM.StaticText)oForm.Items.Item("BDO_TaxTxt").Specific;
                oStaticText.Caption = caption;
                //<-------------------------------------------ანგარიშ-ფაქტურა-----------------------------------
            }
            catch (Exception ex)
            {
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


        public static void uiApp_FormDataEvent(  ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(BusinessObjectInfo.FormTypeEx, Program.currentFormCount);

            if (oForm.TypeEx == "60091")
            {
                if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD & BusinessObjectInfo.BeforeAction == false)
                {
                    formDataLoad( oForm, out errorText);
                    setVisibleFormItems(oForm, out errorText);
                }
            }
          
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
                    createFormItems( oForm, out errorText);
                    formDataLoad( oForm, out errorText);
                    SetVisibility(oForm);
                    oForm.Items.Item("4").Click();
                    Program.FORM_LOAD_FOR_ACTIVATE = true;
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
                {
                    if (pVal.ItemUID == "1" && pVal.BeforeAction == true)
                    {
                        CommonFunctions.fillDocRate( oForm, "OINV");
                    }
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_VALIDATE & pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "1980002192")
                    {
                        setVisibleFormItems(oForm, out errorText);
                    }

                    if (oForm.Items.Item("DiscountE").Visible)
                    {
                        if (Program.FORM_LOAD_FOR_ACTIVATE) return;

                        if (pVal.ItemUID == "38" &&
                            (pVal.ItemChanged && (pVal.ColUID == "14" || pVal.ColUID == "1" ||
                                                  (pVal.ColUID == "15" || pVal.ColUID == "11" && !pVal.InnerEvent)) ||
                             (pVal.ColUID == "1" && !pVal.InnerEvent)))
                        {
                            SetInitialLineNetTotals(oForm, pVal.ColUID, pVal.Row);
                            ApplyDiscount(oForm);
                        }

                        else if (pVal.ItemUID == "DiscountE" &&
                                 !pVal.InnerEvent && pVal.ItemChanged)
                        {
                            ApplyDiscount(oForm);
                        }
                    }
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED & pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "BDO_TaxTxt")
                    {
                        oForm.Freeze(true);
                        int newDocEntry = 0;
                        string bstrUDOObjectType = null;

                        itemPressed( oForm, pVal, out newDocEntry, out bstrUDOObjectType, out errorText);

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

                else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE && !pVal.BeforeAction)
                {
                    if (!Program.FORM_LOAD_FOR_ACTIVATE) return;

                    var discount = oForm.Items.Item("DiscountE");

                    if (discount.Visible)
                    {
                        discount.Specific.Value = 0;

                        Matrix oMatrix = oForm.Items.Item("38").Specific;

                        for (var row = 1; row < oMatrix.RowCount; row++)
                        {
                            SetInitialLineNetTotals(oForm, "14", row);
                        }
                    }

                    Program.FORM_LOAD_FOR_ACTIVATE = false;
                }

                else if (pVal.EventType == BoEventTypes.et_FORM_DRAW && !pVal.BeforeAction)
                {
                    CommonFunctions.SetBaseDocRoundingAmountIntoTargetDoc(oForm);
                }

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
            
            if (pVal.ItemUID == "BDO_TaxTxt")
            {
                string taxDoc = oForm.DataSources.UserDataSources.Item("BDO_TaxDoc").ValueEx;
                bstrUDOObjectType = "UDO_F_BDO_TAXS_D";

                if (docEntry != 0 && (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_VIEW_MODE))
                {
                    if (taxDoc == "" && cancelled == "N")
                    {
                        BDO_TaxInvoiceSent.createDocument( objectType, docEntry, "", true, 0, null, false, null, null, out newDocEntry, out errorText);
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

            FormsB1.WB_TAX_AuthorizationsItems(oForm, WBAUT, TXAUT);

        }

        private static void SetVisibility(Form oForm)
        {
            var isDiscountUsed = CompanyDetails.IsDiscountUsed();
            oForm.Items.Item("24").Visible = !isDiscountUsed;
            oForm.Items.Item("283").Visible = !isDiscountUsed;
            oForm.Items.Item("42").Visible = !isDiscountUsed;
            oForm.Items.Item("DiscountE").Visible = isDiscountUsed;

            FormsB1.WB_TAX_AuthorizationsItems(oForm, WBAUT, TXAUT);
        }

        private static void SetInitialLineNetTotals(Form oForm, string column, int row)
        {
            try
            {
                oForm.Freeze(true);

                Matrix oMatrix = oForm.Items.Item("38").Specific;

                var col = oForm.Items.Item("63").Specific.Value == "GEL" ? "21" : "23";

                if (column == "14" && !Program.FORM_LOAD_FOR_ACTIVATE)
                {
                    oMatrix.GetCellSpecific("15", row).Value = 0;
                }

                var initialLineNetTotal =
                    Convert.ToDecimal(FormsB1.cleanStringOfNonDigits(oMatrix.GetCellSpecific(col, row).Value));

                if (initialLineNetTotal == 0) return;
                InitialLineNetTotals[row] = initialLineNetTotal;
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
                var col = oForm.Items.Item("63").Specific.Value == "GEL" ? "21" : "23";

                EditText oEditText = oForm.Items.Item("DiscountE").Specific;
                var discountTotal = string.IsNullOrEmpty(oEditText.Value) ? 0 : Convert.ToDecimal(oEditText.Value);

                decimal docTotal = 0;

                for (var row = 1; row < oMatrix.RowCount; row++)
                {
                    var itemPrice = oMatrix.GetCellSpecific("14", row).Value;
                    if (!string.IsNullOrEmpty(itemPrice))
                    {
                        docTotal += InitialLineNetTotals[row];
                    }
                    else
                    {
                        oEditText.Value = string.Empty;
                        return;
                    }
                }

                for (var row = 1; row < oMatrix.RowCount; row++)
                {
                    var lineNetTotal = InitialLineNetTotals[row];

                    var taxCode = oMatrix.GetCellSpecific("18", row).Value;
                    var taxRate = CommonFunctions.GetVatGroupRate(taxCode, "");

                    var discount = lineNetTotal / docTotal * discountTotal / (1 + taxRate / 100);

                    var lineNetTotalAfterDiscount = Math.Round(lineNetTotal - discount, 4);

                    oMatrix.GetCellSpecific(col, row).Value =
                        FormsB1.ConvertDecimalToStringForEditboxStrings(lineNetTotalAfterDiscount);
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