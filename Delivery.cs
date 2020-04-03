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
    static partial class Delivery
    {
        private static Dictionary<int, decimal> InitialLineNetTotals = new Dictionary<int, decimal>();
        public static void createUserFields( out string errorText)
        {
            errorText = null;
            GC.Collect();
        }

        public static void createFormItems(SAPbouiCOM.Form oForm, out string errorText)
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

            #region Discount field

            
            height = oForm.Items.Item("42").Height;
            top = oForm.Items.Item("42").Top;
            left_e = oForm.Items.Item("42").Left;
            width_e = oForm.Items.Item("42").Width;

            formItems = new Dictionary<string, object>();
            itemName = "DiscountE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "ODLN");
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

        public static void formDataLoad(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            SAPbouiCOM.StaticText oStaticText = null;
            oForm.Freeze(true);
            try
            {
                int docEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("ODLN").GetValue("DocEntry", 0));

                //-------------------------------------------სასაქონლო ზედნადები----------------------------------->
                string caption = BDOSResources.getTranslate("CreateWaybill");
                int wblDocEntry = 0;
                string wblID = "";
                string wblNum = "";
                string wblSts = "";
                string objType = "15";

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

               
            }
            catch (Exception ex)
            {
                oForm.DataSources.UserDataSources.Item("BDO_WblDoc").ValueEx = "";
                oForm.DataSources.UserDataSources.Item("BDO_WblID").ValueEx = "";
                oForm.DataSources.UserDataSources.Item("BDO_WblNum").ValueEx = "";
                oForm.DataSources.UserDataSources.Item("BDO_WblSts").ValueEx = "";

                oStaticText = (SAPbouiCOM.StaticText)oForm.Items.Item("BDO_WblTxt").Specific;
                oStaticText.Caption = BDOSResources.getTranslate("CreateWaybill");

                errorText = ex.Message;
            }
            finally
            {
                oForm.Freeze(false);
                GC.Collect();
            }
        }

        public static void cancellation(SAPbouiCOM.Form oForm, int docEntry, out string errorText)
        {
            errorText = null;

            try
            {
                Dictionary<string, string> wblDocInfo = BDO_Waybills.getWaybillDocumentInfo(docEntry, "15", out errorText);
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
                JournalEntry.cancellation(oForm, docEntry, "15", out errorText);
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

            string docEntrySTR = oForm.DataSources.DBDataSources.Item("ODLN").GetValue("DocEntry", 0);
            docEntrySTR = string.IsNullOrEmpty(docEntrySTR) == true ? "0" : docEntrySTR;
            int docEntry = Convert.ToInt32(docEntrySTR);
            string cancelled = oForm.DataSources.DBDataSources.Item("ODLN").GetValue("CANCELED", 0).Trim();
            string docType = oForm.DataSources.DBDataSources.Item("ODLN").GetValue("DocType", 0).Trim();
            string objectType = "15";

            if (pVal.ItemUID == "BDO_WblTxt")
            {
                string wblDoc = oForm.DataSources.UserDataSources.Item("BDO_WblDoc").ValueEx;
                bstrUDOObjectType = "UDO_F_BDO_WBLD_D";

                if (docEntry != 0 & (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_VIEW_MODE))
                {
                    if (wblDoc == "" && cancelled == "N" && docType == "I")
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
                }
                else
                {
                    errorText = BDOSResources.getTranslate("ToCreateWaybillWriteDocument");
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
                    createFormItems(oForm, out errorText);
                    formDataLoad(oForm, out errorText);
                    SetVisibility(oForm);
                    oForm.Items.Item("4").Click();
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED & pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "BDO_WblTxt")
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

                else if (pVal.EventType == BoEventTypes.et_VALIDATE && !pVal.BeforeAction)
                {
                    if (oForm.Items.Item("DiscountE").Visible)
                    {
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
            }

        }

        public static void uiApp_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(BusinessObjectInfo.FormTypeEx, Program.currentFormCount);

            if (oForm.TypeEx == "140")
            {
                if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD & BusinessObjectInfo.BeforeAction == false)
                {
                    formDataLoad(oForm, out errorText);
                }

                if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD & BusinessObjectInfo.BeforeAction == false & BusinessObjectInfo.ActionSuccess == true)
                {
                    if (Program.canceledDocEntry != 0)
                    {
                        cancellation(oForm, Program.canceledDocEntry, out errorText);
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
                            CommonFunctions.blockNegativeStockByDocDate(oForm, "ODLN", "DLN1", "WhsCode", out rejection);
                            if (rejection)
                            {
                                Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("DocumentCannotBeAdded"));
                                BubbleEvent = false;
                            }
                        }
                    }
                }

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

        private static void SetInitialLineNetTotals(Form oForm, string column, int row)
        {
            try
            {
                oForm.Freeze(true);

                Matrix oMatrix = oForm.Items.Item("38").Specific;

                var col = oForm.Items.Item("63").Specific.Value == "GEL" ? "21" : "23";

                if (column == "14")
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