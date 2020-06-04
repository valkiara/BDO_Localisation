using System;
using System.Collections.Generic;
using SAPbouiCOM;

namespace BDO_Localisation_AddOn
{
    static class SalesOrder
    {
        private static Dictionary<int, decimal> InitialLineNetTotals = new Dictionary<int, decimal>();

        private static void CreateFormItems(Form oForm)
        {
            #region Discount field

            var height = oForm.Items.Item("42").Height;
            var top = oForm.Items.Item("42").Top;
            var leftE = oForm.Items.Item("42").Left;
            var widthE = oForm.Items.Item("42").Width;

            var formItems = new Dictionary<string, object>();
            var itemName = "DiscountE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "ORDR");
            formItems.Add("Alias", "U_Discount");
            formItems.Add("Bound", true);
            formItems.Add("Type", BoFormItemTypes.it_EDIT);
            formItems.Add("DataType", BoDataType.dt_PRICE);
            formItems.Add("Left", leftE);
            formItems.Add("Width", widthE);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Discount"));
            formItems.Add("DisplayDesc", true);
            formItems.Add("SetAutoManaged", true);
            formItems.Add("Visible", false);

            FormsB1.createFormItem(oForm, formItems, out var errorText);
            if (errorText != null)
            {
                return;
            }

            #endregion
        }

        public static void uiApp_ItemEvent(  string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            
            if (pVal.EventType != BoEventTypes.et_FORM_UNLOAD)
            {
                Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (pVal.EventType == BoEventTypes.et_FORM_LOAD & pVal.BeforeAction)
                {
                    CreateFormItems(oForm);
                    SetVisibility(oForm);
                    oForm.Items.Item("4").Click();
                    Program.FORM_LOAD_FOR_ACTIVATE = true;
                }

                else if (pVal.EventType == BoEventTypes.et_VALIDATE && !pVal.BeforeAction)
                {
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