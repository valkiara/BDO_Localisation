using System;
using System.Collections.Generic;
using SAPbouiCOM;

namespace BDO_Localisation_AddOn
{
    static class SalesOrder
    {
        private static Dictionary<int, decimal> InitialItemGrossPrices = new Dictionary<int, decimal>();

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
            formItems.Add("DataType", BoDataType.dt_SUM);
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