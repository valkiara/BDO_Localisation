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

            var leftS = oForm.Items.Item("34").Left;
            var widthS = oForm.Items.Item("34").Width;
            var height = oForm.Items.Item("34").Height;
            var top = oForm.Items.Item("34").Top + height + 1;

            var leftE = oForm.Items.Item("33").Left;
            var widthE = oForm.Items.Item("33").Width;

            var formItems = new Dictionary<string, object>();
            var itemName = "DiscountS"; //10 characters
            formItems.Add("Type", BoFormItemTypes.it_STATIC);
            formItems.Add("Left", leftS);
            formItems.Add("Width", widthS);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Discount"));
            formItems.Add("Enabled", true);

            FormsB1.createFormItem(oForm, formItems, out var errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "DiscountE"; //10 characters
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

            FormsB1.createFormItem(oForm, formItems, out errorText);
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
                    oForm.Items.Item("4").Click();
                }
                
                else if (pVal.EventType == BoEventTypes.et_VALIDATE && !pVal.BeforeAction && !pVal.InnerEvent)
                {
                    if (pVal.ItemUID == "DiscountE")
                    {
                        ApplyDiscount(oForm);
                    }
                }
            }
        }

        private static void SetInitialItemGrossPrices(Form oForm, int row)
        {
            Matrix oMatrix = oForm.Items.Item("38").Specific;
            var initialItemGrossPrice = Convert.ToDecimal(FormsB1.cleanStringOfNonDigits(oMatrix.GetCellSpecific("20", row).Value));

            InitialItemGrossPrices[row] = initialItemGrossPrice;
        }

        private static void ApplyDiscount(Form oForm)
        {
            try
            {
                oForm.Freeze(true);
                Matrix oMatrix = oForm.Items.Item("38").Specific;
                EditText oEditText = oForm.Items.Item("DiscountE").Specific;

                var discountTotal = oEditText.Value;
                if (string.IsNullOrEmpty(discountTotal)) return;

                var quantityTotal = 0;

                for (var row = 1; row < oMatrix.RowCount; row++)
                {
                    var unitPrice = oMatrix.GetCellSpecific("14", row).Value;
                    if (!string.IsNullOrEmpty(unitPrice))
                    {
                        quantityTotal += Convert.ToDecimal(oMatrix.GetCellSpecific("11", row).Value);
                    }
                    else
                    {
                        Program.uiApp.StatusBar.SetSystemMessage("Fill Item Prices", BoMessageTime.bmt_Short);
                        oEditText.Value = string.Empty;
                        return;
                    }
                }

                var discount = Convert.ToDecimal(discountTotal) / quantityTotal;

                for (var row = 1; row < oMatrix.RowCount; row++)
                {
                    var grossUnitAmt = InitialItemGrossPrices[row];

                    var grossAfterDiscount = Math.Round(grossUnitAmt - discount, 4);

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