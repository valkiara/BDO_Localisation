using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BDO_Localisation_AddOn
{
    static partial class IssueForProduction
    {
        public static string itemCodeOld;
        public static string warehouseOld;

        public static void uiApp_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (pVal.ItemUID == "13" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_GOT_FOCUS)
                {
                    if (!pVal.BeforeAction)
                    {
                        if (pVal.ColUID == "61") //Order No.
                        {
                            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item(pVal.ItemUID).Specific));
                            itemCodeOld = oMatrix.GetCellSpecific("1", pVal.Row).Value;
                        }
                        else if (pVal.ColUID == "15") //Whse
                        {
                            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item(pVal.ItemUID).Specific));
                            warehouseOld = oMatrix.GetCellSpecific(pVal.ColUID, pVal.Row).Value;
                        }
                    }
                }

                if (pVal.ItemUID == "13" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_LOST_FOCUS)
                {
                    if (!pVal.BeforeAction)
                    {
                        if (pVal.ColUID == "61" || pVal.ColUID == "15") //Order No. || //Whse
                        {
                            updateInStockByWarehouseAndDate(oForm, pVal.Row);
                        }
                    }
                }

                if (pVal.ItemUID == "9" && pVal.ItemChanged && !pVal.BeforeAction)
                {
                    updateInStockByWarehouseAndDate(oForm);
                }

                if (Program.selectItemsToCopyOkClick && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE && !pVal.BeforeAction)
                {
                    Program.selectItemsToCopyOkClick = false;
                    updateInStockByWarehouseAndDate(oForm);
                }
            }
        }

        private static void updateInStockByWarehouseAndDate(SAPbouiCOM.Form oForm, int rowIndex = 0)
        {
            string docDate = oForm.DataSources.DBDataSources.Item("OIGE").GetValue("DocDate", 0);
            try
            {
                oForm.Freeze(true);

                SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("13").Specific));
                int rowCount = rowIndex == 0 ? oMatrix.RowCount : rowIndex;
                int i = rowIndex == 0 ? 1 : rowIndex;

                for (; i <= rowCount; i++)
                {
                    string warehouse = oMatrix.GetCellSpecific("15", i).Value;
                    string itemCode = oMatrix.GetCellSpecific("1", i).Value;

                    if ((itemCode != itemCodeOld || warehouse != warehouseOld) || rowIndex == 0)
                    {
                        if (!string.IsNullOrEmpty(itemCode) && !string.IsNullOrEmpty(warehouse) && !string.IsNullOrEmpty(docDate))
                        {
                            oMatrix.Columns.Item("U_BDOSInStck").Cells.Item(i).Specific.Value = FormsB1.ConvertDecimalToStringForEditboxStrings(CommonFunctions.getInStockByWarehouseAndDate(itemCode, warehouse, docDate));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Program.uiApp.StatusBar.SetSystemMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
                itemCodeOld = null;
                warehouseOld = null;
            }
        }
    }
}
