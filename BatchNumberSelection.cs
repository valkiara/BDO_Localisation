//using SAPbouiCOM;
using static BDO_Localisation_AddOn.Program;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

namespace BDO_Localisation_AddOn
{
    static class BatchNumberSelection
    {
        public static DataTable SelectedBatches { get; set; }

        //private static bool clickOkBtn = false;

        public static void UiApp_ItemEvent(ref SAPbouiCOM.ItemEvent pVal, out bool bubbleEvent)
        {
            bubbleEvent = true;

            if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_CLOSE && !pVal.BeforeAction)
            {
                SAPbouiCOM.Form oForm = uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);
                try
                {
                    oForm.Freeze(true);

                    SelectedBatches = new DataTable();
                    SelectedBatches.Columns.AddRange(new DataColumn[] { new DataColumn("Line", typeof(string)), new DataColumn("ItemCode", typeof(string)), new DataColumn("DistNumber", typeof(string)) });

                    SAPbouiCOM.Matrix oMatrixRowsFromDoc = oForm.Items.Item("3").Specific;

                    for (int j = 1; j <= oMatrixRowsFromDoc.VisualRowCount; j++)
                    {
                        oMatrixRowsFromDoc.Columns.Item("0").Cells.Item(j).Click(SAPbouiCOM.BoCellClickType.ct_Regular);

                        string line = oMatrixRowsFromDoc.Columns.Item("0").Cells.Item(oMatrixRowsFromDoc.GetNextSelectedRow()).Specific.Value;
                        string itemCode = oMatrixRowsFromDoc.Columns.Item("1").Cells.Item(oMatrixRowsFromDoc.GetNextSelectedRow()).Specific.Value;

                        DataRow[] dtRows = SelectedBatches.Select($"Line={line}");
                        foreach (var row in dtRows)
                        {
                            row.Delete();
                        }
                        SelectedBatches.AcceptChanges();

                        SAPbouiCOM.Matrix oMatrixSelectedBatches = oForm.Items.Item("5").Specific;
                        for (int i = 1; i <= oMatrixSelectedBatches.VisualRowCount; i++)
                        {
                            string distNumber = oMatrixSelectedBatches.Columns.Item("1").Cells.Item(i).Specific.Value;
                            SelectedBatches.Rows.Add(line, itemCode, distNumber);
                        }
                    }

                    if (SelectedBatches.Rows.Count == 0)
                        SelectedBatches = null;
                }
                catch (Exception ex)
                {
                    SelectedBatches = null;
                    throw new Exception(ex.Message);
                }
                finally
                {
                    oForm.Freeze(false);
                }
            }
        }
    }
}
