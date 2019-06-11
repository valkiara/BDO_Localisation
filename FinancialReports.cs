using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.IO;

namespace BDO_Localisation_AddOn
{
    class FinancialReports
    {
        public static void ExportToExcel(SAPbouiCOM.Form oForm, string itemMatrix, string indexIndentLevel, string indexName, string indexbalance, out string errorText)
        {
            errorText = null;

            Excel.Application excelApp = new Excel.Application();


            Excel.Workbook Workbook1 = excelApp.Workbooks.Add();
            Worksheet WorkSheet1 = Workbook1.Worksheets[1];


            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item(itemMatrix).Specific));
            int columnIndex = 1;

            for (int k = 0; k < oMatrix.Columns.Count; k++)
            {
                SAPbouiCOM.Column oColumn = oMatrix.Columns.Item(k);
                if (oColumn.Visible == false)
                {
                    continue;
                }

                WorkSheet1.Cells[1, columnIndex] = oColumn.Title;
                WorkSheet1.Cells[1, columnIndex].Font.Bold = true;
                WorkSheet1.Cells[1, columnIndex].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = 3d;
                columnIndex++;
            }

            string indentLevelStr = "";
            string nameStr = "";
            string balanceStr = "";

            bool fontBold = false;

            for (int i = 1; i <= oMatrix.VisualRowCount; i++)
            {
                try
                {
                    indentLevelStr = oMatrix.Columns.Item(indexIndentLevel).Cells.Item(i).Specific.Value;
                    nameStr = oMatrix.Columns.Item(indexName).Cells.Item(i).Specific.Value;
                    balanceStr = oMatrix.Columns.Item(indexbalance).Cells.Item(i).Specific.Value;
                }
                catch { }
                int colIndex = 1;

                for (int k = 0; k < oMatrix.Columns.Count; k++)
                {
                    SAPbouiCOM.Column oColumn = oMatrix.Columns.Item(k);
                    if (oColumn.Visible == false)
                    {
                        continue;
                    }

                    SAPbouiCOM.Cell oCell = oColumn.Cells.Item(i);
                    try
                    {
                        WorkSheet1.Cells[i + 1, colIndex] = oCell.Specific.Value;
                    }
                    catch { }

                    if (fontBold == true)
                    {
                        WorkSheet1.Cells[i + 1, colIndex].Font.Bold = true;

                    }

                    colIndex++;
                }

                if (indentLevelStr != "")
                {
                    try
                    {
                        WorkSheet1.Cells[i + 1, 1].IndentLevel = Convert.ToInt32(indentLevelStr);
                    }
                    catch { }
                }

                if (balanceStr.Contains("_") == true || balanceStr.Contains("=") == true || nameStr.Contains("Total") == true)
                {
                    fontBold = true;
                }
                else
                {
                    fontBold = false;
                }
            }

            WorkSheet1.Columns.AutoFit();

            excelApp.Visible = true;
        }

    }
}
