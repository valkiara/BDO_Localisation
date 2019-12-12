using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace BDO_Localisation_AddOn
{
    static partial class BDOSDeleteUDF
    {
        public static void createForm(  out string errorText)
        {
            errorText = null;

            int formHeight = Program.uiApp.Desktop.Width;
            int formWidth = Program.uiApp.Desktop.Width;

            //ფორმის აუცილებელი თვისებები
            Dictionary<string, object> formProperties = new Dictionary<string, object>();
            formProperties.Add("UniqueID", "BDOSDeleteUDFForm");
            formProperties.Add("BorderStyle", SAPbouiCOM.BoFormBorderStyle.fbs_Sizable);
            formProperties.Add("Title", BDOSResources.getTranslate("DeleteUDF"));
            //formProperties.Add("Left", (Program.uiApp.Desktop.Width - formWidth) / 2);
            formProperties.Add("ClientWidth", formWidth);
            //formProperties.Add("Top", (Program.uiApp.Desktop.Height - formHeight) / 3);
            formProperties.Add("ClientHeight", formHeight);

            SAPbouiCOM.Form oForm;
            bool newForm;
            bool formExist = FormsB1.createForm( formProperties, out oForm, out newForm, out errorText);

            if (errorText != null)
            {
                return;
            }

            if (formExist == true)
            {
                if (newForm == true)
                {
                    errorText = null;
                    Dictionary<string, object> formItems;
                    string itemName = "";

                    int left_s = 6;
                    int left_e = 120;
                    int height = 15;
                    int top = 6;
                    int width_s = 121 - 15;
                    int width_e = 148;

                    formItems = new Dictionary<string, object>();
                    itemName = "tableIDS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("Table"));
                    formItems.Add("LinkTo", "tableIDE");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    Dictionary<string, string> listValidValuesDict = new Dictionary<string, string>();
                    SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    string query = getTableList();
                    try
                    {
                        oRecordSet.DoQuery(query);
                        string tableID;

                        while (!oRecordSet.EoF)
                        {
                            tableID = oRecordSet.Fields.Item("TableID").Value.ToString();
                            listValidValuesDict.Add(tableID, tableID);
                            oRecordSet.MoveNext();
                        }                       
                    }
                    catch
                    {
                        return;
                    }
                    finally
                    {
                        oForm.Freeze(false);
                        Marshal.ReleaseComObject(oRecordSet);
                        oRecordSet = null;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "tableIDE"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("Length", 30);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                    formItems.Add("Left", left_e);
                    formItems.Add("Width", width_e);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
                    formItems.Add("Description", BDOSResources.getTranslate("Table"));
                    formItems.Add("ValidValues", listValidValuesDict);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    top = top + height + 1;

                    formItems = new Dictionary<string, object>();
                    itemName = "aliasIDS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("Field"));
                    formItems.Add("LinkTo", "aliasIDE");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "aliasIDE"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("Length", 30);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("Left", left_e);
                    formItems.Add("Width", width_e);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("Description", BDOSResources.getTranslate("Field"));

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    top = top + 3 * height + 1;

                    itemName = "checkB";
                    formItems = new Dictionary<string, object>();
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", 19);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Image", "HANA_CHECKBOX_CH");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left_s = left_s + 20;

                    itemName = "unCheckB";
                    formItems = new Dictionary<string, object>();
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", 19);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Image", "HANA_CHECKBOX_UH");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left_s = left_s + 20;

                    itemName = "fillB";
                    formItems = new Dictionary<string, object>();
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", 65);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("Fill"));

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    itemName = "deleteB";
                    formItems = new Dictionary<string, object>();
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left_s + 65 + 2);
                    formItems.Add("Width", 65);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("Delete"));

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    top = top + height + 1;
                    left_s = 6;

                    formItems = new Dictionary<string, object>();
                    itemName = "udfMTR"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", oForm.ClientWidth);
                    formItems.Add("Top", top);
                    formItems.Add("Height", (oForm.ClientHeight - 25 - top));
                    formItems.Add("UID", itemName);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("AffectsFormMode", false);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    SAPbouiCOM.Matrix oMatrixImport = ((SAPbouiCOM.Matrix)(oForm.Items.Item("udfMTR").Specific));
                    oMatrixImport.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
                    SAPbouiCOM.Columns oColumns = oMatrixImport.Columns;
                    SAPbouiCOM.Column oColumn;

                    SAPbouiCOM.DataTable oDataTable;
                    oDataTable = oForm.DataSources.DataTables.Add("udfMTR");

                    oDataTable.Columns.Add("LineNum", SAPbouiCOM.BoFieldsType.ft_Integer, 50);
                    oDataTable.Columns.Add("CheckBox", SAPbouiCOM.BoFieldsType.ft_Text, 1);
                    oDataTable.Columns.Add("FieldID", SAPbouiCOM.BoFieldsType.ft_Integer, 50);
                    oDataTable.Columns.Add("AliasID", SAPbouiCOM.BoFieldsType.ft_Text, 50);
                    oDataTable.Columns.Add("Descr", SAPbouiCOM.BoFieldsType.ft_Text, 80);
                    oDataTable.Columns.Add("Descra", SAPbouiCOM.BoFieldsType.ft_Text, 80);

                    string UID = "udfMTR";

                    foreach (SAPbouiCOM.DataColumn column in oDataTable.Columns)
                    {
                        string columnName = column.Name;

                        if (columnName == "LineNum")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = "#";
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind(UID, columnName);
                        }
                        else if (columnName == "CheckBox")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
                            oColumn.TitleObject.Caption = "";
                            oColumn.Editable = true;
                            oColumn.ValOff = "N";
                            oColumn.ValOn = "Y";
                            oColumn.DataBind.Bind(UID, columnName);
                        }
                        else if (columnName == "Descra")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
                            oColumn.TitleObject.Caption = "";
                            oColumn.Editable = true;
                            oColumn.Visible = false;
                            oColumn.DataBind.Bind(UID, columnName);
                        }
                        else
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate(columnName);
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind(UID, columnName);
                        }
                    }

                    oMatrixImport.Clear();
                    oMatrixImport.LoadFromDataSource();
                    oMatrixImport.AutoResizeColumns();
                }
                oForm.Visible = true;
                oForm.Select();
            }
            GC.Collect();
        }

        public static void addMenus()
        {
            SAPbouiCOM.MenuItem menuItem;
            SAPbouiCOM.MenuItem fatherMenuItem;
            SAPbouiCOM.MenuCreationParams oCreationPackage;

            try
            {
                
                fatherMenuItem = Program.uiApp.Menus.Item("8704");
               
                // Add a pop-up menu item
                oCreationPackage = (SAPbouiCOM.MenuCreationParams)Program.uiApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                oCreationPackage.Checked = false;
                oCreationPackage.Enabled = true;
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "BDOSDeleteUDFForm";
                oCreationPackage.String = BDOSResources.getTranslate("DeleteUDF");
                oCreationPackage.Position = -1;

                menuItem = fatherMenuItem.SubMenus.AddEx(oCreationPackage);
            }
            catch
            {
                
            }
        }

        public static void uiApp_ItemEvent(  string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK)
                {
                    if (pVal.ItemUID == "fillB" & pVal.BeforeAction == false)
                    {
                        filludfMTR( oForm, out errorText);
                    }

                    if ((pVal.ItemUID == "checkB" || pVal.ItemUID == "unCheckB") && pVal.BeforeAction == false)
                    {
                        checkUncheckMTR(oForm, pVal.ItemUID, out errorText);
                    }

                    if (pVal.ItemUID == "deleteB" & pVal.BeforeAction == false)
                    {
                        deleteUDF( oForm, out errorText);
                        filludfMTR( oForm, out errorText);
                    }
                }
            }
        }

        public static void checkUncheckMTR(SAPbouiCOM.Form oForm, string checkOperation, out string errorText)
        {
            errorText = null;
            try
            {
                oForm.Freeze(true);
                SAPbouiCOM.CheckBox oCheckBox;
                SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("udfMTR").Specific));

                for (int j = 1; j <= oMatrix.RowCount; j++)
                {
                    oCheckBox = oMatrix.Columns.Item("CheckBox").Cells.Item(j).Specific;
                    oCheckBox.Checked = (checkOperation == "checkB");
                }
                oForm.Freeze(false);
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
            finally
            {
                oForm.Freeze(false);
                GC.Collect();
            }
        }

        public static void filludfMTR(  SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("udfMTR");
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string tableID = oForm.DataSources.UserDataSources.Item("tableIDE").ValueEx;
            string aliasID = oForm.DataSources.UserDataSources.Item("aliasIDE").ValueEx;

            if (string.IsNullOrEmpty(tableID))
            {
                errorText = BDOSResources.getTranslate("TheFollowingFieldIsMandatory") + " : \"" + oForm.Items.Item("tableIDS").Specific.caption + "\"";
                Program.uiApp.SetStatusBarMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                return;
            }

            string query = getQueryForImport(tableID, aliasID);

            oRecordSet.DoQuery(query);
            oDataTable.Rows.Clear();

            try
            {
                int rowIndex = 0;

                while (!oRecordSet.EoF)
                {
                    oDataTable.Rows.Add();
                    oDataTable.SetValue("LineNum", rowIndex, rowIndex + 1);
                    oDataTable.SetValue("CheckBox", rowIndex, "N");
                    oDataTable.SetValue("FieldID", rowIndex, oRecordSet.Fields.Item("FieldID").Value);
                    oDataTable.SetValue("AliasID", rowIndex, oRecordSet.Fields.Item("AliasID").Value);
                    oDataTable.SetValue("Descr", rowIndex, oRecordSet.Fields.Item("Descr").Value);

                    oRecordSet.MoveNext();
                    rowIndex++;
                }

                SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("udfMTR").Specific));
                oForm.Freeze(true);
                oMatrix.Clear();
                oMatrix.LoadFromDataSource();
                oMatrix.AutoResizeColumns();
                oForm.Update();
                oForm.Freeze(false);
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
            finally
            {
                oForm.Freeze(false);
                Marshal.ReleaseComObject(oRecordSet);
                oRecordSet = null;
                Marshal.ReleaseComObject(oDataTable);
                oDataTable = null;              
            }
        }

        public static void deleteUDF(  SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            string tableID = oForm.DataSources.UserDataSources.Item("tableIDE").ValueEx;
            string aliasID = oForm.DataSources.UserDataSources.Item("aliasIDE").ValueEx;

            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("udfMTR").Specific));
            int row = 1;

            while (row <= oMatrix.RowCount)
            {
                if (oMatrix.GetCellSpecific("CheckBox", row).Checked)
                {
                    int fieldID = Convert.ToInt32(oMatrix.GetCellSpecific("FieldID", row).Value);
                    aliasID = oMatrix.GetCellSpecific("AliasID", row).Value.ToString();

                    UDO.DeleteUDF( tableID, fieldID, out errorText);
                    if(string.IsNullOrEmpty(errorText))
                    {
                        Program.uiApp.SetStatusBarMessage(BDOSResources.getTranslate("OperationCompletedSuccessfully") + "! UDF : \"" + aliasID + "\"", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                    }
                    else
                    {
                        Program.uiApp.SetStatusBarMessage(BDOSResources.getTranslate("ErrorRemovingUDF") + " : \"" + aliasID + "\" - " + errorText, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    }
                }
                row++;
            }
        }

        private static string getQueryForImport(string tableID, string aliasID)
        {
            string query = @"SELECT ""FieldID"", ""AliasID"", ""Descr"" FROM ""CUFD""  " +
                           @"WHERE ""TableID"" = '" + tableID + @"'";
            if (string.IsNullOrEmpty(aliasID) == false)
                query = query + @" AND ""AliasID"" = '" + aliasID + "'";

            return query;
        }

        private static string getTableList()
        {
            string query = @"SELECT DISTINCT ""TableID"" FROM ""CUFD""";
            return query;
        }
    }
}
