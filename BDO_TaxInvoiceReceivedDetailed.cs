using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace BDO_Localisation_AddOn
{
    static partial class BDO_TaxInvoiceReceivedDetailed
    {
        private static string CardCodeG;
        private static string opDateG;

        private static SAPbouiCOM.Form FormBDOSInternetBanking;
        public static void createForm(SAPbouiCOM.Form FormBDOSInternetBanking1, out SAPbouiCOM.Form oForm, string CardCode, string opDate, out string errorText)
        {
            FormBDOSInternetBanking = FormBDOSInternetBanking1;

            CardCodeG = CardCode;
            opDateG = opDate;

            oForm = null;

            int formHeight = Program.uiApp.Desktop.Width;
            int formWidth = Program.uiApp.Desktop.Width;

            //ფორმის აუცილებელი თვისებები
            string formTitle = BDOSResources.getTranslate("APDocuments");

            Dictionary<string, object> formProperties = new Dictionary<string, object>();
            formProperties.Add("UniqueID", "BDOSAPDOC");
            formProperties.Add("BorderStyle", SAPbouiCOM.BoFormBorderStyle.fbs_Sizable);
            formProperties.Add("Title", formTitle);
            formProperties.Add("ClientWidth", formWidth);
            formProperties.Add("ClientHeight", formHeight);

            bool newForm;
            bool formExist = FormsB1.createForm(formProperties, out oForm, out newForm, out errorText);

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
                    NumberFormatInfo Nfi = new NumberFormatInfo() { NumberDecimalSeparator = "." };

                    int left = 6;
                    int top = 6;
                    int height_e = 15;
                    int height = oForm.ClientHeight - top - 8 * height_e - 1 - 30;
                    int width = oForm.ClientWidth;

                    int left_s = 6;
                    int left_e = 90;
                    int width_s = 80;
                    int width_e = 148;



                    formItems = new Dictionary<string, object>();
                    itemName = "FiltrS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("Filter"));
                    formItems.Add("LinkTo", "Filtr");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    List<string> listValidValues = new List<string>();
                    listValidValues.Add(BDOSResources.getTranslate("All"));
                    listValidValues.Add(BDOSResources.getTranslate("ApInvoice"));
                    listValidValues.Add(BDOSResources.getTranslate("APCreditNote"));

                    formItems = new Dictionary<string, object>();
                    itemName = "Filtr"; //10 characters
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
                    formItems.Add("Height", height_e);
                    formItems.Add("UID", itemName);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
                    formItems.Add("ValidValues", listValidValues);
                    formItems.Add("Visible", true);
                    formItems.Add("AffectsFormMode", false);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }


                    top = top + 20;

                    formItems = new Dictionary<string, object>();
                    itemName = "InvoiceMTR"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
                    formItems.Add("Left", left);
                    formItems.Add("Width", width);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("AffectsFormMode", false);


                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("InvoiceMTR").Specific));
                    oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single;
                    SAPbouiCOM.Columns oColumns = oMatrix.Columns;
                    SAPbouiCOM.Column oColumn;

                    SAPbouiCOM.DataTable oDataTable;

                    oDataTable = oForm.DataSources.DataTables.Add("InvoiceMTR");
                    oDataTable.Columns.Add("LineNum", SAPbouiCOM.BoFieldsType.ft_Integer, 50);
                    oDataTable.Columns.Add("CheckBox", SAPbouiCOM.BoFieldsType.ft_Text, 1); // 
                    oDataTable.Columns.Add("DocType", SAPbouiCOM.BoFieldsType.ft_Text, 50);
                    oDataTable.Columns.Add("DocEntry", SAPbouiCOM.BoFieldsType.ft_Integer, 50);
                    oDataTable.Columns.Add("DocNum", SAPbouiCOM.BoFieldsType.ft_Integer, 50);
                    oDataTable.Columns.Add("DocDate", SAPbouiCOM.BoFieldsType.ft_Date, 50);
                    oDataTable.Columns.Add("Currency", SAPbouiCOM.BoFieldsType.ft_Text, 50);
                    oDataTable.Columns.Add("Total", SAPbouiCOM.BoFieldsType.ft_Sum);


                    oDataTable.Columns.Add("Test", SAPbouiCOM.BoFieldsType.ft_Text, 100);

                    SAPbouiCOM.LinkedButton oLink;

                    string UID = "InvoiceMTR";

                    foreach (SAPbouiCOM.DataColumn column in oDataTable.Columns)
                    {
                        string columnName = column.Name;

                        if (columnName == "LineNum")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = "#";
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind(UID, columnName);
                            oColumn.AffectsFormMode = false;
                        }
                        else if (columnName == "CheckBox")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Selected");
                            oColumn.Editable = true;
                            oColumn.ValOff = "N";
                            oColumn.ValOn = "Y";
                            oColumn.DataBind.Bind(UID, columnName);
                            oColumn.AffectsFormMode = false;
                        }

                        else if (columnName == "DocType")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate(columnName);
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind(UID, columnName);
                            oColumn.AffectsFormMode = false;
                            oColumn.DisplayDesc = true;
                            oColumn.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;

                            List<string> InvoicelistValidValues = new List<string>();
                            InvoicelistValidValues.Add(BDOSResources.getTranslate("APInvoice")); //0 //შესყიდვა
                            InvoicelistValidValues.Add(BDOSResources.getTranslate("APCreditNote")); //1 //შესყიდვის კორექტირება

                            for (int i = 0; i < InvoicelistValidValues.Count(); i++)
                            {
                                oColumn.ValidValues.Add(i.ToString(), InvoicelistValidValues[i]);
                            }
                        }

                        else if (columnName == "DocEntry")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate(columnName);
                            oColumn.Editable = false;
                            oLink = oColumn.ExtendedObject;
                            oLink.LinkedObjectType = "24";

                            oColumn.DataBind.Bind(UID, columnName);
                            oColumn.AffectsFormMode = false;
                        }


                        else
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate(columnName);
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind(UID, columnName);
                            oColumn.AffectsFormMode = false;
                        }
                    }
                    oMatrix.Clear();
                    oMatrix.LoadFromDataSource();
                    oMatrix.AutoResizeColumns();

                    //ღილაკები
                    top = oForm.ClientHeight - 25;
                    height_e = height_e + 4;
                    width_s = 65;

                    itemName = "1";
                    formItems = new Dictionary<string, object>();
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height_e);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", "OK");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left_s = left_s + width_s + 2;

                    itemName = "2";
                    formItems = new Dictionary<string, object>();
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height_e);
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }
                }

                oForm.Visible = true;
                oForm.Select();
            }
            GC.Collect();
        }

        public static void reArrangeFormItems(SAPbouiCOM.Form oForm)
        {
            int left = 6;
            int top = 6;
            int height_e = 15;
            int height = oForm.ClientHeight - top - 8 * height_e - 1 - 30;
            int width = oForm.ClientWidth;

            SAPbouiCOM.Item oItem = oForm.Items.Item("Filtr");
            oItem.Top = top;

            top = top + 20;

            oItem = oForm.Items.Item("InvoiceMTR");
            oItem.Top = top;
            oItem.Height = height;
            oItem.Width = width;
            oItem.Left = left;

            top = oForm.ClientHeight - 25;

            oItem = oForm.Items.Item("1");
            oItem.Top = top;
            oItem = oForm.Items.Item("2");
            oItem.Top = top;
        }

        public static void resizeForm(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            try
            {
                reArrangeFormItems(oForm);
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
        }



        public static void fillInvoicesMTR(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
           
            SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("InvoiceMTR");
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string query = "";
            string Filter = oForm.Items.Item("Filtr").Specific.Value;
            Filter = Filter == "" ? "0" : Filter;

            query = GetInvoicesMTRQuery(Filter);
            oRecordSet.DoQuery(query);

            oDataTable.Rows.Clear();

            try
            {
                int rowIndex = 0;
                int DocEntry;
                int AgrNo;
                string PrjCode;
                int DocNum;
                DateTime DocDate;
                decimal Total;
                string Currency;
                string DocType;

                while (!oRecordSet.EoF)
                {
                    DocType = oRecordSet.Fields.Item("DocType").Value;
                    DocEntry = Convert.ToInt32(oRecordSet.Fields.Item("DocEntry").Value);
                    DocNum = Convert.ToInt32(oRecordSet.Fields.Item("DocNum").Value);
                    DocDate = oRecordSet.Fields.Item("DocDate").Value;
                    Total = Convert.ToDecimal(oRecordSet.Fields.Item("DocTotal").Value);
                    Currency = oRecordSet.Fields.Item("DocCur").Value;

                    oDataTable.Rows.Add();
                    oDataTable.SetValue("LineNum", rowIndex, rowIndex + 1);
                    oDataTable.SetValue("DocType", rowIndex, DocType);
                    oDataTable.SetValue("DocEntry", rowIndex, DocEntry);
                    oDataTable.SetValue("DocNum", rowIndex, DocNum);
                    oDataTable.SetValue("DocDate", rowIndex, DocDate);
                    oDataTable.SetValue("Total", rowIndex, Convert.ToDouble(Total));
                    oDataTable.SetValue("Currency", rowIndex, Currency);


                    oRecordSet.MoveNext();
                    rowIndex++;
                }

                SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("InvoiceMTR").Specific));
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
            }
        }

        public static string GetInvoicesMTRQuery( string Filter = "0")
        {

            DateTime opDateDt = Convert.ToDateTime(DateTime.ParseExact(opDateG, "yyyyMMdd", CultureInfo.InvariantCulture));
            DateTime firstDayMonth = new DateTime(opDateDt.Year, opDateDt.Month, 1);
            DateTime lastDayMonth = firstDayMonth.AddMonths(1).AddDays(-1);

            string str = @"
                        select ""Inv"".* from
                        (select 
                        '0' as ""DocType"",
                        OPCH.""DocEntry"",
                        OPCH.""DocNum"",
                        OPCH.""DocDate"",
                        OPCH.""DocCur"",
                        OPCH.""DocTotal"" from OPCH
                        where
                        OPCH.""CANCELED""= 'N'
                        AND ""OPCH"".""CardCode"" = N'" + CardCodeG + @"'
                        AND ""OPCH"".""DocDate"" >= '" + firstDayMonth.ToString("yyyyMMdd") + @"' AND ""OPCH"".""DocDate"" <= '" + lastDayMonth.ToString("yyyyMMdd") + @"'

                        union all

                        select
                        '1', 
                        ORPC.""DocEntry"",
                        ORPC.""DocNum"",
                        ORPC.""DocDate"",
                        ORPC.""DocCur"",
                        ORPC.""DocTotal"" from ORPC
                        where
                        ORPC.""CANCELED""= 'N'
                        AND ""ORPC"".""CardCode"" = N'" + CardCodeG + @"'
                        AND ""ORPC"".""DocDate"" >= '" + firstDayMonth.ToString("yyyyMMdd") + @"' AND ""ORPC"".""DocDate"" <= '" + lastDayMonth.ToString("yyyyMMdd") + @"'
                        ) as ""Inv""
                        left join ""@BDO_TXR1"" on 
                        ""@BDO_TXR1"".""U_baseDocT"" = ""Inv"".""DocType"" and ""@BDO_TXR1"".""U_baseDoc"" = ""Inv"".""DocEntry""
                        where ""@BDO_TXR1"".""DocEntry"" is null" + (Filter != "0" ? @" and ""Inv"".""DocType""='" + Convert.ToString(Convert.ToInt32(Filter) - 1) + "'" : "") +
                        @" order by ""Inv"".""DocDate"" 
                        ";

                        
            

            return str;
        }


        public static void uiApp_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;


            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_RESIZE & pVal.BeforeAction == false)
                {
                    oForm.Freeze(true);
                    resizeForm(oForm, out errorText);
                    oForm.Freeze(false);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_COMBO_SELECT & pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "Filtr")
                    {
                        fillInvoicesMTR(oForm, out errorText);
                    }
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED)
                {
                    if (pVal.BeforeAction)
                    {


                        SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("InvoiceMTR").Specific));
                        SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("InvoiceMTR");
                        string DocType = oDataTable.GetValue("DocType", pVal.Row - 1).ToString();

                        SAPbouiCOM.Column oColumn;
                        oColumn = oMatrix.Columns.Item(pVal.ColUID);
                        if (DocType == "0")
                        {
                            SAPbouiCOM.LinkedButton oLink = oColumn.ExtendedObject;
                            oLink.LinkedObjectType = "18"; 
                        }
                        else
                        {
                            SAPbouiCOM.LinkedButton oLink = oColumn.ExtendedObject;
                            oLink.LinkedObjectType = "19"; //SAPbouiCOM.BoLinkedObject.lf_Employee}

                        }
                    }
                }
                    if (pVal.ItemUID == "InvoiceMTR")
                {

                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.ColUID == "CheckBox")
                    {
                        SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("InvoiceMTR").Specific));
                        oForm.Freeze(true);
                        oMatrix.FlushToDataSource();
                        oForm.Update();
                        oForm.Freeze(false);
                    }
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK)
                {
                    if (pVal.BeforeAction == false)
                    {
                        if (pVal.ItemUID == "1")
                        {
                            SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("InvoiceMTR");
                            DataTable NewRowsTable = new DataTable();

                            NewRowsTable.Columns.Add("DocEntry", typeof(Int32));
                            NewRowsTable.Columns.Add("DocType", typeof(string));

                            string checkBox;
                            for (int i = 0; i < oDataTable.Rows.Count; i++)
                            {
                                checkBox = oDataTable.GetValue("CheckBox", i);
                                if (checkBox == "Y")
                                {
                                    DataRow dataRow = NewRowsTable.Rows.Add();
                                    dataRow["DocEntry"] = oDataTable.GetValue("DocEntry", i);
                                    dataRow["DocType"] = oDataTable.GetValue("DocType", i);
                                }
                            }

                            BDO_TaxInvoiceReceived.AddMult(FormBDOSInternetBanking, NewRowsTable);

                            oForm.Close();

                        }
                    }
                }







            }
        }
    }
}
