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
    static partial class BDOSIncomingDocuments
    {
       
        public static void createForm(  SAPbouiCOM.Form FormBDOSInternetBanking, out SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            oForm = null;

            int formHeight = Program.uiApp.Desktop.Width;
            int formWidth = Program.uiApp.Desktop.Width;

            //ფორმის აუცილებელი თვისებები
            string formTitle = BDOSResources.getTranslate("IncomingPayment");

            Dictionary<string, object> formProperties = new Dictionary<string, object>();
            formProperties.Add("UniqueID", "BDOSINCDOC");
            formProperties.Add("BorderStyle", SAPbouiCOM.BoFormBorderStyle.fbs_Sizable);
            formProperties.Add("Title", formTitle);
            formProperties.Add("ClientWidth", formWidth);
            formProperties.Add("ClientHeight", formHeight);

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
                    oDataTable.Columns.Add("LineNum", SAPbouiCOM.BoFieldsType.ft_Integer, 50); //ინდექსი 
                    oDataTable.Columns.Add("DocEntry", SAPbouiCOM.BoFieldsType.ft_Integer, 50); //ენთრი
                    oDataTable.Columns.Add("DocNum", SAPbouiCOM.BoFieldsType.ft_Integer, 50); //ნომერი
                    oDataTable.Columns.Add("DocDate", SAPbouiCOM.BoFieldsType.ft_Date,50); //ნომერი
                    oDataTable.Columns.Add("Total", SAPbouiCOM.BoFieldsType.ft_Sum); //Default - Balance Due
                    oDataTable.Columns.Add("Currency", SAPbouiCOM.BoFieldsType.ft_Text, 50); //დოკუმენტის ვალუტა
                    oDataTable.Columns.Add("Project", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20);
                    oDataTable.Columns.Add("BlnkAgr", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20);
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
                       
                        else if (columnName == "BlnkAgr")
                        {
                            oColumn = oColumns.Add("BlnkAgr", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate(columnName);
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind(UID, columnName);
                            oLink = oColumn.ExtendedObject;
                            oLink.LinkedObjectType = "1250000025";
                        }
                        else if (columnName == "Project")
                        {
                            oColumn = oColumns.Add("Project", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate(columnName);
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind(UID, columnName);
                            oLink = oColumn.ExtendedObject;
                            oLink.LinkedObjectType = "63";
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

            SAPbouiCOM.Item oItem = oForm.Items.Item("InvoiceMTR");
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

        public static void resizeForm( SAPbouiCOM.Form oForm, out string errorText)
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

        
      
        public static void fillInvoicesMTR(  SAPbouiCOM.Form oForm, string DocEntries, out string errorText)
        {
            errorText = null;

            SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("InvoiceMTR");
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                
            string query = "";
           
            query = GetInvoicesMTRQuery(DocEntries);
           
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

                while (!oRecordSet.EoF)
                {
                    DocEntry = Convert.ToInt32(oRecordSet.Fields.Item("DocEntry").Value);
                    AgrNo = Convert.ToInt32(oRecordSet.Fields.Item("AgrNo").Value);
                    PrjCode = Convert.ToString(oRecordSet.Fields.Item("PrjCode").Value);
                    DocNum = Convert.ToInt32(oRecordSet.Fields.Item("DocNum").Value);
                    DocDate = oRecordSet.Fields.Item("DocDate").Value;
                    Total = Convert.ToDecimal(oRecordSet.Fields.Item("DocTotal").Value);
                    Currency = oRecordSet.Fields.Item("DocCurr").Value;

                    oDataTable.Rows.Add();
                    oDataTable.SetValue("LineNum", rowIndex, rowIndex + 1);
                    oDataTable.SetValue("DocEntry", rowIndex, DocEntry);
                    oDataTable.SetValue("DocNum", rowIndex, DocNum);
                    oDataTable.SetValue("DocDate", rowIndex, DocDate);
                    oDataTable.SetValue("BlnkAgr", rowIndex, AgrNo);
                    oDataTable.SetValue("Project", rowIndex, PrjCode);
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

        public static string GetInvoicesMTRQuery( string DocEntries)
        {
            string str = @"SELECT *
            	            
            FROM ORCT
            WHERE ""DocEntry"" in (" + DocEntries+ ")";

            return str;
        }


        public static void uiApp_ItemEvent(  string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            
            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_RESIZE & pVal.BeforeAction == false)
                {
                    oForm.Freeze(true);
                    resizeForm( oForm, out errorText);
                    oForm.Freeze(false);
                }

                
            }
        }
    }
}
