using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;
using System.Data;
using System.Runtime.InteropServices;

namespace BDO_Localisation_AddOn
{
    class BDOSVATReconcilationWizard
    {
        static DataTable ItemsDT = null;

        public static void createForm()
        {
            string errorText;
            int formHeight = Program.uiApp.Desktop.Height;
            int formWidth = Program.uiApp.Desktop.Width;

            //ფორმის აუცილებელი თვისებები
            Dictionary<string, object> formProperties = new Dictionary<string, object>();
            formProperties.Add("UniqueID", "BDOSReconWizz");
            formProperties.Add("BorderStyle", SAPbouiCOM.BoFormBorderStyle.fbs_Sizable);
            formProperties.Add("Title", BDOSResources.getTranslate("VATReconcilationWizard"));
            formProperties.Add("ClientWidth", formWidth);
            formProperties.Add("ClientHeight", formHeight);

            SAPbouiCOM.Form oForm;
            bool newForm;
            bool formExist = FormsB1.createForm(formProperties, out oForm, out newForm, out errorText);

            if (errorText != null)
            {
                return;
            }

            if (formExist)
            {
                if (newForm)
                {
                    Dictionary<string, object> formItems;
                    string itemName;

                    int width_s = 130;
                    int width_e = 130;
                    int left_s = 6;
                    int left_e = left_s + width_s + 20;
                    int height = 15;
                    int top = 10;

                    FormsB1.addChooseFromList(oForm, false, "2", "BusinessPartner_CFL");

                    //პირობის დადება ბიზნესპარტნიორის არჩევის სიაზე
                    SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item("BusinessPartner_CFL");
                    SAPbouiCOM.Conditions oCons = oCFL.GetConditions();
                    SAPbouiCOM.Condition oCon = oCons.Add();
                    oCon.Alias = "CardType";
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCon.CondVal = "C"; //მყიდველი
                    oCFL.SetConditions(oCons);

                    formItems = new Dictionary<string, object>();
                    itemName = "DocPsDtS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("DocumentPostingDate"));

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "DocPstDt";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_DATE);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Left", left_e);
                    formItems.Add("Width", width_e);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("ValueEx", new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1).AddMonths(1).AddDays(-1).ToString("yyyyMMdd"));

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    top = top + height + 1;

                    formItems = new Dictionary<string, object>();
                    itemName = "BPCodeS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("CardCode"));

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "BPCode"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("TableName", "");
                    formItems.Add("Length", 20);
                    formItems.Add("Alias", "BPCode");
                    formItems.Add("Bound", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("Left", left_e);
                    formItems.Add("Width", width_e);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("ChooseFromListUID", "BusinessPartner_CFL");
                    formItems.Add("ChooseFromListAlias", "CardCode");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "BPCodeLB"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    formItems.Add("Left", left_e - 20);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("LinkTo", "BPCode");
                    formItems.Add("LinkedObjectType", "2"); //Business Partner

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    top += 2 * height + 1;

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
                        throw new Exception(errorText);
                    }

                    itemName = "createDocB";
                    formItems = new Dictionary<string, object>();
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left_s + 65 + 2);
                    formItems.Add("Width", 65 * 2);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("CreateDocuments"));

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    top = top + height + 1;
                    left_s = 6;

                    formItems = new Dictionary<string, object>();
                    itemName = "InvoiceMTR"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
                    formItems.Add("Left", left_s);
                    formItems.Add("Height", 150);
                    formItems.Add("Top", top);
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("InvoiceMTR").Specific;
                    oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
                    SAPbouiCOM.Columns oColumns = oMatrix.Columns;
                    SAPbouiCOM.Column oColumn;

                    SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Add("InvoiceMTR");

                    oDataTable.Columns.Add("LineNum", SAPbouiCOM.BoFieldsType.ft_Integer, 50); //ინდექსი 
                    oDataTable.Columns.Add("CheckBox", SAPbouiCOM.BoFieldsType.ft_Text, 1); // 
                    oDataTable.Columns.Add("DocEntry", SAPbouiCOM.BoFieldsType.ft_Text, 50); //ენთრი
                    oDataTable.Columns.Add("DocNum", SAPbouiCOM.BoFieldsType.ft_Text, 50); //ნომერი
                    oDataTable.Columns.Add("DocDate", SAPbouiCOM.BoFieldsType.ft_Date, 50); //თარიღი
                    oDataTable.Columns.Add("CardCode", SAPbouiCOM.BoFieldsType.ft_Text, 50); //ენთრი
                    oDataTable.Columns.Add("CardName", SAPbouiCOM.BoFieldsType.ft_Text, 50); //ნომერი
                    oDataTable.Columns.Add("LicTradNum", SAPbouiCOM.BoFieldsType.ft_Text, 50); //ენთრი
                    oDataTable.Columns.Add("ReconSum", SAPbouiCOM.BoFieldsType.ft_Sum); //თანხა
                    oDataTable.Columns.Add("TransId", SAPbouiCOM.BoFieldsType.ft_Text, 50); //თანხა
                    oDataTable.Columns.Add("AlRcnSum", SAPbouiCOM.BoFieldsType.ft_Sum); //თანხა
                    oDataTable.Columns.Add("AlRcnVat", SAPbouiCOM.BoFieldsType.ft_Sum); //თანხა
                    oDataTable.Columns.Add("DocEntVT", SAPbouiCOM.BoFieldsType.ft_Text, 50); //ენთრი
                    oDataTable.Columns.Add("DocTotal", SAPbouiCOM.BoFieldsType.ft_Sum); //თანხა
                    oDataTable.Columns.Add("DocVtTotal", SAPbouiCOM.BoFieldsType.ft_Sum); //თანხა
                    oDataTable.Columns.Add("Error", SAPbouiCOM.BoFieldsType.ft_Text, 50); //ენთრი

                    string UID = "InvoiceMTR";
                    SAPbouiCOM.LinkedButton oLink;

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
                            oColumn.ValOff = "N";
                            oColumn.ValOn = "Y";
                            oColumn.DataBind.Bind(UID, columnName);
                        }
                        else if (columnName == "DocEntry")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("ARDownPaymentRequest") + " (" + BDOSResources.getTranslate(columnName) + ")";
                            oColumn.Editable = false;
                            oLink = oColumn.ExtendedObject;
                            oLink.LinkedObjectType = "203";
                            oColumn.DataBind.Bind(UID, columnName);
                        }
                        else if (columnName == "DocNum")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("ARDownPaymentRequest") + " (" + BDOSResources.getTranslate(columnName) + ")";
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind(UID, columnName);
                        }
                        else if (columnName == "DocEntVT")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("ARDownPaymentVATAccrual");
                            oColumn.Editable = false;
                            oLink = oColumn.ExtendedObject;
                            oLink.LinkedObjectType = "UDO_F_BDO_ARDPV_D";
                            oColumn.DataBind.Bind(UID, columnName);
                        }
                        else if (columnName == "CardCode")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("CardCode") + " (" + BDOSResources.getTranslate("Code") + ")";
                            oColumn.Editable = false;
                            oLink = oColumn.ExtendedObject;
                            oLink.LinkedObjectType = "2";
                            oColumn.DataBind.Bind(UID, columnName);
                        }
                        else if (columnName == "TransId")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("TransId");
                            oColumn.Editable = false;
                            oLink = oColumn.ExtendedObject;
                            oLink.LinkedObjectType = "30";
                            oColumn.DataBind.Bind(UID, columnName);
                        }
                        else if (columnName == "CardName")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("CardCode") + " (" + BDOSResources.getTranslate("Name") + ")";
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind(UID, columnName);
                        }
                        else if (columnName == "LicTradNum")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Tin");
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind(UID, columnName);
                        }
                        else if (columnName == "DocTotal")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Total");
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind(UID, columnName);
                        }
                        else if (columnName == "DocVtTotal")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("VatAmount");
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind(UID, columnName);
                        }
                        else if (columnName == "InDetail")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_PICTURE);
                            oColumn.Editable = false;
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
                }
                oForm.Visible = true;
                oForm.Select();
            }
        }

        public static void uiApp_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_CLOSE && pVal.BeforeAction)
                {
                    int answer = Program.uiApp.MessageBox(BDOSResources.getTranslate("CloseVATReconcilationWizard") + "?", 1, BDOSResources.getTranslate("Yes"), BDOSResources.getTranslate("No"), "");

                    if (answer != 1)
                        BubbleEvent = false;
                }

                else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_RESIZE && !pVal.BeforeAction)
                {
                    resizeForm(oForm);
                }

                else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                {
                    SAPbouiCOM.IChooseFromListEvent oCFLEvento = (SAPbouiCOM.IChooseFromListEvent)pVal;
                    chooseFromList(oForm, pVal, oCFLEvento);
                }

                else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK)
                {
                    if (!pVal.BeforeAction)
                    {
                        if (pVal.ItemUID == "checkB" || pVal.ItemUID == "unCheckB")
                            checkUncheck(oForm, pVal.ItemUID);
                        else if (pVal.ItemUID == "fillB")
                            fillMTRInvoice(oForm);
                        else if (pVal.ItemUID == "createDocB")
                            createPaymentDocuments(oForm);
                    }
                }

                //else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED)
                //{
                //    if (pVal.ItemUID == "InvoiceMTR" && pVal.ColUID == "DocEntry")
                //        matrixColumnSetLinkedObjectTypeInvoicesMTR(oForm, pVal);
                //}

                else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
                {
                    if (!pVal.BeforeAction)
                    {
                        if (pVal.ItemUID == "InvoiceMTR" && pVal.Row > 0)
                        {
                            setInvDocsMatrixRowBackColor(oForm, pVal.Row);
                        }
                    }
                }

                if (pVal.ItemChanged && !pVal.BeforeAction)
                {
                    if (pVal.ItemUID == "DocPstDt")
                    {
                        string docDateStr = oForm.DataSources.UserDataSources.Item("DocPstDt").ValueEx;
                        if (!string.IsNullOrEmpty(docDateStr))
                        {
                            DateTime endDateOp = DateTime.ParseExact(docDateStr, "yyyyMMdd", null);
                            endDateOp = new DateTime(endDateOp.Year, endDateOp.Month, 1).AddMonths(1).AddDays(-1);
                            oForm.DataSources.UserDataSources.Item("DocPstDt").ValueEx = endDateOp.ToString("yyyyMMdd");
                        }

                        fillMTRInvoice(oForm);
                    }
                }
            }
        }

        public static void setInvDocsMatrixRowBackColor(SAPbouiCOM.Form oForm, int row)
        {
            try
            {
                oForm.Freeze(true);

                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("InvoiceMTR").Specific;

                if (oMatrix.RowCount > 0)
                {
                    for (int i = 1; i <= oMatrix.RowCount; i++)
                    {
                        oMatrix.CommonSetting.SetRowBackColor(i, FormsB1.getLongIntRGB(231, 231, 231));
                    }
                    oMatrix.CommonSetting.SetRowBackColor(row, FormsB1.getLongIntRGB(255, 255, 153));
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        public static void setInvDocsMatrixRowCellColor(SAPbouiCOM.Form oForm)
        {
            try
            {
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("InvoiceMTR").Specific;

                if (oMatrix.RowCount > 0)
                {
                    for (int i = 1; i <= oMatrix.RowCount; i++)
                    {
                        if (oMatrix.GetCellSpecific("ReconSum", i).Value != oMatrix.GetCellSpecific("AlRcnSum", i).Value)
                            oMatrix.CommonSetting.SetRowFontColor(i, FormsB1.getLongIntRGB(255, 0, 0));
                        else
                            oMatrix.CommonSetting.SetRowFontColor(i, FormsB1.getLongIntRGB(0, 0, 0));
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private static void checkUncheck(SAPbouiCOM.Form oForm, string checkOperation)
        {
            try
            {
                oForm.Freeze(true);

                SAPbouiCOM.CheckBox oCheckBox;
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("InvoiceMTR").Specific;

                for (int j = 1; j <= oMatrix.RowCount; j++)
                {
                    oCheckBox = oMatrix.Columns.Item("CheckBox").Cells.Item(j).Specific;

                    oCheckBox.Checked = (checkOperation == "checkB");
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        //public static void matrixColumnSetLinkedObjectTypeInvoicesMTR(SAPbouiCOM.Form oForm, SAPbouiCOM.ItemEvent pVal)
        //{
        //    try
        //    {
        //        oForm.Freeze(true);

        //        if (pVal.ColUID == "DocEntry")
        //        {
        //            if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED && pVal.BeforeAction)
        //            {
        //                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("InvoiceMTR").Specific;

        //                SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("InvoiceMTR");
        //                string docType = oDataTable.GetValue("DocType", pVal.Row - 1);

        //                SAPbouiCOM.Column oColumn;

        //                if (docType == "18")
        //                {
        //                    oColumn = oMatrix.Columns.Item(pVal.ColUID);
        //                    SAPbouiCOM.LinkedButton oLink = oColumn.ExtendedObject;
        //                    oLink.LinkedObjectType = docType; //ARInvoice
        //                }
        //                if (docType == "204")
        //                {
        //                    oColumn = oMatrix.Columns.Item(pVal.ColUID);
        //                    SAPbouiCOM.LinkedButton oLink = oColumn.ExtendedObject;
        //                    oLink.LinkedObjectType = docType; //ARInvoice
        //                }
        //                else if (docType == "163")
        //                {
        //                    oColumn = oMatrix.Columns.Item(pVal.ColUID);
        //                    SAPbouiCOM.LinkedButton oLink = oColumn.ExtendedObject;
        //                    oLink.LinkedObjectType = docType; //ARCreditNote
        //                }
        //            }
        //        }
        //        else
        //        {

        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        throw new Exception(ex.Message);
        //    }
        //    finally
        //    {
        //        oForm.Freeze(false);
        //    }
        //}

        private static void createPaymentDocuments(SAPbouiCOM.Form oForm)
        {
            int answer = Program.uiApp.MessageBox(BDOSResources.getTranslate("CreatePaymentDocuments") + "?", 1, BDOSResources.getTranslate("Yes"), BDOSResources.getTranslate("No"), "");

            if (answer == 2)
                return;

            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("InvoiceMTR").Specific;

            string docDateStr = oForm.DataSources.UserDataSources.Item("DocPstDt").ValueEx;
            if (string.IsNullOrEmpty(docDateStr))
            {
                string errorText = BDOSResources.getTranslate("TheFollowingFieldsAreMandatory") +
                    " : \"" + oForm.Items.Item("DocPsDtS").Specific.caption + "\"";

                Program.uiApp.SetStatusBarMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short);
                return;
            }
            DateTime DocDate = Convert.ToDateTime(DateTime.ParseExact(docDateStr, "yyyyMMdd", CultureInfo.InvariantCulture));
            DateTime DocDateEE = (new DateTime(DocDate.Year, DocDate.Month, DateTime.DaysInMonth(DocDate.Year, DocDate.Month)));

            DataTable AccountTable = CommonFunctions.GetOACTTable();

            DataTable reLines = new DataTable();
            reLines.Columns.Add("month");
            reLines.Columns.Add("vatAccrl");
            reLines.Columns.Add("reconSum");
            reLines.Columns.Add("vatSum");

            DataTable jeLines = JournalEntry.JournalEntryTable();

            for (int row = 1; row <= oMatrix.RowCount; row++)
            {
                bool checkedLine = oMatrix.GetCellSpecific("CheckBox", row).Checked;
                if (checkedLine)
                {
                    jeLines.Rows.Clear();
                    reLines.Rows.Clear();

                    string DocEntVT = oMatrix.GetCellSpecific("DocEntVT", row).Value;

                    Dictionary<string, string> listAccounts = GetVatAcrualJornalEntry(DocEntVT);
                    string VatGrp = listAccounts["VatGroup"].ToString();

                    decimal VatRate = CommonFunctions.GetVatGroupRate(VatGrp, "");

                    string TransId = oMatrix.GetCellSpecific("TransId", row).Value;
                    if (TransId != "")
                    {
                        Program.uiApp.SetStatusBarMessage("დღგ უკვე გატარებულია", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                        continue;
                    }

                    string InvDocEntry = oMatrix.GetCellSpecific("DocEntry", row).Value;
                    decimal DocVtTotal = Convert.ToDecimal(oMatrix.GetCellSpecific("DocVtTotal", row).Value, CultureInfo.InvariantCulture);
                    decimal ReconSum = Convert.ToDecimal(oMatrix.GetCellSpecific("ReconSum", row).Value, CultureInfo.InvariantCulture);
                    decimal ReconSumVAT = CommonFunctions.roundAmountByGeneralSettings(ReconSum * VatRate / (100 + VatRate), "Sum");

                    DocVtTotal = Math.Min(DocVtTotal, ReconSumVAT);

                    JournalEntry.AddJournalEntryRow(AccountTable, jeLines, "OnlyCredit", listAccounts["CreditAccount"], listAccounts["DebitAccount"], DocVtTotal, 0, "", "", "", "", "", "", "", "", "");

                    JournalEntry.AddJournalEntryRow(AccountTable, jeLines, "OnlyDebit", listAccounts["CreditAccount"], listAccounts["DebitAccount"], DocVtTotal, 0, "", "", "", "", "", "", "", VatGrp, "");

                    DataRow reLinesRow = reLines.Rows.Add();
                    reLinesRow["month"] = DocDateEE.ToString("yyyyMMdd");
                    reLinesRow["vatAccrl"] = DocEntVT;
                    reLinesRow["reconSum"] = ReconSum.ToString(CultureInfo.InvariantCulture);
                    reLinesRow["vatSum"] = ReconSumVAT.ToString(CultureInfo.InvariantCulture);

                    string errorText = null;

                    try
                    {
                        JournalEntry.JrnEntry(DocEntVT, "Reconcilation", "Reconcilation ", DocDate, jeLines, out errorText);
                        addRecord(reLines, out errorText);

                        if (errorText != null)
                            return;
                    }
                    catch (Exception ex)
                    {
                        errorText = ex.Message;
                    }
                }
            }
            fillMTRInvoice(oForm);
        }

        private static Dictionary<string, string> GetVatAcrualJornalEntry(string DocEntry)
        {
            Dictionary<string, string> listValidValuesDict = new Dictionary<string, string>();

            string query = @"select  top 2 * 
            from ""JDT1""
            inner join ""OJDT"" on ""JDT1"".""TransId"" = ""OJDT"".""TransId"" and ""OJDT"".""Ref1"" = '" + DocEntry + @"'
                and ""OJDT"".""Ref2""  = 'UDO_F_BDO_ARDPV_D' 
            where ""OJDT"".""TransId"" not in 
            (select ""StornoToTr"" from ""OJDT"" where ""StornoToTr"" is not null)
            AND ""OJDT"".""StornoToTr"" Is NULL";

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRecordSet.DoQuery(query);

            while (!oRecordSet.EoF)
            {
                decimal debit = Convert.ToDecimal(oRecordSet.Fields.Item("Debit").Value, CultureInfo.InvariantCulture);
                string VatGroup = oRecordSet.Fields.Item("VatGroup").Value.ToString();

                if (VatGroup != "")
                {
                    listValidValuesDict.Add("VatGroup", VatGroup);
                }


                if (debit > 0)
                {
                    listValidValuesDict.Add("DebitAccount", oRecordSet.Fields.Item("Account").Value.ToString());
                }
                else
                {
                    listValidValuesDict.Add("CreditAccount", oRecordSet.Fields.Item("Account").Value.ToString());
                }
                oRecordSet.MoveNext();
            }

            return listValidValuesDict;
        }

        private static void chooseFromList(SAPbouiCOM.Form oForm, SAPbouiCOM.ItemEvent pVal, SAPbouiCOM.IChooseFromListEvent oCFLEvento)
        {
            try
            {
                oForm.Freeze(true);

                if (pVal.BeforeAction)
                {

                }
                else
                {
                    SAPbouiCOM.DataTable oDataTable = oCFLEvento.SelectedObjects;

                    if (oDataTable != null)
                    {
                        if (oCFLEvento.ChooseFromListUID == "BusinessPartner_CFL")
                        {
                            string CardCode = oDataTable.GetValue("CardCode", 0);
                            LanguageUtils.IgnoreErrors<string>(() => oForm.Items.Item("BPCode").Specific.Value = CardCode);

                            fillMTRInvoice(oForm);
                        }
                    }
                    else
                        fillMTRInvoice(oForm);
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        public static void createUDO(out string errorText)
        {
            errorText = null;

            string tableName = "BDOSRECWIZ";
            string description = "Reconcilation Wizard History";

            SAPbobsCOM.UserObjectsMD oUserObjectsMD = null;
            oUserObjectsMD = (SAPbobsCOM.UserObjectsMD)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);
            Marshal.ReleaseComObject(oUserObjectsMD);

            int result = UDO.addUserTable(tableName, description, SAPbobsCOM.BoUTBTableType.bott_NoObjectAutoIncrement, out errorText);

            if (result != 0)
            {
                return;
            }

            Dictionary<string, object> fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap = new Dictionary<string, object>(); // docType
            fieldskeysMap.Add("Name", "month");
            fieldskeysMap.Add("TableName", "BDOSRECWIZ");
            fieldskeysMap.Add("Description", "Month");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Date);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // docType
            fieldskeysMap.Add("Name", "vatAccrl");
            fieldskeysMap.Add("TableName", "BDOSRECWIZ");
            fieldskeysMap.Add("Description", "Vat accrual");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // docType
            fieldskeysMap.Add("Name", "reconSum");
            fieldskeysMap.Add("TableName", "BDOSRECWIZ");
            fieldskeysMap.Add("Description", "Reconcilation sum");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // docType
            fieldskeysMap.Add("Name", "vatSum");
            fieldskeysMap.Add("TableName", "BDOSRECWIZ");
            fieldskeysMap.Add("Description", "Vat sum");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);
        }

        public static void addRecord(DataTable reLines, out string errorText)
        {
            errorText = null;
            int returnCode;

            SAPbobsCOM.UserTable oUserTable = Program.oCompany.UserTables.Item("BDOSRECWIZ");
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            DataRow reLine;
            try
            {
                for (int i = 0; i < reLines.Rows.Count; i++)
                {
                    reLine = reLines.Rows[i];
                    string queryOPDF = @"delete from ""@BDOSRECWIZ""
                                    where ""U_month"" = '" + reLine["month"] + @"' and ""U_vatAccrl"" = " + reLine["vatAccrl"].ToString();
                    oRecordSet.DoQuery(queryOPDF);
                    oUserTable.UserFields.Fields.Item("U_month").Value = DateTime.ParseExact(reLine["month"].ToString(), "yyyyMMdd", CultureInfo.InvariantCulture);
                    oUserTable.UserFields.Fields.Item("U_vatAccrl").Value = reLine["vatAccrl"].ToString();
                    oUserTable.UserFields.Fields.Item("U_reconSum").Value = Convert.ToDouble(reLine["reconSum"], CultureInfo.InvariantCulture);
                    oUserTable.UserFields.Fields.Item("U_vatSum").Value = Convert.ToDouble(reLine["vatSum"], CultureInfo.InvariantCulture);
                    returnCode = oUserTable.Add();

                    if (returnCode != 0)
                    {
                        int errCode;
                        string errMsg;

                        Program.oCompany.GetLastError(out errCode, out errMsg);
                        errorText = "Error description : " + errMsg + "! Code : " + errCode;
                    }
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
            finally
            {
                Marshal.ReleaseComObject(oUserTable);
                Marshal.ReleaseComObject(oRecordSet);
            }
        }

        public static void fillMTRInvoice(SAPbouiCOM.Form oForm)
        {
            string docDateStr = oForm.DataSources.UserDataSources.Item("DocPstDt").ValueEx;
            if (string.IsNullOrEmpty(docDateStr))
            {
                string errorText = BDOSResources.getTranslate("TheFollowingFieldsAreMandatory") +
                    " : \"" + oForm.Items.Item("DocPsDtS").Specific.caption + "\"";

                Program.uiApp.SetStatusBarMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short);
                return;
            }

            DateTime date = Convert.ToDateTime(DateTime.ParseExact(docDateStr, "yyyyMMdd", CultureInfo.InvariantCulture));
            DateTime prevDate = date.AddMonths(-1);

            string dateES = new DateTime(date.Year, date.Month, 1).ToString("yyyyMMdd");
            string dateEE = new DateTime(date.Year, date.Month, DateTime.DaysInMonth(date.Year, date.Month)).ToString("yyyyMMdd");

            string prevdateES = new DateTime(prevDate.Year, prevDate.Month, 1).ToString("yyyyMMdd");
            string prevdateEE = new DateTime(prevDate.Year, prevDate.Month, DateTime.DaysInMonth(prevDate.Year, prevDate.Month)).ToString("yyyyMMdd");

            string cardCode = oForm.DataSources.UserDataSources.Item("BPCode").ValueEx;

            SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("InvoiceMTR");
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            StringBuilder query = new StringBuilder();
            query.Append("SELECT \"ODPI\".\"DocEntry\", \n");
            query.Append("       \"ODPI\".\"DocNum\", \n");
            query.Append("       \"ODPI\".\"DocDate\", \n");
            query.Append("       \"ODPI\".\"CardCode\", \n");
            query.Append("       \"ODPI\".\"CardName\", \n");
            query.Append("       \"ITR1\".\"ReconSum\", \n");
            query.Append("       IFNULL(\"OJDT1\".\"Debit\", 0)                            AS \"AlRcnVat\", \n");
            query.Append("       IFNULL(\"OJDT1\".\"TransId\", 0)                          AS \"TransId\", \n");
            query.Append("       IFNULL(\"OCRD\".\"LicTradNum\", '')                       AS \"LicTradNum\", \n");
            query.Append("       IFNULL(\"OJDT1\".\"BaseSum\", 0)                          AS \"U_reconSum\", \n");
            query.Append("       \"@BDOSARDV\".\"U_GrsAmnt\" - IFNULL(\"OJDT\".\"BaseSum\", 0) AS \"DocTotal\", \n");
            query.Append("       \"@BDOSARDV\".\"U_VatAmount\" - IFNULL(\"OJDT\".\"Debit\", 0) AS \"DocVAtTotal\", \n");
            query.Append("       IFNULL(\"@BDOSARDV\".\"DocEntry\", 0)                     AS \"DocEntVT\", \n");
            query.Append("       IFNULL(\"@BDOSARDV\".\"DocNum\", 0)                       AS \"DocNumVT\" \n");
            query.Append("FROM   (SELECT \"IncomingPayment\".\"DownPmntEntry\"        AS \"DocEntry\", \n");
            query.Append("               Sum(\"InternalReconciliation\".\"ReconSum\") AS \"ReconSum\" \n");
            query.Append("        FROM   (SELECT \"SrcObjAbs\"            AS \"IncPmntEntry\", \n");
            query.Append("                       Sum(\"ITR1\".\"ReconSum\") AS \"ReconSum\" \n");
            query.Append("                FROM   \"ITR1\" \n");
            query.Append("                       INNER JOIN \"OITR\" \n");
            query.Append("                               ON \"ITR1\".\"ReconNum\" = \"OITR\".\"ReconNum\" \n");
            query.Append("                                  AND \"OITR\".\"ReconDate\" >= '" + dateES + "' \n");
            query.Append("                                  AND \"OITR\".\"ReconDate\" <= '" + dateEE + "' \n");
            query.Append("                                  AND \"OITR\".\"Canceled\" <> 'C' \n");
            query.Append("                WHERE  \"ITR1\".\"SrcObjTyp\" = 24 \n");
            if (!string.IsNullOrEmpty(cardCode))
                query.Append("                       AND \"ITR1\".\"ShortName\" = '" + cardCode + "' \n");
            query.Append("                GROUP  BY \"SrcObjAbs\") AS \"InternalReconciliation\" \n");
            query.Append("               INNER JOIN (SELECT \"RCT2\".\"DocNum\"   AS \"IncPmntEntry\", \n");
            query.Append("                                  \"RCT2\".\"DocEntry\" AS \"DownPmntEntry\" \n");
            query.Append("                           FROM   \"RCT2\" \n");
            query.Append("                                  INNER JOIN \"ORCT\" \n");
            query.Append("                                          ON \"ORCT\".\"DocEntry\" = \"RCT2\".\"DocNum\" \n");
            query.Append("                                             AND \"ORCT\".\"Canceled\" = 'N' \n");
            if (!string.IsNullOrEmpty(cardCode))
                query.Append("                                             AND \"ORCT\".\"CardCode\" = '" + cardCode + "' \n");
            query.Append("                                             AND \"RCT2\".\"InvType\" = 203) AS \n");
            query.Append("                          \"IncomingPayment\" \n");
            query.Append("                       ON \"InternalReconciliation\".\"IncPmntEntry\" = \n");
            query.Append("                          \"IncomingPayment\".\"IncPmntEntry\" \n");
            query.Append("        GROUP  BY \"IncomingPayment\".\"DownPmntEntry\") AS \"ITR1\" \n");
            query.Append("       INNER JOIN \"ODPI\" \n");
            query.Append("               ON \"ITR1\".\"DocEntry\" = \"ODPI\".\"DocEntry\" \n");
            query.Append("                  AND \"DocStatus\" = 'C' \n");
            if (!string.IsNullOrEmpty(cardCode))
                query.Append("                  AND \"ODPI\".\"CardCode\" = '" + cardCode + "' \n");
            query.Append("       INNER JOIN \"@BDOSARDV\" \n");
            query.Append("               ON \"ODPI\".\"DocEntry\" = \"@BDOSARDV\".\"U_baseDoc\" \n");
            query.Append("                  AND \"@BDOSARDV\".\"U_baseDocT\" = 203 \n");
            query.Append("                  AND \"@BDOSARDV\".\"Canceled\" = 'N' \n");
            query.Append("                  AND \"@BDOSARDV\".\"U_DocDate\" <= '" + dateEE + "' \n");
            query.Append("       INNER JOIN \"OCRD\" \n");
            query.Append("               ON \"ODPI\".\"CardCode\" = \"OCRD\".\"CardCode\" \n");
            query.Append("       left JOIN (SELECT \"OJDT\".\"Ref1\", \n");
            query.Append("                         Sum(\"OJDT\".\"BaseSum\") AS \"BaseSum\", \n");
            query.Append("                         Sum(\"OJDT\".\"Debit\")   AS \"Debit\" \n");
            query.Append("                  FROM   (SELECT \"OJDT\".\"TransId\"                       AS \n");
            query.Append("                                 \"TransId\", \n");
            query.Append("                                 \"OJDT\".\"Ref1\", \n");
            query.Append("                                 Sum(\"JDT1\".\"BaseSum\" + \"JDT1\".\"Debit\") AS \n");
            query.Append("                                 \"BaseSum\", \n");
            query.Append("                                 Sum(\"JDT1\".\"Debit\")                    AS \n");
            query.Append("                                 \"Debit\" \n");
            query.Append("                          FROM   \"JDT1\" \n");
            query.Append("                                 INNER JOIN \"OJDT\" \n");
            query.Append("                                         ON \"JDT1\".\"TransId\" = \"OJDT\".\"TransId\" \n");
            query.Append("                                            AND \"OJDT\".\"TaxDate\" <= '" + dateEE + "' \n");
            query.Append("                                            AND \"OJDT\".\"Ref2\" = 'Reconcilation' \n");
            query.Append("                                            AND \"JDT1\".\"VatGroup\" <> '' \n");
            query.Append("                                            AND \"OJDT\".\"StornoToTr\" IS NULL \n");
            query.Append("                          GROUP  BY \"OJDT\".\"Ref1\", \n");
            query.Append("                                    \"OJDT\".\"TransId\" \n");
            query.Append("                          UNION ALL \n");
            query.Append("                          SELECT \"OJDT\".\"StornoToTr\"     AS \"TransId\", \n");
            query.Append("                                 \"OJDT\".\"Ref1\", \n");
            query.Append("                                 Sum(\"JDT1\".\"BaseSum\" + CASE WHEN \n");
            query.Append("                                     \"JDT1\".\"Credit\">0 \n");
            query.Append("                                     THEN \n");
            query.Append("                                     -\"JDT1\".\"Credit\" \n");
            query.Append("                                     ELSE \n");
            query.Append("                                     \"JDT1\".\"Debit\" END) AS \"BaseSum\", \n");
            query.Append("                                 Sum(CASE \n");
            query.Append("                                       WHEN \"JDT1\".\"Credit\" > 0 THEN \n");
            query.Append("                                       -\"JDT1\".\"Credit\" \n");
            query.Append("                                       ELSE \"JDT1\".\"Debit\" \n");
            query.Append("                                     END)                AS \"Credit\" \n");
            query.Append("                          FROM   \"JDT1\" \n");
            query.Append("                                 INNER JOIN \"OJDT\" \n");
            query.Append("                                         ON \"JDT1\".\"TransId\" = \"OJDT\".\"TransId\" \n");
            query.Append("                                            AND \"OJDT\".\"TaxDate\" <= '" + dateEE + "' \n");
            query.Append("                                            AND \"OJDT\".\"Ref2\" = 'Reconcilation' \n");
            query.Append("                                            AND \"JDT1\".\"VatGroup\" <> '' \n");
            query.Append("                                            AND not \"OJDT\".\"StornoToTr\" IS NULL \n");
            query.Append("                          GROUP  BY \"OJDT\".\"Ref1\", \n");
            query.Append("                                    \"OJDT\".\"StornoToTr\") AS \"OJDT\" \n");
            query.Append("                  GROUP  BY \"OJDT\".\"Ref1\" \n");
            query.Append("                  Having Sum(\"OJDT\".\"Debit\") > 0) AS \"OJDT\" \n");
            query.Append("              ON \"@BDOSARDV\".\"DocEntry\" = \"OJDT\".\"Ref1\" \n");
            query.Append("       left JOIN (SELECT \"OJDT\".\"TransId\"      AS \"TransId\", \n");
            query.Append("                         \"OJDT\".\"Ref1\", \n");
            query.Append("                         Sum(\"OJDT\".\"BaseSum\") AS \"BaseSum\", \n");
            query.Append("                         Sum(\"OJDT\".\"Debit\")   AS \"Debit\" \n");
            query.Append("                  FROM   (SELECT \"OJDT\".\"TransId\"        AS \"TransId\", \n");
            query.Append("                                 \"OJDT\".\"Ref1\", \n");
            query.Append("                                 Sum(\"JDT1\".\"BaseSum\" + CASE WHEN \n");
            query.Append("                                     \"JDT1\".\"Credit\">0 \n");
            query.Append("                                     THEN \n");
            query.Append("                                     -\"JDT1\".\"Credit\" \n");
            query.Append("                                     ELSE \n");
            query.Append("                                     \"JDT1\".\"Debit\" END) AS \"BaseSum\", \n");
            query.Append("                                 Sum(\"JDT1\".\"Debit\")     AS \"Debit\" \n");
            query.Append("                          FROM   \"JDT1\" \n");
            query.Append("                                 INNER JOIN \"OJDT\" \n");
            query.Append("                                         ON \"JDT1\".\"TransId\" = \"OJDT\".\"TransId\" \n");
            query.Append("                                            AND \"OJDT\".\"TaxDate\" >= '" + dateES + "' \n");
            query.Append("                                            AND \"OJDT\".\"TaxDate\" <= '" + dateEE + "' \n");
            query.Append("                                            AND \"OJDT\".\"Ref2\" = 'Reconcilation' \n");
            query.Append("                                            AND \"JDT1\".\"VatGroup\" <> '' \n");
            query.Append("                                            AND \"OJDT\".\"StornoToTr\" IS NULL \n");
            query.Append("                          GROUP  BY \"OJDT\".\"Ref1\", \n");
            query.Append("                                    \"OJDT\".\"TransId\" \n");
            query.Append("                          UNION ALL \n");
            query.Append("                          SELECT \"OJDT\".\"StornoToTr\"     AS \"TransId\", \n");
            query.Append("                                 \"OJDT\".\"Ref1\", \n");
            query.Append("                                 Sum(\"JDT1\".\"BaseSum\" + CASE WHEN \n");
            query.Append("                                     \"JDT1\".\"Credit\">0 \n");
            query.Append("                                     THEN \n");
            query.Append("                                     -\"JDT1\".\"Credit\" \n");
            query.Append("                                     ELSE \n");
            query.Append("                                     \"JDT1\".\"Debit\" END) AS \"BaseSum\", \n");
            query.Append("                                 Sum(CASE \n");
            query.Append("                                       WHEN \"JDT1\".\"Credit\" > 0 THEN \n");
            query.Append("                                       -\"JDT1\".\"Credit\" \n");
            query.Append("                                       ELSE \"JDT1\".\"Debit\" \n");
            query.Append("                                     END)                AS \"Credit\" \n");
            query.Append("                          FROM   \"JDT1\" \n");
            query.Append("                                 INNER JOIN \"OJDT\" \n");
            query.Append("                                         ON \"JDT1\".\"TransId\" = \"OJDT\".\"TransId\" \n");
            query.Append("                                            AND \"OJDT\".\"TaxDate\" >= '" + dateES + "' \n");
            query.Append("                                            AND \"OJDT\".\"TaxDate\" <= '" + dateEE + "' \n");
            query.Append("                                            AND \"OJDT\".\"Ref2\" = 'Reconcilation' \n");
            query.Append("                                            AND \"JDT1\".\"VatGroup\" <> '' \n");
            query.Append("                                            AND not \"OJDT\".\"StornoToTr\" IS NULL \n");
            query.Append("                          GROUP  BY \"OJDT\".\"Ref1\", \n");
            query.Append("                                    \"OJDT\".\"StornoToTr\") AS \"OJDT\" \n");
            query.Append("                  GROUP  BY \"OJDT\".\"TransId\", \n");
            query.Append("                            \"OJDT\".\"Ref1\" \n");
            query.Append("                  Having Sum(\"OJDT\".\"Debit\") > 0) AS \"OJDT1\" \n");
            query.Append("              ON \"@BDOSARDV\".\"DocEntry\" = \"OJDT1\".\"Ref1\" \n");
            query.Append("ORDER  BY \"ODPI\".\"DocDate\" ASC");

            if (Program.oCompany.DbServerType != SAPbobsCOM.BoDataServerTypes.dst_HANADB)
            {
                query = query.Replace("IFNULL", "ISNULL");
            }

            oRecordSet.DoQuery(query.ToString());

            oDataTable.Rows.Clear();
            try
            {
                int rowIndex = 0;
                string DocEntry;
                string DocNum;
                string DocDate;
                string DocEntVT;
                string CardCode;
                string CardName;
                string LicTradNum;
                string TransId;
                decimal DocTotal = 0;
                decimal DocVatTotal = 0;
                decimal ReconSum = 0;
                decimal AlRcnSum = 0;
                decimal AlRcnVat = 0;

                ItemsDT = new DataTable();
                ItemsDT.Columns.Add("DocEntry");
                ItemsDT.Columns.Add("ItemCode");
                ItemsDT.Columns.Add("Dscptn");
                ItemsDT.Columns.Add("GrsAmnt");
                ItemsDT.Columns.Add("VatGrp");
                ItemsDT.Columns.Add("VatAmnt");

                while (!oRecordSet.EoF)
                {
                    DocEntry = oRecordSet.Fields.Item("DocEntry").Value == 0 ? "" : oRecordSet.Fields.Item("DocEntry").Value.ToString();
                    DocNum = oRecordSet.Fields.Item("DocNum").Value == 0 ? "" : oRecordSet.Fields.Item("DocNum").Value.ToString();
                    DocDate = oRecordSet.Fields.Item("DocDate").Value.ToString("yyyyMMdd");
                    DocEntVT = oRecordSet.Fields.Item("DocEntVT").Value == 0 ? "" : oRecordSet.Fields.Item("DocEntVT").Value.ToString();
                    decimal allowableDeviation = Convert.ToDecimal(CommonFunctions.getOADM("U_BDOSAllDev").ToString(), CultureInfo.InvariantCulture);
                    TransId = oRecordSet.Fields.Item("TransId").Value == 0 ? "" : oRecordSet.Fields.Item("TransId").Value.ToString();
                    DocTotal = Convert.ToDecimal(oRecordSet.Fields.Item("DocTotal").Value, CultureInfo.InvariantCulture);
                    DocVatTotal = Convert.ToDecimal(oRecordSet.Fields.Item("DocVatTotal").Value, CultureInfo.InvariantCulture);
                    AlRcnVat = Convert.ToDecimal(oRecordSet.Fields.Item("AlRcnVat").Value, CultureInfo.InvariantCulture);
                    AlRcnSum = Convert.ToDecimal(oRecordSet.Fields.Item("U_reconSum").Value, CultureInfo.InvariantCulture);
                    ReconSum = Convert.ToDecimal(oRecordSet.Fields.Item("ReconSum").Value, CultureInfo.InvariantCulture);
                    CardCode = oRecordSet.Fields.Item("CardCode").Value;
                    CardName = oRecordSet.Fields.Item("CardName").Value;
                    LicTradNum = oRecordSet.Fields.Item("LicTradNum").Value;

                    if (Math.Abs(DocTotal) <= allowableDeviation)
                        DocTotal = 0;
                    if (Math.Abs(AlRcnSum - ReconSum) <= allowableDeviation)
                        AlRcnSum = ReconSum;

                    if (TransId == "" && ReconSum == 0)
                    {
                        oRecordSet.MoveNext();
                        continue;
                    }

                    oDataTable.Rows.Add();
                    oDataTable.SetValue("LineNum", rowIndex, rowIndex + 1);
                    oDataTable.SetValue("CheckBox", rowIndex, "N");
                    oDataTable.SetValue("DocEntry", rowIndex, DocEntry);
                    oDataTable.SetValue("DocNum", rowIndex, DocNum);
                    oDataTable.SetValue("DocEntVT", rowIndex, DocEntVT);
                    oDataTable.SetValue("DocDate", rowIndex, DocDate);
                    oDataTable.SetValue("CardCode", rowIndex, CardCode);
                    oDataTable.SetValue("CardName", rowIndex, CardName);
                    oDataTable.SetValue("LicTradNum", rowIndex, LicTradNum);
                    oDataTable.SetValue("TransId", rowIndex, TransId);
                    oDataTable.SetValue("DocTotal", rowIndex, Convert.ToDouble(DocTotal, CultureInfo.InvariantCulture));
                    oDataTable.SetValue("DocVtTotal", rowIndex, Convert.ToDouble(DocVatTotal, CultureInfo.InvariantCulture));
                    oDataTable.SetValue("ReconSum", rowIndex, Convert.ToDouble(ReconSum, CultureInfo.InvariantCulture));
                    oDataTable.SetValue("AlRcnVat", rowIndex, Convert.ToDouble(AlRcnVat, CultureInfo.InvariantCulture));
                    oDataTable.SetValue("AlRcnSum", rowIndex, Convert.ToDouble(AlRcnSum, CultureInfo.InvariantCulture));

                    oRecordSet.MoveNext();
                    rowIndex++;
                }

                oForm.Freeze(true);
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("InvoiceMTR").Specific;
                oMatrix.Clear();
                oMatrix.LoadFromDataSource();
                oForm.Update();

                setInvDocsMatrixRowCellColor(oForm);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                oForm.Freeze(false);
                Marshal.ReleaseComObject(oRecordSet);
            }
        }

        public static void addMenus()
        {
            try
            {
                SAPbouiCOM.MenuItem fatherMenuItem = Program.uiApp.Menus.Item("2048");

                // Add a pop-up menu item
                SAPbouiCOM.MenuCreationParams oCreationPackage = (SAPbouiCOM.MenuCreationParams)Program.uiApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                oCreationPackage.Checked = false;
                oCreationPackage.Enabled = true;
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "BDOSReconWizz";
                oCreationPackage.String = BDOSResources.getTranslate("VATReconcilationWizard");
                oCreationPackage.Position = fatherMenuItem.SubMenus.Count - 1;

                SAPbouiCOM.MenuItem menuItem = fatherMenuItem.SubMenus.AddEx(oCreationPackage);
            }
            catch
            {

            }
        }

        public static void resizeForm(SAPbouiCOM.Form oForm)
        {
            try
            {
                oForm.Freeze(true);

                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("InvoiceMTR").Specific;
                int mtrWidth = oForm.ClientWidth;
                oForm.Items.Item("InvoiceMTR").Width = mtrWidth;
                oForm.Items.Item("InvoiceMTR").Height = oForm.ClientHeight - 25;
                int columnsCount = oMatrix.Columns.Count - 2;
                oMatrix.Columns.Item("LineNum").Width = 19;
                oMatrix.Columns.Item("CheckBox").Width = 19;
                mtrWidth -= 38;
                mtrWidth /= columnsCount;

                foreach (SAPbouiCOM.Column column in oMatrix.Columns)
                {
                    if (column.UniqueID == "LineNum" || column.UniqueID == "CheckBox")
                        continue;
                    column.Width = mtrWidth;
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }
    }
}


