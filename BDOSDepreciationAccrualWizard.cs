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
    class BDOSDepreciationAccrualWizard
    {
        public static void createForm()
        {
            string errorText;
            int formHeight = Program.uiApp.Desktop.Height;
            int formWidth = Program.uiApp.Desktop.Width;

            //ფორმის აუცილებელი თვისებები
            Dictionary<string, object> formProperties = new Dictionary<string, object>();
            formProperties.Add("UniqueID", "BDOSDepAccrForm");
            formProperties.Add("BorderStyle", SAPbouiCOM.BoFormBorderStyle.fbs_Sizable);
            formProperties.Add("Title", BDOSResources.getTranslate("DepreciationAccruingWizard"));
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

                    formItems = new Dictionary<string, object>();
                    itemName = "DeprMonthS"; //10 characters
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
                        throw new Exception(errorText);
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "DeprMonth";
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

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    top += height + 10;

                    formItems = new Dictionary<string, object>();
                    itemName = "InvDepr"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("Length", 1);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_OPTION_BUTTON);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("Retirement"));
                    formItems.Add("ValOn", "Y");
                    formItems.Add("ValOff", "N");
                    formItems.Add("Value", 1);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "StckDepr"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("Length", 1);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_OPTION_BUTTON);
                    formItems.Add("Left", left_e);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("Depreciation"));
                    formItems.Add("GroupWith", "InvDepr");
                    formItems.Add("ValOn", "Y");
                    formItems.Add("ValOff", "N");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    top += height + 10;

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

                    top += height + 1;

                    formItems = new Dictionary<string, object>();
                    itemName = "ItemsMTR"; //10 characters
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

                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("ItemsMTR").Specific;
                    oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
                    SAPbouiCOM.Columns oColumns = oMatrix.Columns;
                    SAPbouiCOM.Column oColumn;

                    SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Add("ItemMTRTmp");
                    oDataTable.Columns.Add("DocType", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 50);
                    oDataTable.Columns.Add("DocEntry", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 50);
                    oDataTable.Columns.Add("Project", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 50);
                    oDataTable.Columns.Add("ItemCode", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 50);
                    oDataTable.Columns.Add("DistNumber", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 50);
                    oDataTable.Columns.Add("UseLife", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 50);
                    oDataTable.Columns.Add("Quantity", SAPbouiCOM.BoFieldsType.ft_Quantity);
                    oDataTable.Columns.Add("DeprAmt", SAPbouiCOM.BoFieldsType.ft_Sum);
                    oDataTable.Columns.Add("DepcDoc", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 50);
                    oDataTable.Columns.Add("AlrDeprAmt", SAPbouiCOM.BoFieldsType.ft_Sum);

                    oDataTable = oForm.DataSources.DataTables.Add("ItemsMTR");
                    oDataTable.Columns.Add("LineNum", SAPbouiCOM.BoFieldsType.ft_Integer, 50); //ინდექსი 
                    oDataTable.Columns.Add("DocType", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 50);
                    oDataTable.Columns.Add("DocEntry", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 50);
                    oDataTable.Columns.Add("Project", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 50);
                    oDataTable.Columns.Add("ItemGrp", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 50);
                    oDataTable.Columns.Add("ItemCode", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 50);
                    oDataTable.Columns.Add("ItemName", SAPbouiCOM.BoFieldsType.ft_Text, 50);
                    oDataTable.Columns.Add("DistNumber", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 50);
                    oDataTable.Columns.Add("UseLife", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 50);
                    oDataTable.Columns.Add("AlrDeprAmt", SAPbouiCOM.BoFieldsType.ft_Sum);
                    oDataTable.Columns.Add("Quantity", SAPbouiCOM.BoFieldsType.ft_Quantity);
                    oDataTable.Columns.Add("APCost", SAPbouiCOM.BoFieldsType.ft_Sum);
                    oDataTable.Columns.Add("NtBookVal", SAPbouiCOM.BoFieldsType.ft_Sum);
                    oDataTable.Columns.Add("DeprAmt", SAPbouiCOM.BoFieldsType.ft_Sum);
                    oDataTable.Columns.Add("CurMnthAmt", SAPbouiCOM.BoFieldsType.ft_Sum);
                    oDataTable.Columns.Add("DepcDoc", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 50);

                    string UID = "ItemsMTR";
                    SAPbouiCOM.LinkedButton oLink;

                    for (int count = 0; count < oDataTable.Columns.Count; count++)
                    {
                        var column = oDataTable.Columns.Item(count);
                        string columnName = column.Name;

                        if (columnName == "LineNum")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = "#";
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind(UID, columnName);
                        }
                        else if (columnName == "Project")
                        {
                            oColumn = oColumns.Add("Project", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Project");
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind(UID, columnName);
                            oLink = oColumn.ExtendedObject;
                            oLink.LinkedObjectType = "63";
                        }
                        else if (columnName == "ItemGrp")
                        {
                            oColumn = oColumns.Add("ItemGrp", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("ItmsGrpCod");
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind(UID, columnName);
                            oLink = oColumn.ExtendedObject;
                            oLink.LinkedObjectType = "52";
                        }
                        else if (columnName == "ItemCode")
                        {
                            oColumn = oColumns.Add("ItemCode", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("ItemCode");
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind(UID, columnName);
                            oLink = oColumn.ExtendedObject;
                            oLink.LinkedObjectType = "4";
                        }
                        else if (columnName == "DepcDoc")
                        {
                            oColumn = oColumns.Add("DepcDoc", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("DepcDoc");
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind(UID, columnName);
                            oLink = oColumn.ExtendedObject;
                            oLink.LinkedObjectType = "UDO_F_BDOSDEPACR_D";
                        }
                        else if (columnName == "DocEntry")
                        {
                            oColumn = oColumns.Add("DocEntry", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("DocEntry");
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind(UID, columnName);
                        }
                        else if (columnName == "DistNumber")
                        {
                            oColumn = oColumns.Add("DistNumber", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("DistNumber");
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

        public static void matrixColumnSetLinkedObjectTypeInvoicesMTR(SAPbouiCOM.Form oForm, SAPbouiCOM.ItemEvent pVal)
        {
            try
            {
                oForm.Freeze(true);

                if (pVal.BeforeAction)
                {
                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("ItemsMTR").Specific;

                    SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("ItemsMTR");
                    string docType = oDataTable.GetValue("DocType", pVal.Row - 1);

                    SAPbouiCOM.Column oColumn;

                    if (docType == "13")
                    {
                        oColumn = oMatrix.Columns.Item(pVal.ColUID);
                        SAPbouiCOM.LinkedButton oLink = oColumn.ExtendedObject;
                        oLink.LinkedObjectType = docType; //Invoice object
                    }
                    else if (docType == "60")
                    {
                        oColumn = oMatrix.Columns.Item(pVal.ColUID);
                        SAPbouiCOM.LinkedButton oLink = oColumn.ExtendedObject;
                        oLink.LinkedObjectType = docType; //Goods Issue object
                    }
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

        public static void uiApp_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_CLOSE && pVal.BeforeAction)
                {
                    int answer = Program.uiApp.MessageBox(BDOSResources.getTranslate("CloseDepreciationAccruingWizard") + "?", 1, BDOSResources.getTranslate("Yes"), BDOSResources.getTranslate("No"), "");

                    if (answer != 1)
                        BubbleEvent = false;
                }

                else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_RESIZE && !pVal.BeforeAction)
                {
                    resizeForm(oForm);
                }

                else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK)
                {
                    if (!pVal.BeforeAction)
                    {
                        if (pVal.ItemUID == "fillB")
                            fillMTRItems(oForm);
                        else if (pVal.ItemUID == "createDocB")
                        {
                            CreateDocuments(oForm);
                            fillMTRItems(oForm);
                        }
                    }
                }

                else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED)
                {
                    if (pVal.ItemUID == "ItemsMTR" && pVal.ColUID == "DocEntry")
                        matrixColumnSetLinkedObjectTypeInvoicesMTR(oForm, pVal);
                }

                else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
                {
                    if (!pVal.BeforeAction)
                    {
                        if (pVal.ItemUID == "StckDepr" || pVal.ItemUID == "InvDepr")
                            setVisibleFormItems(oForm);
                    }
                }

                else if (pVal.ItemChanged && !pVal.BeforeAction)
                {
                    if (pVal.ItemUID == "DeprMonth")
                        itemPressed(oForm);
                }
            }
        }

        private static void itemPressed(SAPbouiCOM.Form oForm)
        {
            try
            {
                oForm.Freeze(true);

                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("ItemsMTR").Specific;

                string dateStr = oForm.DataSources.UserDataSources.Item("DeprMonth").ValueEx;
                if (!string.IsNullOrEmpty(dateStr))
                {
                    DateTime accrMnth = DateTime.ParseExact(dateStr, "yyyyMMdd", null);
                    accrMnth = new DateTime(accrMnth.Year, accrMnth.Month, 1);
                    accrMnth = accrMnth.AddMonths(1).AddDays(-1);

                    oForm.DataSources.UserDataSources.Item("DeprMonth").ValueEx = accrMnth.ToString("yyyyMMdd");
                }
                if (oMatrix.RowCount > 0)
                    oMatrix.Clear();
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

        private static int CreateDocument(SAPbouiCOM.Form oForm, DateTime AccrMnth, DateTime PostingDate)
        {
            string errorText = null;

            SAPbobsCOM.CompanyService oCompanyService = Program.oCompany.GetCompanyService();
            SAPbobsCOM.GeneralService oGeneralService = oCompanyService.GetGeneralService("UDO_F_BDOSDEPACR_D");
            SAPbobsCOM.GeneralData oGeneralData = (SAPbobsCOM.GeneralData)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);

            oGeneralData.SetProperty("U_AccrMnth", AccrMnth);
            oGeneralData.SetProperty("U_DocDate", PostingDate);

            SAPbobsCOM.GeneralDataCollection oChildren = oGeneralData.Child("BDOSDEPAC1");
            SAPbouiCOM.DataTable DepreciationLines = oForm.DataSources.DataTables.Item("ItemMTRTmp");

            for (int i = 0; i < DepreciationLines.Rows.Count; i++)
            {
                SAPbobsCOM.GeneralData oChild = oChildren.Add();
                oChild.SetProperty("U_ItemCode", DepreciationLines.GetValue("ItemCode", i));
                oChild.SetProperty("U_DistNumber", DepreciationLines.GetValue("DistNumber", i));
                oChild.SetProperty("U_BDOSUsLife", DepreciationLines.GetValue("UseLife", i));
                oChild.SetProperty("U_Project", DepreciationLines.GetValue("Project", i));
                oChild.SetProperty("U_Quantity", DepreciationLines.GetValue("Quantity", i));
                oChild.SetProperty("U_DeprAmt", DepreciationLines.GetValue("DeprAmt", i));
                oChild.SetProperty("U_InvEntry", DepreciationLines.GetValue("DocEntry", i));
                oChild.SetProperty("U_InvType", DepreciationLines.GetValue("DocType", i));
                oChild.SetProperty("U_AlrDeprAmt", DepreciationLines.GetValue("AlrDeprAmt", i));
            }
            int docEntry = 0;

            try
            {
                CommonFunctions.StartTransaction();

                var response = oGeneralService.Add(oGeneralData);
                docEntry = response.GetProperty("DocEntry");

                if (docEntry > 0)
                {
                    DataTable JrnLinesDT = BDOSDepreciationAccrualDocument.createAdditionalEntries(null, oGeneralData, 0, PostingDate, "", null, null);

                    BDOSDepreciationAccrualDocument.JrnEntry(docEntry.ToString(), docEntry.ToString(), PostingDate, JrnLinesDT, "", out errorText);

                    if (errorText != null)
                    {
                        CommonFunctions.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                        Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("DocumentNotCreated") + ". " + BDOSResources.getTranslate("ReasonIs") + ": " + errorText, SAPbouiCOM.BoMessageTime.bmt_Short);
                    }
                    else
                    {
                        Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("DocumentCreatedSuccesfully") + ": " + docEntry.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                        CommonFunctions.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                    }
                }
            }
            catch (Exception Ex)
            {
                Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("DocumentNotCreated") + ". " + BDOSResources.getTranslate("ReasonIs") + ": " + Ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short);
            }
            return docEntry;
        }

        private static void CreateDocuments(SAPbouiCOM.Form oForm)
        {
            bool isInvoice = oForm.Items.Item("InvDepr").Specific.Selected;

            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("ItemsMTR").Specific;
            SAPbouiCOM.DataTable depreciationLinesTmp = oForm.DataSources.DataTables.Item("ItemMTRTmp");
            depreciationLinesTmp.Rows.Clear();
            SAPbouiCOM.DataTable depreciationLines = oForm.DataSources.DataTables.Item("ItemsMTR");

            string dateStr = oForm.DataSources.UserDataSources.Item("DeprMonth").ValueEx;
            if (string.IsNullOrEmpty(dateStr))
            {
                string errorText = BDOSResources.getTranslate("TheFollowingFieldsAreMandatory") +
                    " : \"" + oForm.Items.Item("DeprMonthS").Specific.caption + "\"";

                Program.uiApp.SetStatusBarMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short);
                return;
            }

            DateTime accrMnth = DateTime.ParseExact(dateStr, "yyyyMMdd", null);

            int newRow = 0;
            for (int i = 0; i < oMatrix.RowCount; i++)
            {
                string itemCode = depreciationLines.GetValue("ItemCode", i);
                string distNumber = depreciationLines.GetValue("DistNumber", i);
                string project = depreciationLines.GetValue("Project", i);
                double curMnthAmt = depreciationLines.GetValue("CurMnthAmt", i);

                if (curMnthAmt == 0)
                {
                    newRow = depreciationLinesTmp.Rows.Count;
                    depreciationLinesTmp.Rows.Add();
                    depreciationLinesTmp.SetValue("ItemCode", newRow, itemCode);
                    depreciationLinesTmp.SetValue("DistNumber", newRow, distNumber);
                    depreciationLinesTmp.SetValue("UseLife", newRow, depreciationLines.GetValue("UseLife", i));
                    depreciationLinesTmp.SetValue("Project", newRow, project);
                    depreciationLinesTmp.SetValue("Quantity", newRow, depreciationLines.GetValue("Quantity", i));
                    depreciationLinesTmp.SetValue("DeprAmt", newRow, depreciationLines.GetValue("DeprAmt", i));
                    depreciationLinesTmp.SetValue("AlrDeprAmt", newRow, depreciationLines.GetValue("AlrDeprAmt", i));

                    if (isInvoice)
                    {
                        depreciationLinesTmp.SetValue("DocEntry", newRow, depreciationLines.GetValue("DocEntry", i));
                        depreciationLinesTmp.SetValue("DocType", newRow, depreciationLines.GetValue("DocType", i));

                        SAPbobsCOM.BoObjectTypes DocType;

                        if (depreciationLines.GetValue("DocType", i) == "13") //Invoice object
                            DocType = SAPbobsCOM.BoObjectTypes.oInvoices;
                        else
                            DocType = SAPbobsCOM.BoObjectTypes.oInventoryGenExit;

                        SAPbobsCOM.Documents oInvoice = Program.oCompany.GetBusinessObject(DocType);
                        DateTime DocDate = new DateTime();
                        if (oInvoice.GetByKey(depreciationLines.GetValue("DocEntry", i)))
                            DocDate = oInvoice.DocDate;

                        CreateDocument(oForm, accrMnth, DocDate);

                        depreciationLinesTmp.Rows.Clear();
                    }
                }
                else
                {
                    Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("DocumentAlreadyCreated") + "! " + itemCode + " - " + distNumber + " - " + project, SAPbouiCOM.BoMessageTime.bmt_Short);
                    continue;
                }

                //StringBuilder query = new StringBuilder();
                ////query.Append("SELECT * \n");
                ////query.Append("FROM \n");
                //query.Append("(SELECT TOP 1 \n");
                //query.Append("\"@BDOSDEPACR\".\"DocEntry\", \n");
                //query.Append("\"@BDOSDEPACR\".\"U_DocDate\" \n");
                //query.Append("FROM \"@BDOSDEPACR\" \n");
                //query.Append("INNER JOIN \"@BDOSDEPAC1\" ON \"@BDOSDEPACR\".\"DocEntry\" = \"@BDOSDEPAC1\".\"DocEntry\" \n");
                //query.Append("WHERE \"U_DocDate\" <= '" + dateStr + "' \n");
                //query.Append("AND \"Canceled\" = 'N' \n");
                //query.Append("AND \"@BDOSDEPAC1\".\"U_ItemCode\" = '" + itemCode + "' \n");
                //query.Append("AND \"@BDOSDEPAC1\".\"U_DistNumber\" = '" + distNumber + "' \n");
                //query.Append("AND \"@BDOSDEPAC1\".\"U_Project\" = '" + project + "' \n");
                //query.Append("ORDER BY \n");
                //query.Append("\"U_DocDate\" DESC, \n");
                //query.Append("\"DocEntry\" DESC) \n");
                ////query.Append("WHERE ABS((MONTH('" + dateStr + "') - MONTH(\"U_DocDate\")))>1");

                //SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                //oRecordSet.DoQuery(query.ToString());
                //if (!oRecordSet.EoF)
                //{
                //    if (Math.Abs(accrMnth.Month - oRecordSet.Fields.Item("U_DocDate").Value.Month) > 1)
                //    {
                //        depreciationLinesTmp.Rows.Clear();
                //        Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("CreateDocumentForThePreviousMonths") + "! " + itemCode + " - " + distNumber + " - " + project, SAPbouiCOM.BoMessageTime.bmt_Short);
                //        Marshal.ReleaseComObject(oRecordSet);
                //        continue;
                //    }
                //}
            }

            if (!isInvoice && depreciationLinesTmp.Rows.Count > 0)
                CreateDocument(oForm, accrMnth, accrMnth);
        }

        public static void resizeForm(SAPbouiCOM.Form oForm)
        {
            try
            {
                oForm.Freeze(true);

                int mtrWidth = oForm.ClientWidth;
                oForm.Items.Item("ItemsMTR").Width = mtrWidth;
                oForm.Items.Item("ItemsMTR").Height = oForm.ClientHeight - 25;
                FormsB1.resetWidthMatrixColumns(oForm, "ItemsMTR", "LineNum", mtrWidth);
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

        public static string BatchDepreciaionQuery(DateTime DeprMonth, string ItemCodes, string BatchNumbers, string WhsCode, bool isInvoice = false)
        {



            string query = @"select 
                            ((""FinTable"".""APCost"" * ""DeprDocs"".""DeprDocAmnt"")/""FinTable"".""UseLife"") as ""AlrDeprAmt"",
                            ""DeprDocs"".""DeprDocAmnt"",
                            ""FinTable"".""LocCode"","
                            +
                            (isInvoice ? @" ""FinTable"".""DocEntry"", ""FinTable"".""DocType"", " : "")
                            +
                             @"""FinTable"".""Project"",
                             ""FinTable"".""ItemGrp"",
                             ""FinTable"".""ItemCode"",
	 	                     ""FinTable"".""ItemName"",
                             ""FinTable"".""UseLife"" ,
	                         ""FinTable"".""DistNumber"",
                             ""FinTable"".""InDate"",
	                         ""FinTable"".""APCost"",
                             ""DepcAccInvoice"".""DepcDoc"",
                             ""DepcAcc"".""DeprAmt"" as ""DeprAmt"","
                            +
                            (isInvoice ? @"""DepcAccInvoice"".""CurrDeprAmt"" as ""CurrDeprAmt""," : @"""CurrDepcAcc"".""CurrDeprAmt"" as ""CurrDeprAmt"", ""CurrDepcAcc"".""FutureDeprAmt"" as ""FutureDeprAmt"",")
                            +
                            @"""DepcAcc"".""DeprQty"" as ""DeprQty"", 
	                         sum(""FinTable"".""NtBookVal"") as ""NtBookVal"",
	                         sum(""FinTable"".""Quantity"") as ""Quantity""

                             from (select distinct
	                         ""OIVL"".""LocCode"","
                            +
                            (isInvoice ? @" case when ""OBVL"".""BaseDocEn""= 0 
                                                then ""OBVL"".""DocEntry""
                                                else ""OBVL"".""BaseDocEn"" end as ""DocEntry"", ""OBVL"".""DocType"", " : "")
                            +
                             @"""OWHS"".""U_BDOSPrjCod"" as ""Project"",
                             ""OITM"".""ItmsGrpCod"" as ""ItemGrp"", 	                         
                             ""OBVL"".""ItemCode"",
	 	                     ""OITM"".""ItemName"",
                             ""OITM"".""U_BDOSUsLife"" as ""UseLife"",
	                         ""OBVL"".""DistNumber"",
                             ""OBTN"".""InDate"",
	                         ""OBTN"".""CostTotal"" / ""OBTN"".""Quantity"" as ""APCost"",
                             
	                         ""OBTN"".""CostTotal""/""OBTN"".""Quantity"" * ""OBVL"".""Quantity""*case when ""OBVL"".""TransValue"">0 then 1 else -1 end as ""NtBookVal"",
	                         ""OBVL"".""Quantity""*case when ""OBVL"".""TransValue"">0 then 1 else -1 end  as ""Quantity"" 
                        from ""OBVL""
                        
                        inner join ""OBTN"" on ""OBTN"".""DistNumber"" = ""OBVL"".""DistNumber"" and ""OBTN"".""ItemCode"" = ""OBVL"".""ItemCode"" and ""OBTN"".""Quantity"">0
                        and #ItemFilter#
                        and #BatchFilter#
                        and  ADD_MONTHS(NEXT_DAY(LAST_DAY(""OBTN"".""InDate"")),-1)<= ADD_MONTHS(NEXT_DAY(LAST_DAY('" + DeprMonth.ToString("yyyyMMdd") + @"')),-1)
                        inner join ""OITM"" on  ""OBVL"".""ItemCode"" = ""OITM"".""ItemCode""
                        inner join ""OITB"" on  ""OITB"".""ItmsGrpCod"" = ""OITM"".""ItmsGrpCod"" and ""OITB"".""U_BDOSFxAs""='Y'
                        inner join ""OIVL"" on ""OIVL"".""ItemCode"" = ""OBVL"".""ItemCode"" 
                        and ""OIVL"".""DocDate"" <='" + DeprMonth.ToString("yyyyMMdd") + @"'                        
                        and ""OIVL"".""CreatedBy"" = ""OBVL"".""DocEntry"" 
                        and ""OIVL"".""TransType"" = ""OBVL"".""DocType"" 
                        and case when ""OIVL"".""ParentID"" = -1 
                        then ""OIVL"".""ParentID"" 
                        else ""OIVL"".""TransType"" 
                        END = ""OBVL"".""BaseType"" 
                        inner join ""OWHS"" on ""OWHS"".""WhsCode"" = ""OIVL"".""LocCode"""
                         +
                            (isInvoice ? @" where (""OBVL"".""DocType""= 13 or ""OBVL"".""DocType""= 60) and LAST_DAY(""OIVL"".""DocDate"") = '" + DeprMonth.ToString("yyyyMMdd") + "'" : "")
                            + @"

                        ) as ""FinTable""
	                         left  join (select
                        ""@BDOSDEPAC1"".""U_ItemCode"",
                        ""@BDOSDEPAC1"".""U_DistNumber"",
                        SUM(case when ISNULL(""@BDOSDEPAC1"".""U_Quantity"",0)=0 then 0 else ""@BDOSDEPAC1"".""U_DeprAmt""/""@BDOSDEPAC1"".""U_Quantity"" end) as ""DeprAmt"",

                        SUM(case when ""@BDOSDEPACR"".""U_AccrMnth"" = '" + DeprMonth.ToString("yyyyMMdd") + @"' 
                        then ""@BDOSDEPAC1"".""U_DeprAmt""
	                    else 0
                        end) as ""CurrDeprAmt"",

                        SUM(case when ""@BDOSDEPACR"".""U_AccrMnth"" > '" + DeprMonth.ToString("yyyyMMdd") + @"' 
                        then ""@BDOSDEPAC1"".""U_DeprAmt""
	                    else 0
                        end) as ""FutureDeprAmt"",

	                    SUM(""@BDOSDEPAC1"".""U_Quantity"") as ""DeprQty""
                        from ""@BDOSDEPAC1"" 
                        inner join   ""@BDOSDEPACR"" on ""@BDOSDEPACR"".""DocEntry"" = ""@BDOSDEPAC1"".""DocEntry"" and ""@BDOSDEPACR"".""Canceled"" = 'N' and ISNULL(""@BDOSDEPAC1"".""U_InvEntry"",'')=''
                        
                        group by 
                        ""@BDOSDEPAC1"".""U_ItemCode"",
                        ""@BDOSDEPAC1"".""U_DistNumber"" ) as ""DepcAcc"" 
                        on ""DepcAcc"".""U_ItemCode"" = ""FinTable"".""ItemCode"" 
                        and ""DepcAcc"".""U_DistNumber"" = ""FinTable"".""DistNumber""
------------------

                        left  join (select
                        ""@BDOSDEPAC1"".""U_Project"",
                        ""@BDOSDEPAC1"".""U_ItemCode"",
                        ""@BDOSDEPAC1"".""U_DistNumber"",
                        SUM(case when ISNULL(""@BDOSDEPAC1"".""U_Quantity"",0)=0 then 0 else ""@BDOSDEPAC1"".""U_DeprAmt""/""@BDOSDEPAC1"".""U_Quantity"" end) as ""DeprAmt"",

                        SUM(case when ""@BDOSDEPACR"".""U_AccrMnth"" = '" + DeprMonth.ToString("yyyyMMdd") + @"' 
                        then ""@BDOSDEPAC1"".""U_DeprAmt""
	                    else 0
                        end) as ""CurrDeprAmt"",

                        SUM(case when ""@BDOSDEPACR"".""U_AccrMnth"" > '" + DeprMonth.ToString("yyyyMMdd") + @"' 
                        then ""@BDOSDEPAC1"".""U_DeprAmt""
	                    else 0
                        end) as ""FutureDeprAmt""
                        from ""@BDOSDEPAC1"" 
                        inner join   ""@BDOSDEPACR"" on ""@BDOSDEPACR"".""DocEntry"" = ""@BDOSDEPAC1"".""DocEntry"" and ""@BDOSDEPACR"".""Canceled"" = 'N' and ISNULL(""@BDOSDEPAC1"".""U_InvEntry"",'')=''
                        
                        group by 
                        ""@BDOSDEPAC1"".""U_Project"",
                        ""@BDOSDEPAC1"".""U_ItemCode"",
                        ""@BDOSDEPAC1"".""U_DistNumber"" ) as ""CurrDepcAcc"" 
                        on ""CurrDepcAcc"".""U_ItemCode"" = ""FinTable"".""ItemCode"" 
                        and ""CurrDepcAcc"".""U_DistNumber"" = ""FinTable"".""DistNumber""
                        and ""CurrDepcAcc"".""U_Project"" = ""FinTable"".""Project""  


------------------

                        left join (select 
                        ""@BDOSDEPAC1"".""U_ItemCode"",
                        ""@BDOSDEPAC1"".""U_DistNumber"",count(distinct ""@BDOSDEPAC1"".""DocEntry"") as ""DeprDocAmnt""
                        from ""@BDOSDEPAC1""
                        inner join   ""@BDOSDEPACR"" on ""@BDOSDEPACR"".""DocEntry"" = ""@BDOSDEPAC1"".""DocEntry"" and ""@BDOSDEPACR"".""Canceled"" = 'N'
                        group by
                        ""@BDOSDEPAC1"".""U_ItemCode"",
                        ""@BDOSDEPAC1"".""U_DistNumber"") as ""DeprDocs"" 
                        on ""DeprDocs"".""U_ItemCode"" = ""FinTable"".""ItemCode""
                        and ""DeprDocs"".""U_DistNumber"" = ""FinTable"".""DistNumber""



------------------
                        left  join (select
                        ""@BDOSDEPACR"".""DocEntry"" as ""DepcDoc"",
                        ""@BDOSDEPAC1"".""U_InvEntry"",
                        ""@BDOSDEPAC1"".""U_InvType"",
                        
                        ""@BDOSDEPAC1"".""U_ItemCode"",
                        ""@BDOSDEPAC1"".""U_DistNumber"",
                        SUM(case when ISNULL(""@BDOSDEPAC1"".""U_Quantity"",0)=0 then 0 else ""@BDOSDEPAC1"".""U_DeprAmt""/""@BDOSDEPAC1"".""U_Quantity"" end) as ""DeprAmt"",

                        SUM(case when ""@BDOSDEPACR"".""U_AccrMnth"" = '" + DeprMonth.ToString("yyyyMMdd") + @"' 
                        then ""@BDOSDEPAC1"".""U_DeprAmt""
	                    else 0
                        end) as ""CurrDeprAmt"",

                        SUM(case when ""@BDOSDEPACR"".""U_AccrMnth"" > '" + DeprMonth.ToString("yyyyMMdd") + @"' 
                        then ""@BDOSDEPAC1"".""U_DeprAmt""
	                    else 0
                        end) as ""FutureDeprAmt"",

	                    SUM(""@BDOSDEPAC1"".""U_Quantity"") as ""DeprQty""
                        from ""@BDOSDEPAC1"" 
                        inner join ""@BDOSDEPACR"" on ""@BDOSDEPACR"".""DocEntry"" = ""@BDOSDEPAC1"".""DocEntry"" and ""@BDOSDEPACR"".""Canceled"" = 'N' and ""@BDOSDEPACR"".""U_AccrMnth"" = '" + DeprMonth.ToString("yyyyMMdd") + @"' and "
                        +
                        (isInvoice ? @" ISNULL(""@BDOSDEPAC1"".""U_InvEntry"",'')<>'' " : @" ISNULL(""@BDOSDEPAC1"".""U_InvEntry"",'')='' ")
                        +
                        @"group by 
                        ""@BDOSDEPACR"".""DocEntry"",
                        ""@BDOSDEPAC1"".""U_InvEntry"",
                        ""@BDOSDEPAC1"".""U_InvType"",
                        ""@BDOSDEPAC1"".""U_ItemCode"",
                        ""@BDOSDEPAC1"".""U_DistNumber"" ) as ""DepcAccInvoice""
                        on ""DepcAccInvoice"".""U_ItemCode"" = ""FinTable"".""ItemCode"" 
                        and ""DepcAccInvoice"".""U_DistNumber"" = ""FinTable"".""DistNumber""
                         "
                         +
                         (isInvoice ? @" and ""DepcAccInvoice"".""U_InvEntry"" =  ""FinTable"".""DocEntry"" and   ""DepcAccInvoice"".""U_InvType"" = ""FinTable"".""DocType"" " : "")
                         +
                            @"group by ""FinTable"".""LocCode"","
                            +
                            (isInvoice ? @" ""FinTable"".""DocEntry"", ""FinTable"".""DocType"", " : "")
                            +
                             @"""FinTable"".""Project"",
                                (""FinTable"".""APCost"" * ""DeprDocs"".""DeprDocAmnt"")/""FinTable"".""UseLife"",
                                ""DeprDocs"".""DeprDocAmnt"",
                             ""FinTable"".""ItemGrp"",
                             ""FinTable"".""ItemCode"",
	 	                     ""FinTable"".""ItemName"",
                             ""FinTable"".""UseLife"" ,
	                         ""FinTable"".""DistNumber"",
                             ""FinTable"".""InDate"",
	                         ""FinTable"".""APCost"",
                             ""DepcAccInvoice"".""DepcDoc"",
                             ""DepcAcc"".""DeprAmt"","
                            +
                            (isInvoice ? @"""DepcAccInvoice"".""CurrDeprAmt""," : @"""CurrDepcAcc"".""CurrDeprAmt"", ""CurrDepcAcc"".""FutureDeprAmt"",")
                            +
                            @"""DepcAcc"".""DeprQty""";



            if (ItemCodes == "")
            {
                query = query.Replace("#ItemFilter#", "1=1");
            }
            else
            {
                query = query.Replace("#ItemFilter#", @"""OBTN"".""ItemCode"" in (" + ItemCodes + @")");
            }

            if (BatchNumbers == "")
            {
                query = query.Replace("#BatchFilter#", "1=1");
            }
            else
            {
                query = query.Replace("#BatchFilter#", @"""OBTN"".""DistNumber"" in (" + BatchNumbers + @")");
            }

            if (Program.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
            {

                query = query.Replace("ISNULL", "IFNULL");
            }

            return query;
        }

        public static void fillMTRItems(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("ItemsMTR");
            SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("ItemsMTR").Specific;
            oDataTable.Rows.Clear();

            string dateStr = oForm.DataSources.UserDataSources.Item("DeprMonth").ValueEx;
            if (string.IsNullOrEmpty(dateStr))
            {
                string errorText = BDOSResources.getTranslate("TheFollowingFieldsAreMandatory") +
                    " : \"" + oForm.Items.Item("DeprMonthS").Specific.caption + "\"";

                Program.uiApp.SetStatusBarMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short);
                return;
            }
            DateTime deprMonth = Convert.ToDateTime(DateTime.ParseExact(dateStr, "yyyyMMdd", CultureInfo.InvariantCulture));
            bool isInvoice = oForm.Items.Item("InvDepr").Specific.Selected;

            string query = BatchDepreciaionQuery(deprMonth, "", "", "", isInvoice);

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRecordSet.DoQuery(query);
            int rowIndex = 0;
            int i = 0;
            while (!oRecordSet.EoF)
            {
                DateTime InDateStart = oRecordSet.Fields.Item("InDate").Value;
                DateTime InDateEnd = InDateStart.AddMonths(oRecordSet.Fields.Item("UseLife").Value);
                InDateEnd = new DateTime(InDateEnd.Year, InDateEnd.Month, 1);
                InDateEnd = InDateEnd.AddMonths(1).AddDays(-1);
                DateTime AccrMnth = InDateStart;
                AccrMnth = new DateTime(AccrMnth.Year, AccrMnth.Month, 1);
                AccrMnth = AccrMnth.AddMonths(1).AddDays(-1);

                decimal Quantity = Convert.ToDecimal(oRecordSet.Fields.Item("Quantity").Value);
                Quantity = Quantity * (isInvoice ? -1 : 1);
                i++;

                if (deprMonth > InDateEnd || deprMonth <= AccrMnth || Quantity == 0)
                {
                    oRecordSet.MoveNext();
                    continue;
                }

                decimal CurrDeprAmt = Convert.ToDecimal(oRecordSet.Fields.Item("CurrDeprAmt").Value, CultureInfo.InvariantCulture);
                int monthsApart = 12 * (InDateEnd.Year - deprMonth.Year) + (InDateEnd.Month - deprMonth.Month) + 1;
                monthsApart = Math.Abs(monthsApart);

                decimal AlrDeprAmt = 0;
                //AlrDeprAmt = Convert.ToDecimal(oRecordSet.Fields.Item("DeprAmt").Value)  * Quantity;
                AlrDeprAmt = Convert.ToDecimal(oRecordSet.Fields.Item("AlrDeprAmt").Value) * Quantity;
                AlrDeprAmt -= CurrDeprAmt;

                decimal NtBookVal = Convert.ToDecimal(oRecordSet.Fields.Item("APCost").Value * Convert.ToDouble(Quantity)) - AlrDeprAmt;
                decimal DeprAmt = NtBookVal / monthsApart;

                oDataTable.Rows.Add();
                oDataTable.SetValue("LineNum", rowIndex, rowIndex + 1);
                oDataTable.SetValue("Project", rowIndex, oRecordSet.Fields.Item("Project").Value);
                oDataTable.SetValue("ItemGrp", rowIndex, oRecordSet.Fields.Item("ItemGrp").Value);
                oDataTable.SetValue("ItemCode", rowIndex, oRecordSet.Fields.Item("ItemCode").Value);
                oDataTable.SetValue("ItemName", rowIndex, oRecordSet.Fields.Item("ItemName").Value);
                oDataTable.SetValue("DistNumber", rowIndex, oRecordSet.Fields.Item("DistNumber").Value);
                if (oRecordSet.Fields.Item("DepcDoc").Value != 0)
                    oDataTable.SetValue("DepcDoc", rowIndex, oRecordSet.Fields.Item("DepcDoc").Value);
                oDataTable.SetValue("UseLife", rowIndex, monthsApart);
                oDataTable.SetValue("Quantity", rowIndex, Convert.ToDouble(Quantity));
                oDataTable.SetValue("APCost", rowIndex, oRecordSet.Fields.Item("APCost").Value * Convert.ToDouble(Quantity));
                oDataTable.SetValue("AlrDeprAmt", rowIndex, Convert.ToDouble(AlrDeprAmt));
                oDataTable.SetValue("NtBookVal", rowIndex, Convert.ToDouble(NtBookVal));
                if (CurrDeprAmt > 0)
                    oDataTable.SetValue("DeprAmt", rowIndex, 0);
                else
                    oDataTable.SetValue("DeprAmt", rowIndex, Convert.ToDouble(DeprAmt));
                oDataTable.SetValue("CurMnthAmt", rowIndex, Convert.ToDouble(CurrDeprAmt));
                if (isInvoice)
                {
                    oDataTable.SetValue("DocEntry", rowIndex, oRecordSet.Fields.Item("DocEntry").Value);
                    oDataTable.SetValue("DocType", rowIndex, oRecordSet.Fields.Item("DocType").Value);
                }

                rowIndex++;
                oRecordSet.MoveNext();
            }

            oForm.Freeze(true);
            oMatrix.Clear();
            oMatrix.LoadFromDataSource();
            oForm.Update();
            oForm.Freeze(false);
        }

        public static void addMenus()
        {
            try
            {
                SAPbouiCOM.MenuItem fatherMenuItem = Program.uiApp.Menus.Item("9201");

                // Add a pop-up menu item
                SAPbouiCOM.MenuCreationParams oCreationPackage = (SAPbouiCOM.MenuCreationParams)Program.uiApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                oCreationPackage.Checked = false;
                oCreationPackage.Enabled = true;
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "BDOSDepAccrForm";
                oCreationPackage.String = BDOSResources.getTranslate("DepreciationAccruingWizard");
                oCreationPackage.Position = fatherMenuItem.SubMenus.Count - 1;

                SAPbouiCOM.MenuItem menuItem = fatherMenuItem.SubMenus.AddEx(oCreationPackage);
            }
            catch
            {

            }
        }

        public static void setVisibleFormItems(SAPbouiCOM.Form oForm)
        {
            try
            {
                oForm.Freeze(true);

                bool isRetirement = oForm.Items.Item("InvDepr").Specific.Selected;

                SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("ItemsMTR").Specific;

                if (isRetirement)
                {
                    oMatrix.Columns.Item("DocType").Visible = true;
                    oMatrix.Columns.Item("DocEntry").Visible = true;
                }
                else
                {
                    oMatrix.Columns.Item("DocType").Visible = false;
                    oMatrix.Columns.Item("DocEntry").Visible = false;
                }

                int mtrWidth = oForm.ClientWidth;
                oForm.Items.Item("ItemsMTR").Width = mtrWidth;
                FormsB1.resetWidthMatrixColumns(oForm, "ItemsMTR", "LineNum", mtrWidth);
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
