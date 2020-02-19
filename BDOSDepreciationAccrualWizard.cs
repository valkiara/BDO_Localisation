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
        const int clientHeight = 600;
        const int clientWidth = 800;

        public static void createForm()
        {
            string errorText;

            //ფორმის აუცილებელი თვისებები
            Dictionary<string, object> formProperties = new Dictionary<string, object>();
            formProperties.Add("UniqueID", "BDOSDepAccrForm");
            formProperties.Add("BorderStyle", SAPbouiCOM.BoFormBorderStyle.fbs_Sizable);
            formProperties.Add("Title", BDOSResources.getTranslate("DepreciationAccruingWizard"));
            formProperties.Add("ClientWidth", clientWidth);
            formProperties.Add("ClientHeight", clientHeight);

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

                    //top += height + 10;

                    //formItems = new Dictionary<string, object>();
                    //itemName = "InvDepr"; //10 characters
                    //formItems.Add("isDataSource", true);
                    //formItems.Add("DataSource", "UserDataSources");
                    //formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    //formItems.Add("Length", 1);
                    //formItems.Add("TableName", "");
                    //formItems.Add("Alias", itemName);
                    //formItems.Add("Bound", true);
                    //formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_OPTION_BUTTON);
                    //formItems.Add("Left", left_s);
                    //formItems.Add("Width", width_s);
                    //formItems.Add("Top", top);
                    //formItems.Add("Height", height);
                    //formItems.Add("UID", itemName);
                    //formItems.Add("Caption", BDOSResources.getTranslate("Retirement"));
                    //formItems.Add("ValOn", "Y");
                    //formItems.Add("ValOff", "N");
                    ////formItems.Add("Value", 1);

                    //FormsB1.createFormItem(oForm, formItems, out errorText);
                    //if (errorText != null)
                    //{
                    //    throw new Exception(errorText);
                    //}

                    //formItems = new Dictionary<string, object>();
                    //itemName = "StckDepr"; //10 characters
                    //formItems.Add("isDataSource", true);
                    //formItems.Add("DataSource", "UserDataSources");
                    //formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    //formItems.Add("Length", 1);
                    //formItems.Add("TableName", "");
                    //formItems.Add("Alias", itemName);
                    //formItems.Add("Bound", true);
                    //formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_OPTION_BUTTON);
                    //formItems.Add("Left", left_e);
                    //formItems.Add("Width", width_s);
                    //formItems.Add("Top", top);
                    //formItems.Add("Height", height);
                    //formItems.Add("UID", itemName);
                    //formItems.Add("Caption", BDOSResources.getTranslate("Depreciation"));
                    //formItems.Add("GroupWith", "InvDepr");
                    //formItems.Add("ValOn", "Y");
                    //formItems.Add("ValOff", "N");
                    //formItems.Add("Value", 2);

                    //FormsB1.createFormItem(oForm, formItems, out errorText);
                    //if (errorText != null)
                    //{
                    //    throw new Exception(errorText);
                    //}

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

                    SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Add("ItemsMTR");
                    oDataTable.Columns.Add("LineNum", SAPbouiCOM.BoFieldsType.ft_Integer);
                    oDataTable.Columns.Add("DistNumber", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 36);
                    oDataTable.Columns.Add("WhsCode", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 8);
                    oDataTable.Columns.Add("PrjCode", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20);
                    oDataTable.Columns.Add("ItemCode", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 50);
                    oDataTable.Columns.Add("ItemName", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);
                    oDataTable.Columns.Add("ItmsGrpCod", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 6);
                    oDataTable.Columns.Add("ItmsGrpNam", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20);
                    oDataTable.Columns.Add("UsefulLife", SAPbouiCOM.BoFieldsType.ft_Integer);
                    oDataTable.Columns.Add("RemainingLife", SAPbouiCOM.BoFieldsType.ft_Integer);
                    oDataTable.Columns.Add("Quantity", SAPbouiCOM.BoFieldsType.ft_Quantity);
                    oDataTable.Columns.Add("PurchasePrice", SAPbouiCOM.BoFieldsType.ft_Price);
                    oDataTable.Columns.Add("PurchaseCost", SAPbouiCOM.BoFieldsType.ft_Sum);
                    oDataTable.Columns.Add("NetBookValue", SAPbouiCOM.BoFieldsType.ft_Sum);
                    oDataTable.Columns.Add("DepreciationAmt", SAPbouiCOM.BoFieldsType.ft_Sum);
                    oDataTable.Columns.Add("AlreadyDepreciatedAmt", SAPbouiCOM.BoFieldsType.ft_Sum);
                    oDataTable.Columns.Add("DepreciationDocEntry", SAPbouiCOM.BoFieldsType.ft_Integer);
                    oDataTable.Columns.Add("AccumulatedDepreciationAmt", SAPbouiCOM.BoFieldsType.ft_Sum);

                    SAPbouiCOM.DataTable oDataTableTmp = oForm.DataSources.DataTables.Add("ItemMTRTmp");
                    oDataTableTmp.CopyFrom(oDataTable);

                    string UID = "ItemsMTR";
                    SAPbouiCOM.LinkedButton oLink;

                    oColumn = oColumns.Add("LineNum", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = "#";
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "LineNum");

                    oColumn = oColumns.Add("DistNumber", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("DistNumber");
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "DistNumber");

                    oColumn = oColumns.Add("WhsCode", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("Warehouse");
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "WhsCode");
                    oLink = oColumn.ExtendedObject;
                    oLink.LinkedObjectType = "64";

                    oColumn = oColumns.Add("PrjCode", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("Project");
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "PrjCode");
                    oLink = oColumn.ExtendedObject;
                    oLink.LinkedObjectType = "63";

                    oColumn = oColumns.Add("ItemCode", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("AssetCode");
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "ItemCode");
                    oLink = oColumn.ExtendedObject;
                    oLink.LinkedObjectType = "4";

                    oColumn = oColumns.Add("ItemName", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("AssetName");
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "ItemName");

                    oColumn = oColumns.Add("ItmsGrpCod", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("AssetGroupCode");
                    oColumn.Editable = false;
                    oColumn.Visible = false;
                    oColumn.DataBind.Bind(UID, "ItmsGrpCod");

                    oColumn = oColumns.Add("ItmsGrpNam", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("AssetGroupName");
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "ItmsGrpNam");

                    oColumn = oColumns.Add("UsefulLife", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("UsefulLife");
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "UsefulLife");

                    oColumn = oColumns.Add("RmnngLife", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("RemainingLife");
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "RemainingLife");

                    oColumn = oColumns.Add("Quantity", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("Quantity");
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "Quantity");

                    oColumn = oColumns.Add("PrchsPrice", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("PurchasePrice");
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "PurchasePrice");

                    oColumn = oColumns.Add("PrchsCost", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("PurchaseCost");
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "PurchaseCost");

                    oColumn = oColumns.Add("NetBook", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("NetBookValue");
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "NetBookValue");

                    oColumn = oColumns.Add("DprAmt", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("DepreciationAmount");
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "DepreciationAmt");

                    oColumn = oColumns.Add("AlrdDprAmt", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("AlreadyDepreciatedAmountInCurrentMonth");
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "AlreadyDepreciatedAmt");

                    oColumn = oColumns.Add("DprEntry", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("DepreciationAccrualDocumentInCurrentMonth");
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "DepreciationDocEntry");
                    oLink = oColumn.ExtendedObject;
                    oLink.LinkedObjectType = "UDO_F_BDOSDEPACR_D";

                    oColumn = oColumns.Add("AccmDprAmt", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("AccumulatedDepreciationAmount");
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "AccumulatedDepreciationAmt");
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
            SAPbobsCOM.GeneralDataCollection oChildren = oGeneralData.Child("BDOSDEPAC1");
            int docEntry = 0;

            try
            {
                CommonFunctions.StartTransaction();

                oGeneralData.SetProperty("U_AccrMnth", AccrMnth);
                oGeneralData.SetProperty("U_DocDate", PostingDate);

                SAPbouiCOM.DataTable DepreciationLines = oForm.DataSources.DataTables.Item("ItemMTRTmp");

                for (int i = 0; i < DepreciationLines.Rows.Count; i++)
                {
                    SAPbobsCOM.GeneralData oChild = oChildren.Add();
                    oChild.SetProperty("U_ItemCode", DepreciationLines.GetValue("ItemCode", i));
                    oChild.SetProperty("U_DistNumber", DepreciationLines.GetValue("DistNumber", i));
                    oChild.SetProperty("U_BDOSUsLife", DepreciationLines.GetValue("UsefulLife", i));
                    oChild.SetProperty("U_Project", DepreciationLines.GetValue("PrjCode", i));
                    oChild.SetProperty("U_Quantity", DepreciationLines.GetValue("Quantity", i));
                    if (DepreciationLines.GetValue("RemainingLife", i) == 1)
                        oChild.SetProperty("U_DeprAmt", DepreciationLines.GetValue("NetBookValue", i));
                    else
                        oChild.SetProperty("U_DeprAmt", DepreciationLines.GetValue("DepreciationAmt", i));
                    oChild.SetProperty("U_AccmDprAmt", DepreciationLines.GetValue("AccumulatedDepreciationAmt", i));
                }

                var response = oGeneralService.Add(oGeneralData);
                docEntry = response.GetProperty("DocEntry");

                if (docEntry > 0)
                {
                    DataTable JrnLinesDT = BDOSDepreciationAccrualDocument.createAdditionalEntries(null, oGeneralData, 0, PostingDate, "", null, null);

                    BDOSDepreciationAccrualDocument.JrnEntry(docEntry.ToString(), docEntry.ToString(), PostingDate, JrnLinesDT, "", out errorText);

                    if (errorText != null)
                        throw new Exception(errorText);
                    else
                    {
                        CommonFunctions.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                        Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("DocumentCreatedSuccesfully") + " #" + docEntry.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                    }
                }
            }
            catch (Exception ex)
            {
                CommonFunctions.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("DocumentNotCreated") + ". " + BDOSResources.getTranslate("ReasonIs") + ": " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short);
            }
            finally
            {
                Marshal.ReleaseComObject(oChildren);
                Marshal.ReleaseComObject(oGeneralData);
                Marshal.ReleaseComObject(oGeneralService);
                Marshal.ReleaseComObject(oCompanyService);
            }
            return docEntry;
        }

        private static void CreateDocuments(SAPbouiCOM.Form oForm)
        {
            //bool isInvoice = oForm.Items.Item("InvDepr").Specific.Selected;

            SAPbouiCOM.DataTable depreciationLinesTmp = oForm.DataSources.DataTables.Item("ItemMTRTmp");
            depreciationLinesTmp.Rows.Clear();
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("ItemsMTR").Specific;
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
                string project = depreciationLines.GetValue("PrjCode", i);

                if (depreciationLines.GetValue("DepreciationDocEntry", i) == 0)
                {
                    newRow = depreciationLinesTmp.Rows.Count;
                    depreciationLinesTmp.Rows.Add();
                    depreciationLinesTmp.SetValue("ItemCode", newRow, itemCode);
                    depreciationLinesTmp.SetValue("DistNumber", newRow, distNumber);
                    depreciationLinesTmp.SetValue("UsefulLife", newRow, depreciationLines.GetValue("UsefulLife", i));
                    depreciationLinesTmp.SetValue("RemainingLife", newRow, depreciationLines.GetValue("RemainingLife", i));
                    depreciationLinesTmp.SetValue("PrjCode", newRow, project);
                    depreciationLinesTmp.SetValue("Quantity", newRow, depreciationLines.GetValue("Quantity", i));
                    depreciationLinesTmp.SetValue("NetBookValue", newRow, depreciationLines.GetValue("NetBookValue", i));
                    depreciationLinesTmp.SetValue("DepreciationAmt", newRow, depreciationLines.GetValue("DepreciationAmt", i));
                    depreciationLinesTmp.SetValue("AccumulatedDepreciationAmt", newRow, depreciationLines.GetValue("AccumulatedDepreciationAmt", i));

                    ////if (isInvoice)
                    ////{
                    //depreciationLinesTmp.SetValue("DocEntry", newRow, depreciationLines.GetValue("DocEntry", i));
                    //depreciationLinesTmp.SetValue("DocType", newRow, depreciationLines.GetValue("DocType", i));

                    //SAPbobsCOM.BoObjectTypes DocType;

                    //if (depreciationLines.GetValue("DocType", i) == "13") //Invoice object
                    //    DocType = SAPbobsCOM.BoObjectTypes.oInvoices;
                    //else
                    //    DocType = SAPbobsCOM.BoObjectTypes.oInventoryGenExit;

                    //SAPbobsCOM.Documents oInvoice = Program.oCompany.GetBusinessObject(DocType);
                    //DateTime DocDate = new DateTime();
                    //if (oInvoice.GetByKey(depreciationLines.GetValue("DocEntry", i)))
                    //    DocDate = oInvoice.DocDate;

                    //CreateDocument(oForm, accrMnth, DocDate);

                    //depreciationLinesTmp.Rows.Clear();
                    ////}
                }
                else
                {
                    Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("DocumentAlreadyCreated") + "! " + itemCode + " - " + distNumber + " - " + project, SAPbouiCOM.BoMessageTime.bmt_Short);
                    continue;
                }

                StringBuilder query = new StringBuilder();
                //query.Append("SELECT * \n");
                //query.Append("FROM \n");
                query.Append("(SELECT TOP 1 \n");
                query.Append("\"@BDOSDEPACR\".\"DocEntry\", \n");
                query.Append("\"@BDOSDEPACR\".\"U_DocDate\" \n");
                query.Append("FROM \"@BDOSDEPACR\" \n");
                query.Append("INNER JOIN \"@BDOSDEPAC1\" ON \"@BDOSDEPACR\".\"DocEntry\" = \"@BDOSDEPAC1\".\"DocEntry\" \n");
                query.Append("WHERE \"U_DocDate\" <= '" + dateStr + "' \n");
                query.Append("AND \"Canceled\" = 'N' \n");
                query.Append("AND \"@BDOSDEPAC1\".\"U_ItemCode\" = '" + itemCode + "' \n");
                query.Append("AND \"@BDOSDEPAC1\".\"U_DistNumber\" = '" + distNumber + "' \n");
                query.Append("AND \"@BDOSDEPAC1\".\"U_Project\" = '" + project + "' \n");
                query.Append("ORDER BY \n");
                query.Append("\"U_DocDate\" DESC, \n");
                query.Append("\"DocEntry\" DESC) \n");
                //query.Append("WHERE ABS((MONTH('" + dateStr + "') - MONTH(\"U_DocDate\")))>1");

                SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRecordSet.DoQuery(query.ToString());
                if (!oRecordSet.EoF)
                {
                    if (Math.Abs(accrMnth.Month - oRecordSet.Fields.Item("U_DocDate").Value.Month) > 1)
                    {
                        depreciationLinesTmp.Rows.Remove(newRow);
                        Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("CreateDocumentForThePreviousMonths") + "! " + itemCode + " - " + distNumber + " - " + project, SAPbouiCOM.BoMessageTime.bmt_Short);
                        Marshal.ReleaseComObject(oRecordSet);
                        continue;
                    }
                }
            }

            if (depreciationLinesTmp.Rows.Count > 0) //!isInvoice && 
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
            //bool isInvoice = oForm.Items.Item("InvDepr").Specific.Selected;

            //string query = BatchDepreciaionQuery(deprMonth, "", "", "", isInvoice);

            StringBuilder query = new StringBuilder();
            query.Append("SELECT T0.*, \n");
            query.Append("       CASE \n");
            query.Append("         WHEN T0.\"AlreadyDepreciatedAmt\" IS NULL THEN \n");
            query.Append("         T0.\"PurchaseCost\" / T0.\"UsefulLife\" \n");
            query.Append("         ELSE 0 \n");
            query.Append("       END                                                  AS \"DepreciationAmt\", \n");
            query.Append("       T0.\"PurchaseCost\" - \"AccumulatedDepreciationAmt\" AS \"NetBookValue\", \n");
            query.Append("       T0.\"UsefulLife\" - T0.\"AllDeprDocQty\"             AS \"RemainingLife\" \n");
            query.Append("FROM   (SELECT \"OBTN\".\"DistNumber\", \n");
            query.Append("               \"OBTQ\".\"WhsCode\", \n");
            query.Append("               \"OWHS\".\"WhsName\", \n");
            query.Append("               \"OWHS\".\"U_BDOSPrjCod\" AS \"PrjCode\", \n");
            query.Append("               \"OBTQ\".\"ItemCode\", \n");
            query.Append("               \"OITM\".\"ItemName\", \n");
            query.Append("               \"OITM\".\"ItmsGrpCod\", \n");
            query.Append("               \"OITB\".\"ItmsGrpNam\", \n");
            query.Append("               \"OBTQ\".\"Quantity\", \n");
            query.Append("               \"OITM\".\"U_BDOSUsLife\" AS \"UsefulLife\", \n");
            query.Append("               T1.\"U_DeprAmt\"                                                 AS \"AccumulatedDepreciationAmt\", \n");
            query.Append("               T2.\"U_DeprAmt\"                                                 AS \"AlreadyDepreciatedAmt\", \n");
            query.Append("               T2.\"DepreciationDocEntry\", \n");
            query.Append("               \"OBTN\".\"CostTotal\" / \"OBTN\".\"Quantity\"                             AS \"PurchasePrice\", \n");
            query.Append("               ( \"OBTN\".\"CostTotal\" / \"OBTN\".\"Quantity\" ) * \"OBTQ\".\"Quantity\" AS \"PurchaseCost\", \n");
            query.Append("               (SELECT Count(DISTINCT \"@BDOSDEPAC1\".\"DocEntry\") \n");
            query.Append("               FROM   \"@BDOSDEPAC1\" \n");
            query.Append("                       INNER JOIN \"@BDOSDEPACR\" \n");
            query.Append("                               ON \"@BDOSDEPAC1\".\"DocEntry\" = \"@BDOSDEPACR\".\"DocEntry\" \n");
            query.Append("               WHERE  \"@BDOSDEPACR\".\"Canceled\" = 'N' \n");
            query.Append("                       AND \"@BDOSDEPAC1\".\"U_ItemCode\" = \"OBTQ\".\"ItemCode\")    AS \"AllDeprDocQty\" \n");
            query.Append("        FROM   \"OBTQ\" \n");
            query.Append("               INNER JOIN \"OBTN\" \n");
            query.Append("                       ON \"OBTQ\".\"ItemCode\" = \"OBTN\".\"ItemCode\" \n");
            query.Append("                          AND \"OBTQ\".\"SysNumber\" = \"OBTN\".\"SysNumber\" \n");
            query.Append("                          AND \"OBTQ\".\"MdAbsEntry\" = \"OBTN\".\"AbsEntry\" \n");
            query.Append("               INNER JOIN \"OWHS\" \n");
            query.Append("                       ON \"OBTQ\".\"WhsCode\" = \"OWHS\".\"WhsCode\" \n");
            query.Append("               INNER JOIN \"OITM\" \n");
            query.Append("                       ON \"OBTQ\".\"ItemCode\" = \"OITM\".\"ItemCode\" \n");
            query.Append("               INNER JOIN \"OITB\" \n");
            query.Append("                       ON \"OITM\".\"ItmsGrpCod\" = \"OITB\".\"ItmsGrpCod\" \n");
            query.Append("                          AND \"OITB\".\"U_BDOSFxAs\" = 'Y' \n");
            query.Append("               LEFT JOIN (SELECT \"@BDOSDEPAC1\".\"U_Project\", \n");
            query.Append("                                 \"@BDOSDEPAC1\".\"U_DistNumber\", \n");
            query.Append("                                 \"@BDOSDEPAC1\".\"U_ItemCode\", \n");
            query.Append("                                 \"@BDOSDEPAC1\".\"U_DeprAmt\" \n");
            query.Append("                          FROM   \"@BDOSDEPAC1\" \n");
            query.Append("                                 INNER JOIN \"@BDOSDEPACR\" \n");
            query.Append("                                         ON \"@BDOSDEPAC1\".\"DocEntry\" = \"@BDOSDEPACR\".\"DocEntry\" \n");
            query.Append("                          WHERE  \"@BDOSDEPACR\".\"Canceled\" = 'N' \n");
            query.Append("                                 AND \"@BDOSDEPACR\".\"U_AccrMnth\" <= '" + dateStr + "') AS T1 \n");
            query.Append("                      ON T1.\"U_ItemCode\" = \"OBTQ\".\"ItemCode\" \n");
            query.Append("                         AND T1.\"U_DistNumber\" = \"OBTN\".\"DistNumber\" \n");
            query.Append("                         AND T1.\"U_Project\" = \"OWHS\".\"U_BDOSPrjCod\" \n");
            query.Append("               LEFT JOIN (SELECT \"@BDOSDEPACR\".\"DocEntry\" AS \"DepreciationDocEntry\", \n");
            query.Append("                                 \"@BDOSDEPAC1\".\"U_Project\", \n");
            query.Append("                                 \"@BDOSDEPAC1\".\"U_DistNumber\", \n");
            query.Append("                                 \"@BDOSDEPAC1\".\"U_ItemCode\", \n");
            query.Append("                                 \"@BDOSDEPAC1\".\"U_DeprAmt\" \n");
            query.Append("                          FROM   \"@BDOSDEPAC1\" \n");
            query.Append("                                 INNER JOIN \"@BDOSDEPACR\" \n");
            query.Append("                                         ON \"@BDOSDEPAC1\".\"DocEntry\" = \"@BDOSDEPACR\".\"DocEntry\" \n");
            query.Append("                          WHERE  \"@BDOSDEPACR\".\"Canceled\" = 'N' \n");
            query.Append("                                 AND \"@BDOSDEPACR\".\"U_AccrMnth\" = '" + dateStr + "') AS T2 \n");
            query.Append("                      ON T2.\"U_ItemCode\" = \"OBTQ\".\"ItemCode\" \n");
            query.Append("                         AND T2.\"U_DistNumber\" = \"OBTN\".\"DistNumber\" \n");
            query.Append("                         AND T2.\"U_Project\" = \"OWHS\".\"U_BDOSPrjCod\" \n");
            query.Append("        WHERE \"OITM\".\"U_BDOSUsLife\" > 0 AND \"OBTN\".\"Quantity\" > 0) AS T0");

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRecordSet.DoQuery(query.ToString());
            int rowIndex = 0;
            //int i = 0;
            while (!oRecordSet.EoF)
            {
                oDataTable.Rows.Add();
                oDataTable.SetValue("LineNum", rowIndex, rowIndex + 1);
                oDataTable.SetValue("DistNumber", rowIndex, oRecordSet.Fields.Item("DistNumber").Value);
                oDataTable.SetValue("WhsCode", rowIndex, oRecordSet.Fields.Item("WhsCode").Value);
                oDataTable.SetValue("PrjCode", rowIndex, oRecordSet.Fields.Item("PrjCode").Value);
                oDataTable.SetValue("ItemCode", rowIndex, oRecordSet.Fields.Item("ItemCode").Value);
                oDataTable.SetValue("ItemName", rowIndex, oRecordSet.Fields.Item("ItemName").Value);
                oDataTable.SetValue("ItmsGrpCod", rowIndex, oRecordSet.Fields.Item("ItmsGrpCod").Value);
                oDataTable.SetValue("ItmsGrpNam", rowIndex, oRecordSet.Fields.Item("ItmsGrpNam").Value);
                oDataTable.SetValue("UsefulLife", rowIndex, oRecordSet.Fields.Item("UsefulLife").Value);
                oDataTable.SetValue("RemainingLife", rowIndex, oRecordSet.Fields.Item("RemainingLife").Value);
                oDataTable.SetValue("Quantity", rowIndex, oRecordSet.Fields.Item("Quantity").Value);
                oDataTable.SetValue("PurchasePrice", rowIndex, oRecordSet.Fields.Item("PurchasePrice").Value);
                oDataTable.SetValue("PurchaseCost", rowIndex, oRecordSet.Fields.Item("PurchaseCost").Value);
                oDataTable.SetValue("NetBookValue", rowIndex, oRecordSet.Fields.Item("NetBookValue").Value);
                oDataTable.SetValue("DepreciationAmt", rowIndex, oRecordSet.Fields.Item("DepreciationAmt").Value);
                oDataTable.SetValue("AlreadyDepreciatedAmt", rowIndex, oRecordSet.Fields.Item("AlreadyDepreciatedAmt").Value);
                if ((int)oRecordSet.Fields.Item("DepreciationDocEntry").Value != 0)
                    oDataTable.SetValue("DepreciationDocEntry", rowIndex, oRecordSet.Fields.Item("DepreciationDocEntry").Value);
                oDataTable.SetValue("AccumulatedDepreciationAmt", rowIndex, oRecordSet.Fields.Item("AccumulatedDepreciationAmt").Value);

                //DateTime InDateStart = oRecordSet.Fields.Item("InDate").Value;
                //DateTime InDateEnd = InDateStart.AddMonths(oRecordSet.Fields.Item("UsefulLife").Value);
                //InDateEnd = new DateTime(InDateEnd.Year, InDateEnd.Month, 1);
                //InDateEnd = InDateEnd.AddMonths(1).AddDays(-1);
                //DateTime AccrMnth = InDateStart;
                //AccrMnth = new DateTime(AccrMnth.Year, AccrMnth.Month, 1);
                //AccrMnth = AccrMnth.AddMonths(1).AddDays(-1);

                //decimal Quantity = Convert.ToDecimal(oRecordSet.Fields.Item("Quantity").Value);
                //Quantity = Quantity * (isInvoice ? -1 : 1);
                //i++;

                //if (deprMonth > InDateEnd || deprMonth <= AccrMnth || Quantity == 0)
                //{
                //    oRecordSet.MoveNext();
                //    continue;
                //}

                //decimal CurrDeprAmt = Convert.ToDecimal(oRecordSet.Fields.Item("CurrDeprAmt").Value, CultureInfo.InvariantCulture);
                //int monthsApart = 12 * (InDateEnd.Year - deprMonth.Year) + (InDateEnd.Month - deprMonth.Month) + 1;
                //monthsApart = Math.Abs(monthsApart);

                //decimal AlrDeprAmt = 0;
                ////AlrDeprAmt = Convert.ToDecimal(oRecordSet.Fields.Item("DeprAmt").Value)  * Quantity;
                //AlrDeprAmt = Convert.ToDecimal(oRecordSet.Fields.Item("AlrDeprAmt").Value) * Quantity;
                //AlrDeprAmt -= CurrDeprAmt;

                //decimal NtBookVal = Convert.ToDecimal(oRecordSet.Fields.Item("APCost").Value * Convert.ToDouble(Quantity)) - AlrDeprAmt;
                //decimal DeprAmt = NtBookVal / monthsApart;

                //oDataTable.Rows.Add();
                //oDataTable.SetValue("LineNum", rowIndex, rowIndex + 1);
                //oDataTable.SetValue("PrjCode", rowIndex, oRecordSet.Fields.Item("PrjCode").Value);
                //oDataTable.SetValue("ItmsGrpCod", rowIndex, oRecordSet.Fields.Item("ItmsGrpCod").Value);
                //oDataTable.SetValue("ItemCode", rowIndex, oRecordSet.Fields.Item("ItemCode").Value);
                //oDataTable.SetValue("ItemName", rowIndex, oRecordSet.Fields.Item("ItemName").Value);
                //oDataTable.SetValue("DistNumber", rowIndex, oRecordSet.Fields.Item("DistNumber").Value);
                //if (oRecordSet.Fields.Item("DepcDoc").Value != 0)
                //    oDataTable.SetValue("DepcDoc", rowIndex, oRecordSet.Fields.Item("DepcDoc").Value);
                //oDataTable.SetValue("UsefulLife", rowIndex, monthsApart);
                //oDataTable.SetValue("Quantity", rowIndex, Convert.ToDouble(Quantity));
                //oDataTable.SetValue("APCost", rowIndex, oRecordSet.Fields.Item("APCost").Value * Convert.ToDouble(Quantity));
                //oDataTable.SetValue("AlrDeprAmt", rowIndex, Convert.ToDouble(AlrDeprAmt));
                //oDataTable.SetValue("NtBookVal", rowIndex, Convert.ToDouble(NtBookVal));
                //if (CurrDeprAmt > 0)
                //    oDataTable.SetValue("DeprAmt", rowIndex, 0);
                //else
                //    oDataTable.SetValue("DeprAmt", rowIndex, Convert.ToDouble(DeprAmt));
                //oDataTable.SetValue("CurMnthAmt", rowIndex, Convert.ToDouble(CurrDeprAmt));
                //if (isInvoice)
                //{
                //    oDataTable.SetValue("DocEntry", rowIndex, oRecordSet.Fields.Item("DocEntry").Value);
                //    oDataTable.SetValue("DocType", rowIndex, oRecordSet.Fields.Item("DocType").Value);
                //}

                oRecordSet.MoveNext();
                rowIndex++;
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

                //bool isRetirement = oForm.Items.Item("InvDepr").Specific.Selected;

                //SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("ItemsMTR").Specific;

                //if (isRetirement)
                //{
                //    oMatrix.Columns.Item("DocType").Visible = true;
                //    oMatrix.Columns.Item("DocEntry").Visible = true;
                //}
                //else
                //{
                //    oMatrix.Columns.Item("DocType").Visible = false;
                //    oMatrix.Columns.Item("DocEntry").Visible = false;
                //}

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
