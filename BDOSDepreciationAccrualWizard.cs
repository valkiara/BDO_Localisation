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

                    top += height + 10;

                    var oItem = oForm.Items.Add("Rtrmnt", SAPbouiCOM.BoFormItemTypes.it_OPTION_BUTTON);
                    oItem.Left = left_s;
                    oItem.Top = top;
                    oItem.Height = height;
                    oItem.Width = width_s;
                    SAPbouiCOM.OptionBtn optBtn = oItem.Specific;
                    optBtn.Caption = BDOSResources.getTranslate("Retirement");
                    var oUserDataSource = oForm.DataSources.UserDataSources.Add("Rtrmnt", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                    optBtn.DataBind.SetBound(true, "", "Rtrmnt");
                    //optBtn.PressedAfter += (o, a) =>
                    //{
                    //    if (!a.InnerEvent)
                    //        OptBtnPressedAfter(oForm);
                    //};

                    oItem = oForm.Items.Add("Dprctn", SAPbouiCOM.BoFormItemTypes.it_OPTION_BUTTON);
                    oItem.Left = left_e;
                    oItem.Top = top;
                    oItem.Height = height;
                    oItem.Width = width_s;
                    optBtn = oItem.Specific;
                    optBtn.Caption = BDOSResources.getTranslate("Depreciation");
                    optBtn.GroupWith("Rtrmnt");
                    //optBtn.PressedAfter += (o, a) =>
                    //{
                    //    if (!a.InnerEvent)
                    //        OptBtnPressedAfter(oForm);
                    //};

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
                    oDataTable.Columns.Add("InDate", SAPbouiCOM.BoFieldsType.ft_Date);
                    oDataTable.Columns.Add("BaseType", SAPbouiCOM.BoFieldsType.ft_Integer);
                    oDataTable.Columns.Add("LastDeprDocDate", SAPbouiCOM.BoFieldsType.ft_Date);
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

                    oColumn = oColumns.Add("InDate", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("AdmissionDate");
                    oColumn.Editable = false;
                    //oColumn.Visible = false;
                    oColumn.DataBind.Bind(UID, "InDate");

                    oColumn = oColumns.Add("BaseType", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("Type");
                    oColumn.Editable = false;
                    oColumn.Visible = false;
                    oColumn.DataBind.Bind(UID, "BaseType");

                    oColumn = oColumns.Add("LstDprDcDt", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("DateOfLastExistingDepreciationAccrualDocument");
                    oColumn.Editable = false;
                    //oColumn.Visible = false;
                    oColumn.DataBind.Bind(UID, "LastDeprDocDate");

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

        private static void OptBtnPressedAfter(SAPbouiCOM.Form oForm)
        {
            oForm.Freeze(true);
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("ItemsMTR").Specific;
            if (oMatrix.RowCount > 0)
                oMatrix.Clear();
            oForm.Freeze(false);
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

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE || pVal.EventType == SAPbouiCOM.BoEventTypes.et_GOT_FOCUS)
                    return;

                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_CLOSE && pVal.BeforeAction)
                {
                    int answer = Program.uiApp.MessageBox(BDOSResources.getTranslate("CloseDepreciationAccruingWizard") + "?", 1, BDOSResources.getTranslate("Yes"), BDOSResources.getTranslate("No"), "");

                    if (answer != 1)
                        BubbleEvent = false;
                }

                else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE && !pVal.BeforeAction)
                {
                    oForm.Freeze(true);
                    SAPbouiCOM.OptionBtn optBtn = oForm.Items.Item("Rtrmnt").Specific;
                    optBtn.Selected = true;
                    oForm.Freeze(false);
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
                            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("ItemsMTR").Specific;
                            if (oMatrix.RowCount > 0)
                            {
                                CreateDocuments(oForm);
                                fillMTRItems(oForm);
                            }
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
                        if (!pVal.InnerEvent && (pVal.ItemUID == "Dprctn" || pVal.ItemUID == "Rtrmnt"))
                            OptBtnPressedAfter(oForm);
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

        private static int CreateDocument(SAPbouiCOM.Form oForm, DateTime accrMnth, DateTime postingDate, bool isRetirement)
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

                oGeneralData.SetProperty("U_AccrMnth", accrMnth);
                oGeneralData.SetProperty("U_DocDate", postingDate);
                if (isRetirement)
                    oGeneralData.SetProperty("U_Retirement", "Y");

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
                    oChild.SetProperty("U_AccmDprAmt", DepreciationLines.GetValue("AccumulatedDepreciationAmt", i) + DepreciationLines.GetValue("DepreciationAmt", i));
                }

                var response = oGeneralService.Add(oGeneralData);
                docEntry = response.GetProperty("DocEntry");

                if (docEntry > 0)
                {
                    DataTable jrnLinesDT = BDOSDepreciationAccrualDocument.createAdditionalEntries(null, oGeneralData, 0, postingDate);

                    BDOSDepreciationAccrualDocument.JrnEntry(docEntry.ToString(), docEntry.ToString(), postingDate, jrnLinesDT, "", out errorText);

                    if (errorText != null)
                        throw new Exception(errorText);
                    else
                    {
                        CommonFunctions.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                        Program.uiApp.StatusBar.SetSystemMessage($"{BDOSResources.getTranslate("DocumentCreatedSuccesfully")}", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success, "", "", $"{BDOSResources.getTranslate("DocEntry")}: {docEntry.ToString()}");
                    }
                }
            }
            catch (Exception ex)
            {
                CommonFunctions.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                Program.uiApp.StatusBar.SetSystemMessage($"{BDOSResources.getTranslate("DocumentNotCreated")}", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error, "", "", $"{ex.Message}");
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
            bool isRetirement = oForm.Items.Item("Rtrmnt").Specific.Selected;

            //if (isRetirement)
            //    return;

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
                DateTime? lastDeprDocDate = depreciationLines.GetValue("LastDeprDocDate", i);
                int lineNum = depreciationLines.GetValue("LineNum", i);

                if (depreciationLines.GetValue("DepreciationDocEntry", i) == 0)
                {
                    if (!isRetirement)
                    {
                        DateTime inDate = depreciationLines.GetValue("InDate", i);
                        DateTime dateForCheck = lastDeprDocDate.HasValue ? lastDeprDocDate.Value : inDate;
                        DateTime lastDayOfDateForCheckNextMonth;

                        if (depreciationLines.GetValue("BaseType", i) == 67 && !lastDeprDocDate.HasValue)
                            lastDayOfDateForCheckNextMonth = new DateTime(dateForCheck.Year, dateForCheck.Month, 1).AddMonths(1).AddDays(-1);
                        else
                            lastDayOfDateForCheckNextMonth = new DateTime(dateForCheck.Year, dateForCheck.Month, 1).AddMonths(2).AddDays(-1);

                        if (lastDayOfDateForCheckNextMonth != accrMnth)
                        {
                            string text = lastDeprDocDate.HasValue ? BDOSResources.getTranslate("PayAttentionToTheDateOfLastExistingDepreciationAccrualDocument") : BDOSResources.getTranslate("PayAttentionToTheAdmissionDate");
                            Program.uiApp.StatusBar.SetSystemMessage($"{BDOSResources.getTranslate("CreateDocumentForThePreviousMonths")}", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error, "", "", $"{text}! {BDOSResources.getTranslate("TableRow")}: {lineNum}");
                            continue;
                        }
                    }

                    newRow = depreciationLinesTmp.Rows.Count;
                    depreciationLinesTmp.Rows.Add();
                    depreciationLinesTmp.SetValue("LineNum", newRow, lineNum);
                    //depreciationLinesTmp.SetValue("InDate", newRow, inDate);
                    depreciationLinesTmp.SetValue("ItemCode", newRow, itemCode);
                    depreciationLinesTmp.SetValue("DistNumber", newRow, distNumber);
                    depreciationLinesTmp.SetValue("UsefulLife", newRow, depreciationLines.GetValue("UsefulLife", i));
                    depreciationLinesTmp.SetValue("RemainingLife", newRow, depreciationLines.GetValue("RemainingLife", i));
                    depreciationLinesTmp.SetValue("PrjCode", newRow, project);
                    depreciationLinesTmp.SetValue("Quantity", newRow, depreciationLines.GetValue("Quantity", i));
                    depreciationLinesTmp.SetValue("NetBookValue", newRow, depreciationLines.GetValue("NetBookValue", i));
                    depreciationLinesTmp.SetValue("DepreciationAmt", newRow, depreciationLines.GetValue("DepreciationAmt", i));
                    depreciationLinesTmp.SetValue("AccumulatedDepreciationAmt", newRow, depreciationLines.GetValue("AccumulatedDepreciationAmt", i));
                }
                else
                {
                    Program.uiApp.StatusBar.SetSystemMessage($"{BDOSResources.getTranslate("DocumentAlreadyCreated")}", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error, "", "", $"{BDOSResources.getTranslate("TableRow")}: {lineNum}");
                    continue;
                }
            }

            if (depreciationLinesTmp.Rows.Count > 0)
                CreateDocument(oForm, accrMnth, accrMnth, isRetirement);
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

            StringBuilder query = new StringBuilder();

            bool isRetirement = oForm.Items.Item("Rtrmnt").Specific.Selected;

            if (isRetirement)
                query = getQueryForRetirement(dateStr);
            else
                query = getQueryForDepreciation(dateStr);

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
                if (!isRetirement)
                    oDataTable.SetValue("InDate", rowIndex, oRecordSet.Fields.Item("InDate").Value);
                oDataTable.SetValue("BaseType", rowIndex, oRecordSet.Fields.Item("BaseType").Value);
                if (oRecordSet.Fields.Item("LastDeprDocDate").Value.ToString("yyyyMMdd") != "18991230")
                    oDataTable.SetValue("LastDeprDocDate", rowIndex, oRecordSet.Fields.Item("LastDeprDocDate").Value);
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
                if (isRetirement)
                    oDataTable.SetValue("AccumulatedDepreciationAmt", rowIndex, oRecordSet.Fields.Item("AccumulatedDepreciationAmt").Value);
                else
                    oDataTable.SetValue("AccumulatedDepreciationAmt", rowIndex, oRecordSet.Fields.Item("Coefficient").Value * oRecordSet.Fields.Item("AccumulatedDepreciationAmt").Value);

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

        static StringBuilder getQueryForRetirement(string dateStr)
        {
            StringBuilder query = new StringBuilder();

            query.Append("SELECT T0.*, \n");
            query.Append("       CASE \n");
            query.Append("         WHEN T0.\"AlreadyDepreciatedAmt\" = 0 THEN \n");
            query.Append("         T0.\"PurchaseCost\" / T0.\"UsefulLife\" \n");
            query.Append("         ELSE 0 \n");
            query.Append("       END                                                   AS \"DepreciationAmt\", \n");
            query.Append("       T0.\"PurchaseCost\" - (T0.\"AccumulatedDepreciationAmt\") AS \"NetBookValue\", \n");
            query.Append("       T0.\"UsefulLife\" - T0.\"AllDeprDocQty\"                  AS \"RemainingLife\" \n");
            query.Append("FROM   (SELECT \"OBTN\".\"DistNumber\", \n");
            query.Append("               \"OIBT\".\"WhsCode\", \n");
            query.Append("               \"OWHS\".\"WhsName\", \n");
            query.Append("               \"OWHS\".\"U_BDOSPrjCod\"                                        AS \"PrjCode\", \n");
            query.Append("               \"OIBT\".\"BaseType\", \n");
            query.Append("               \"T3\".\"LastDeprDocDate\", \n");
            query.Append("               \"OIBT\".\"ItemCode\", \n");
            query.Append("               \"OITM\".\"ItemName\", \n");
            query.Append("               \"OITM\".\"ItmsGrpCod\", \n");
            query.Append("               \"OITB\".\"ItmsGrpNam\", \n");
            query.Append("               \"OIBT\".\"Quantity\", \n");
            query.Append("               \"OITM\".\"U_BDOSUsLife\"                                        AS \"UsefulLife\", \n");
            query.Append("               CASE WHEN T1.\"U_DeprAmt\" IS NULL THEN 0 ELSE T1.\"U_DeprAmt\" END        AS \"AccumulatedDepreciationAmt\", \n");
            query.Append("               CASE WHEN T2.\"U_DeprAmt\" IS NULL THEN 0 ELSE T2.\"U_DeprAmt\" END        AS \"AlreadyDepreciatedAmt\", \n");
            query.Append("               T2.\"DepreciationDocEntry\", \n");
            query.Append("               \"OBTN\".\"CostTotal\" / \"OBTN\".\"Quantity\"                             AS \"PurchasePrice\", \n");
            query.Append("               ( \"OBTN\".\"CostTotal\" / \"OBTN\".\"Quantity\" ) * \"OIBT\".\"Quantity\" AS \"PurchaseCost\", \n");
            query.Append("               CASE WHEN T4.\"DocEntry\" IS NULL THEN 0 ELSE T4.\"DocEntry\" END AS \"AllDeprDocQty\" \n");
            query.Append("        FROM   (SELECT B0.\"SysNumber\", \n");
            query.Append("               B0.\"ItemCode\", \n");
            query.Append("               B0.\"BatchNum\", \n");
            query.Append("               B0.\"WhsCode\", \n");
            query.Append("               B1.\"BaseType\" AS \"BaseType\", \n");
            query.Append("               SUM(CASE WHEN B1.\"Direction\" = 1 THEN B1.\"Quantity\" ELSE( -1 ) * B1.\"Quantity\" END) AS \"Quantity\" \n");
            query.Append("        FROM   \"OIBT\" B0 \n");
            query.Append("               INNER JOIN \"IBT1\" B1 \n");
            query.Append("                       ON B0.\"ItemCode\" = B1.\"ItemCode\" \n");
            query.Append("                          AND B0.\"BatchNum\" = B1.\"BatchNum\" \n");
            query.Append("                          AND B0.\"WhsCode\" = B1.\"WhsCode\" \n");
            query.Append("        WHERE  B1.\"BaseType\" IN(13, 60) \n");
            query.Append($"               AND B1.\"DocDate\" <= '{dateStr}' \n");
            query.Append("        GROUP  BY B0.\"SysNumber\", \n");
            query.Append("                  B0.\"ItemCode\", \n");
            query.Append("                  B0.\"BatchNum\", \n");
            query.Append("                  B0.\"WhsCode\", \n");
            query.Append("                  B1.\"BaseType\" \n");
            query.Append("        HAVING SUM(CASE WHEN B1.\"Direction\" = 1 THEN B1.\"Quantity\" ELSE( -1 ) * B1.\"Quantity\" END) > 0) \"OIBT\" \n");
            query.Append("               INNER JOIN \"OBTN\" \n");
            query.Append("                       ON \"OIBT\".\"ItemCode\" = \"OBTN\".\"ItemCode\" \n");
            query.Append("                          AND \"OIBT\".\"SysNumber\" = \"OBTN\".\"SysNumber\" \n");
            query.Append("                          AND \"OIBT\".\"BatchNum\" = \"OBTN\".\"DistNumber\" \n");
            query.Append("               INNER JOIN \"OWHS\" \n");
            query.Append("                       ON \"OIBT\".\"WhsCode\" = \"OWHS\".\"WhsCode\" \n");
            query.Append("               INNER JOIN \"OITM\" \n");
            query.Append("                       ON \"OIBT\".\"ItemCode\" = \"OITM\".\"ItemCode\" \n");
            query.Append("               INNER JOIN \"OITB\" \n");
            query.Append("                       ON \"OITM\".\"ItmsGrpCod\" = \"OITB\".\"ItmsGrpCod\" \n");
            query.Append("                          AND \"OITB\".\"U_BDOSFxAs\" = 'Y' \n");
            query.Append("               LEFT JOIN (SELECT \n");
            query.Append("                                 \"@BDOSDEPAC1\".\"U_DistNumber\", \n");
            query.Append("                                 \"@BDOSDEPAC1\".\"U_ItemCode\", \n");
            query.Append("                                 SUM(\"@BDOSDEPAC1\".\"U_DeprAmt\") AS \"U_DeprAmt\" \n");
            query.Append("                          FROM   \"@BDOSDEPAC1\" \n");
            query.Append("                                 INNER JOIN \"@BDOSDEPACR\" \n");
            query.Append("                                         ON \"@BDOSDEPAC1\".\"DocEntry\" = \"@BDOSDEPACR\".\"DocEntry\" \n");
            query.Append("                          WHERE  \"@BDOSDEPACR\".\"Canceled\" = 'N' /*AND \"@BDOSDEPACR\".\"U_Retirement\" = 'Y'*/ \n");
            query.Append($"                                 AND \"@BDOSDEPACR\".\"U_AccrMnth\" <= '{dateStr}' \n");
            query.Append("                          GROUP BY \"@BDOSDEPAC1\".\"U_DistNumber\", \"@BDOSDEPAC1\".\"U_ItemCode\") AS T1 \n");
            query.Append("                      ON T1.\"U_ItemCode\" = \"OIBT\".\"ItemCode\" \n");
            query.Append("                         AND T1.\"U_DistNumber\" = \"OBTN\".\"DistNumber\" \n");
            query.Append("               LEFT JOIN (SELECT \"@BDOSDEPACR\".\"DocEntry\" AS \"DepreciationDocEntry\", \n");
            query.Append("                                 \"@BDOSDEPAC1\".\"U_Project\", \n");
            query.Append("                                 \"@BDOSDEPAC1\".\"U_DistNumber\", \n");
            query.Append("                                 \"@BDOSDEPAC1\".\"U_ItemCode\", \n");
            query.Append("                                 \"@BDOSDEPAC1\".\"U_DeprAmt\" \n");
            query.Append("                          FROM   \"@BDOSDEPAC1\" \n");
            query.Append("                                 INNER JOIN \"@BDOSDEPACR\" \n");
            query.Append("                                         ON \"@BDOSDEPAC1\".\"DocEntry\" = \"@BDOSDEPACR\".\"DocEntry\" \n");
            query.Append("                          WHERE  \"@BDOSDEPACR\".\"Canceled\" = 'N' AND \"@BDOSDEPACR\".\"U_Retirement\" = 'Y' \n");
            query.Append($"                                 AND \"@BDOSDEPACR\".\"U_AccrMnth\" = '{dateStr}') AS T2 \n");
            query.Append("                      ON T2.\"U_ItemCode\" = \"OIBT\".\"ItemCode\" \n");
            query.Append("                         AND T2.\"U_DistNumber\" = \"OBTN\".\"DistNumber\" \n");
            query.Append("                         AND T2.\"U_Project\" = \"OWHS\".\"U_BDOSPrjCod\" \n");
            query.Append("               LEFT JOIN (SELECT MAX(\"@BDOSDEPACR\".\"U_DocDate\") AS \"LastDeprDocDate\", \n");
            query.Append("                                 \"@BDOSDEPAC1\".\"U_Project\", \n");
            query.Append("                                 \"@BDOSDEPAC1\".\"U_DistNumber\", \n");
            query.Append("                                 \"@BDOSDEPAC1\".\"U_ItemCode\" \n");
            query.Append("                          FROM   \"@BDOSDEPACR\" \n");
            query.Append("                                 INNER JOIN \"@BDOSDEPAC1\" \n");
            query.Append("                                         ON \"@BDOSDEPACR\".\"DocEntry\" = \"@BDOSDEPAC1\".\"DocEntry\" \n");
            query.Append("                          WHERE  \"Canceled\" = 'N' AND \"U_Retirement\" = 'Y' \n");
            query.Append($"                                 AND \"U_DocDate\" <= '{dateStr}' \n");
            query.Append("                          GROUP BY \"@BDOSDEPAC1\".\"U_Project\", \"@BDOSDEPAC1\".\"U_DistNumber\", \"@BDOSDEPAC1\".\"U_ItemCode\") AS T3 \n");
            query.Append("                      ON T3.\"U_ItemCode\" = \"OIBT\".\"ItemCode\" \n");
            query.Append("                         AND T3.\"U_DistNumber\" = \"OBTN\".\"DistNumber\" \n");
            query.Append("                         AND T3.\"U_Project\" = \"OWHS\".\"U_BDOSPrjCod\" \n");
            query.Append("               LEFT JOIN (SELECT Count(DISTINCT \"@BDOSDEPAC1\".\"DocEntry\") AS \"DocEntry\", \n");
            query.Append("                                 \"@BDOSDEPAC1\".\"U_DistNumber\", \n");
            query.Append("                                 \"@BDOSDEPAC1\".\"U_ItemCode\" \n");
            query.Append("                          FROM   \"@BDOSDEPACR\" \n");
            query.Append("                                 INNER JOIN \"@BDOSDEPAC1\" \n");
            query.Append("                                         ON \"@BDOSDEPACR\".\"DocEntry\" = \"@BDOSDEPAC1\".\"DocEntry\" \n");
            query.Append("                          WHERE  \"Canceled\" = 'N' \n");
            query.Append($"                                 AND \"U_DocDate\" <= '{dateStr}' \n");
            query.Append("                          GROUP BY \"@BDOSDEPAC1\".\"U_DistNumber\", \"@BDOSDEPAC1\".\"U_ItemCode\") AS T4 \n");
            query.Append("                      ON T4.\"U_ItemCode\" = \"OIBT\".\"ItemCode\" \n");
            query.Append("                         AND T4.\"U_DistNumber\" = \"OBTN\".\"DistNumber\" \n");
            query.Append("          WHERE \"OITM\".\"U_BDOSUsLife\" > 0 AND \"OBTN\".\"Quantity\" > 0 AND \"OIBT\".\"Quantity\" > 0 \n");
            query.Append("        ) AS T0 \n");
            query.Append("ORDER BY T0.\"ItemCode\", T0.\"DistNumber\", \n");
            query.Append("T0.\"LastDeprDocDate\" DESC");

            return query;
        }

        static StringBuilder getQueryForDepreciation(string dateStr)
        {
            StringBuilder query = new StringBuilder();

            query.Append("SELECT T0.*, \n");
            query.Append("       CASE \n");
            query.Append("         WHEN T0.\"AlreadyDepreciatedAmt\" = 0 THEN \n");
            query.Append("         T0.\"PurchaseCost\" / T0.\"UsefulLife\" \n");
            query.Append("         ELSE 0 \n");
            query.Append("       END                                                  AS \"DepreciationAmt\", \n");
            query.Append("       T0.\"PurchaseCost\" - (T0.\"AccumulatedDepreciationAmt\" * T0.\"Coefficient\") AS \"NetBookValue\", \n");
            query.Append("       T0.\"UsefulLife\" - T0.\"AllDeprDocQty\"             AS \"RemainingLife\" \n");
            query.Append("FROM   (SELECT \"OBTN\".\"DistNumber\", \n");
            query.Append("               \"OIBT\".\"WhsCode\", \n");
            query.Append("               \"OWHS\".\"WhsName\", \n");
            query.Append("               \"OWHS\".\"U_BDOSPrjCod\"                                        AS \"PrjCode\", \n");
            query.Append("               \"OIBT\".\"InDate\", \n");
            query.Append("               \"OIBT\".\"BaseType\", \n");
            query.Append("               \"T3\".\"LastDeprDocDate\", \n");
            query.Append("               \"OIBT\".\"ItemCode\", \n");
            query.Append("               \"OITM\".\"ItemName\", \n");
            query.Append("               \"OITM\".\"ItmsGrpCod\", \n");
            query.Append("               \"OITB\".\"ItmsGrpNam\", \n");
            query.Append("               \"OIBT\".\"Quantity\", \n");
            query.Append("               \"OIBT\".\"QuantityAll\", \n");
            query.Append("               \"OIBT\".\"Quantity\" / \"OIBT\".\"QuantityAll\"                 AS \"Coefficient\", \n");
            query.Append("               \"OITM\".\"U_BDOSUsLife\"                                        AS \"UsefulLife\", \n");
            query.Append("               CASE WHEN T1.\"U_DeprAmt\" IS NULL THEN 0 ELSE T1.\"U_DeprAmt\" END        AS \"AccumulatedDepreciationAmt\", \n");
            query.Append("               CASE WHEN T2.\"U_DeprAmt\" IS NULL THEN 0 ELSE T2.\"U_DeprAmt\" END        AS \"AlreadyDepreciatedAmt\", \n");
            query.Append("               T2.\"DepreciationDocEntry\", \n");
            query.Append("               \"OBTN\".\"CostTotal\" / \"OBTN\".\"Quantity\"                             AS \"PurchasePrice\", \n");
            query.Append("               ( \"OBTN\".\"CostTotal\" / \"OBTN\".\"Quantity\" ) * \"OIBT\".\"Quantity\" AS \"PurchaseCost\", \n");
            query.Append("               CASE WHEN T4.\"DocEntry\" IS NULL THEN 0 ELSE T4.\"DocEntry\" END AS \"AllDeprDocQty\" \n");
            query.Append("        FROM   (\n");
            query.Append("SELECT B2.*, \n");
            query.Append("       B3.\"Quantity\", \n");
            query.Append("       B3.\"QuantityAll\" \n");
            query.Append("FROM   (SELECT B0.\"SysNumber\", \n");
            query.Append("               B0.\"ItemCode\", \n");
            query.Append("               B0.\"BatchNum\", \n");
            query.Append("               B0.\"WhsCode\", \n");
            query.Append("               B1.\"BaseType\"     AS \"BaseType\", \n");
            query.Append("               Min(B1.\"DocDate\") AS \"InDate\" \n");
            query.Append("        FROM   \"OIBT\" B0 \n");
            query.Append("               INNER JOIN \"IBT1\" B1 \n");
            query.Append("                       ON B0.\"ItemCode\" = B1.\"ItemCode\" \n");
            query.Append("                          AND B0.\"BatchNum\" = B1.\"BatchNum\" \n");
            query.Append("                          AND B0.\"WhsCode\" = B1.\"WhsCode\" \n");
            query.Append("        WHERE B1.\"BaseType\" IN(18, 67) \n");
            query.Append("               AND B1.\"Direction\" = 0 \n");
            query.Append($"               AND B1.\"DocDate\" <= '{dateStr}' \n");
            query.Append("        GROUP  BY B0.\"SysNumber\", \n");
            query.Append("                  B0.\"ItemCode\", \n");
            query.Append("                  B0.\"BatchNum\", \n");
            query.Append("                  B0.\"WhsCode\", \n");
            query.Append("                  B1.\"BaseType\", \n");
            query.Append("                  B1.\"DocDate\" \n");
            query.Append("        ORDER  BY B1.\"DocDate\") B2 \n");
            query.Append("       LEFT JOIN (SELECT DISTINCT \n");
            query.Append("                                   B0.\"SysNumber\", \n");
            query.Append("                                   B0.\"ItemCode\", \n");
            query.Append("                                   B0.\"BatchNum\", \n");
            query.Append("                                   B0.\"WhsCode\", \n");
            query.Append("                                   Sum(CASE WHEN B1.\"Direction\" = 0 THEN B1.\"Quantity\" ELSE( -1 ) * B1.\"Quantity\" END) OVER(PARTITION BY B0.\"ItemCode\", B0.\"BatchNum\", B0.\"SysNumber\", B0.\"WhsCode\") AS \"Quantity\", \n");
            query.Append("                                   Sum(CASE WHEN B1.\"Direction\" = 0 THEN B1.\"Quantity\" ELSE( -1 ) * B1.\"Quantity\" END) OVER(PARTITION BY B0.\"ItemCode\", B0.\"BatchNum\", B0.\"SysNumber\") AS \"QuantityAll\" \n");
            query.Append("                   FROM   \"OIBT\" B0 \n");
            query.Append("                          INNER JOIN \"IBT1\" B1 \n");
            query.Append("                                  ON B0.\"ItemCode\" = B1.\"ItemCode\" \n");
            query.Append("                                     AND B0.\"BatchNum\" = B1.\"BatchNum\" \n");
            query.Append("                                     AND B0.\"WhsCode\" = B1.\"WhsCode\" \n");
            query.Append($"                   WHERE  B1.\"DocDate\" <= '{dateStr}') B3 \n");
            query.Append("               ON B2.\"SysNumber\" = B3.\"SysNumber\" \n");
            query.Append("                  AND B2.\"ItemCode\" = B3.\"ItemCode\" \n");
            query.Append("                  AND B2.\"BatchNum\" = B3.\"BatchNum\" \n");
            query.Append("                  AND B2.\"WhsCode\" = B3.\"WhsCode\" \n");
            query.Append("ORDER  BY B2.\"InDate\") \"OIBT\" \n");
            query.Append("               INNER JOIN \"OBTN\" \n");
            query.Append("                       ON \"OIBT\".\"ItemCode\" = \"OBTN\".\"ItemCode\" \n");
            query.Append("                          AND \"OIBT\".\"SysNumber\" = \"OBTN\".\"SysNumber\" \n");
            query.Append("                          AND \"OIBT\".\"BatchNum\" = \"OBTN\".\"DistNumber\" \n");
            query.Append("               INNER JOIN \"OWHS\" \n");
            query.Append("                       ON \"OIBT\".\"WhsCode\" = \"OWHS\".\"WhsCode\" \n");
            query.Append("               INNER JOIN \"OITM\" \n");
            query.Append("                       ON \"OIBT\".\"ItemCode\" = \"OITM\".\"ItemCode\" \n");
            query.Append("               INNER JOIN \"OITB\" \n");
            query.Append("                       ON \"OITM\".\"ItmsGrpCod\" = \"OITB\".\"ItmsGrpCod\" \n");
            query.Append("                          AND \"OITB\".\"U_BDOSFxAs\" = 'Y' \n");
            query.Append("               LEFT JOIN (SELECT --\"@BDOSDEPAC1\".\"U_Project\", \n");
            query.Append("                                 \"@BDOSDEPAC1\".\"U_DistNumber\", \n");
            query.Append("                                 \"@BDOSDEPAC1\".\"U_ItemCode\", \n");
            query.Append("                                 SUM(\"@BDOSDEPAC1\".\"U_DeprAmt\") AS \"U_DeprAmt\" \n");
            query.Append("                          FROM   \"@BDOSDEPAC1\" \n");
            query.Append("                                 INNER JOIN \"@BDOSDEPACR\" \n");
            query.Append("                                         ON \"@BDOSDEPAC1\".\"DocEntry\" = \"@BDOSDEPACR\".\"DocEntry\" \n");
            query.Append("                          WHERE  \"@BDOSDEPACR\".\"Canceled\" = 'N' AND \"@BDOSDEPACR\".\"U_Retirement\" = 'N' \n");
            query.Append($"                                 AND \"@BDOSDEPACR\".\"U_AccrMnth\" <= '{dateStr}' \n");
            query.Append("                          GROUP BY /*\"@BDOSDEPAC1\".\"U_Project\",*/ \"@BDOSDEPAC1\".\"U_DistNumber\", \"@BDOSDEPAC1\".\"U_ItemCode\") AS T1 \n");
            query.Append("                      ON T1.\"U_ItemCode\" = \"OIBT\".\"ItemCode\" \n");
            query.Append("                         AND T1.\"U_DistNumber\" = \"OBTN\".\"DistNumber\" \n");
            query.Append("                         --AND T1.\"U_Project\" = \"OWHS\".\"U_BDOSPrjCod\" \n");
            query.Append("               LEFT JOIN (SELECT \"@BDOSDEPACR\".\"DocEntry\" AS \"DepreciationDocEntry\", \n");
            query.Append("                                 \"@BDOSDEPAC1\".\"U_Project\", \n");
            query.Append("                                 \"@BDOSDEPAC1\".\"U_DistNumber\", \n");
            query.Append("                                 \"@BDOSDEPAC1\".\"U_ItemCode\", \n");
            query.Append("                                 \"@BDOSDEPAC1\".\"U_DeprAmt\" \n");
            query.Append("                          FROM   \"@BDOSDEPAC1\" \n");
            query.Append("                                 INNER JOIN \"@BDOSDEPACR\" \n");
            query.Append("                                         ON \"@BDOSDEPAC1\".\"DocEntry\" = \"@BDOSDEPACR\".\"DocEntry\" \n");
            query.Append("                          WHERE  \"@BDOSDEPACR\".\"Canceled\" = 'N' AND \"@BDOSDEPACR\".\"U_Retirement\" = 'N' \n");
            query.Append($"                                 AND \"@BDOSDEPACR\".\"U_AccrMnth\" = '{dateStr}') AS T2 \n");
            query.Append("                      ON T2.\"U_ItemCode\" = \"OIBT\".\"ItemCode\" \n");
            query.Append("                         AND T2.\"U_DistNumber\" = \"OBTN\".\"DistNumber\" \n");
            query.Append("                         AND T2.\"U_Project\" = \"OWHS\".\"U_BDOSPrjCod\" \n");
            query.Append("               LEFT JOIN (SELECT MAX(\"@BDOSDEPACR\".\"U_DocDate\") AS \"LastDeprDocDate\", \n");
            query.Append("                                 \"@BDOSDEPAC1\".\"U_Project\", \n");
            query.Append("                                 \"@BDOSDEPAC1\".\"U_DistNumber\", \n");
            query.Append("                                 \"@BDOSDEPAC1\".\"U_ItemCode\" \n");
            query.Append("                          FROM   \"@BDOSDEPACR\" \n");
            query.Append("                                 INNER JOIN \"@BDOSDEPAC1\" \n");
            query.Append("                                         ON \"@BDOSDEPACR\".\"DocEntry\" = \"@BDOSDEPAC1\".\"DocEntry\" \n");
            query.Append("                          WHERE  \"Canceled\" = 'N' AND \"U_Retirement\" = 'N' \n");
            query.Append($"                                 AND \"U_DocDate\" <= '{dateStr}' \n");
            query.Append("                          GROUP BY \"@BDOSDEPAC1\".\"U_Project\", \"@BDOSDEPAC1\".\"U_DistNumber\", \"@BDOSDEPAC1\".\"U_ItemCode\") AS T3 \n");
            query.Append("                      ON T3.\"U_ItemCode\" = \"OIBT\".\"ItemCode\" \n");
            query.Append("                         AND T3.\"U_DistNumber\" = \"OBTN\".\"DistNumber\" \n");
            query.Append("                         AND T3.\"U_Project\" = \"OWHS\".\"U_BDOSPrjCod\" \n");
            query.Append("               LEFT JOIN (SELECT Count(DISTINCT \"@BDOSDEPAC1\".\"DocEntry\") AS \"DocEntry\", \n");
            query.Append("                                 \"@BDOSDEPAC1\".\"U_DistNumber\", \n");
            query.Append("                                 \"@BDOSDEPAC1\".\"U_ItemCode\" \n");
            query.Append("                          FROM   \"@BDOSDEPACR\" \n");
            query.Append("                                 INNER JOIN \"@BDOSDEPAC1\" \n");
            query.Append("                                         ON \"@BDOSDEPACR\".\"DocEntry\" = \"@BDOSDEPAC1\".\"DocEntry\" \n");
            query.Append("                          WHERE  \"Canceled\" = 'N' \n");
            query.Append($"                                 AND \"U_DocDate\" <= '{dateStr}' \n");
            query.Append("                          GROUP BY \"@BDOSDEPAC1\".\"U_DistNumber\", \"@BDOSDEPAC1\".\"U_ItemCode\") AS T4 \n");
            query.Append("                      ON T4.\"U_ItemCode\" = \"OIBT\".\"ItemCode\" \n");
            query.Append("                         AND T4.\"U_DistNumber\" = \"OBTN\".\"DistNumber\" \n");
            query.Append("          WHERE \"OITM\".\"U_BDOSUsLife\" > 0 AND \"OBTN\".\"Quantity\" > 0 AND \"OIBT\".\"Quantity\" > 0 \n");
            query.Append($"           AND (NEXT_DAY(LAST_DAY(\"OIBT\".\"InDate\")) < '{dateStr}' OR (\"OIBT\".\"BaseType\" = 67 AND LAST_DAY(\"OIBT\".\"InDate\") = '{dateStr}')) \n");
            query.Append("        ) AS T0 \n");
            query.Append("ORDER BY T0.\"ItemCode\", T0.\"DistNumber\", T0.\"InDate\", T0.\"LastDeprDocDate\" DESC");

            return query;
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
