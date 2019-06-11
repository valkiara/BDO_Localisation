using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BDO_Localisation_AddOn
{
    class BDOSDepreciationAccrualWizard
    {
        public static void createForm(out string errorText)
        {
            errorText = null;
            Dictionary<string, object> formItems;
            string itemName;
            SAPbouiCOM.Columns oColumns;
            SAPbouiCOM.Column oColumn;

            SAPbouiCOM.DataTable oDataTable;

            bool multiSelection;

            int left_s = 5;
            int left_s1 = 310;

            int top = 10;
            int height = 15;

            //ფორმის აუცილებელი თვისებები
            Dictionary<string, object> formProperties = new Dictionary<string, object>();
            formProperties.Add("UniqueID", "BDOSDepAccrForm");
            formProperties.Add("BorderStyle", SAPbouiCOM.BoFormBorderStyle.fbs_Sizable);
            formProperties.Add("Title", BDOSResources.getTranslate("DepreciationAccruingWizard"));
            formProperties.Add("Left", 558);
            formProperties.Add("ClientWidth", 800);
            formProperties.Add("Top", 335);
            formProperties.Add("ClientHeight", 600);

            SAPbouiCOM.Form oForm;
            bool newForm;
            bool formExist = FormsB1.createForm(formProperties, out oForm, out newForm, out errorText);

            if (formExist == true)
            {
                if (newForm)
                {


                    formItems = new Dictionary<string, object>();
                    itemName = "DeprMonthS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", 150);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("DocumentPostingDate"));
                    formItems.Add("FromPane", 0);
                    formItems.Add("ToPane", 0);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "DeprMonth";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_DATE);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("Length", 1);
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Left", left_s + 25 + 100);
                    formItems.Add("Width", 150);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    //formItems.Add("ValueEx", DateTime.Now.ToString("yyyyMMdd"));
                    formItems.Add("FromPane", 0);
                    formItems.Add("ToPane", 0);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }
                    
                    top = top + height + 10;

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
                    formItems.Add("Width", 150);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("InvoiceDepreciation"));
                    formItems.Add("ValOn", "Y");
                    formItems.Add("ValOff", "N");
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("Selected", true);
                    formItems.Add("FromPane", 0);
                    formItems.Add("ToPane", 0);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
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
                    formItems.Add("Left", left_s + 150);
                    formItems.Add("Width", 150);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("StockDepreciation"));
                    formItems.Add("GroupWith", "InvDepr");
                    formItems.Add("ValOn", "Y");
                    formItems.Add("ValOff", "N");
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("FromPane", 0);
                    formItems.Add("ToPane", 0);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }



                    top = top + height + 10;

                    //საკონტროლო პანელი
                    formItems = new Dictionary<string, object>();
                    itemName = "fillMTR"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("fillMTR"));
                    //formItems.Add("SetAutoManaged", true);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    //საკონტროლო პანელი
                    formItems = new Dictionary<string, object>();
                    itemName = "CreatDoc"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left_s + 100 + 5);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("CreateDepreciationDocument"));
                    //formItems.Add("SetAutoManaged", true);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    top = top + height + 10;

                    formItems = new Dictionary<string, object>();
                    itemName = "ItemsMTR"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", 600);
                    formItems.Add("Top", top);
                    formItems.Add("Height", 550);
                    formItems.Add("UID", itemName);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("AffectsFormMode", false);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }
                                                                          
                    SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("ItemsMTR").Specific;

                    oColumns = oMatrix.Columns;

                    SAPbouiCOM.LinkedButton oLink;
                    oDataTable = oForm.DataSources.DataTables.Add("ItemsMTR");

                    oDataTable.Columns.Add("LineNum", SAPbouiCOM.BoFieldsType.ft_Integer, 50); //ინდექსი 
                    oDataTable.Columns.Add("CheckBox", SAPbouiCOM.BoFieldsType.ft_Text, 1); // 
                    oDataTable.Columns.Add("DocType", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 50);
                    oDataTable.Columns.Add("DocEntry", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 50);
                    oDataTable.Columns.Add("Project", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 50);
                    oDataTable.Columns.Add("ItemGrp", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 50);
                    oDataTable.Columns.Add("ItemCode", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 50);
                    oDataTable.Columns.Add("ItemName", SAPbouiCOM.BoFieldsType.ft_Text, 50);
                    oDataTable.Columns.Add("DistNumber", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 50);
                    oDataTable.Columns.Add("UseLife", SAPbouiCOM.BoFieldsType.ft_Quantity);
                    oDataTable.Columns.Add("AlrDeprAmt", SAPbouiCOM.BoFieldsType.ft_Sum);
                    oDataTable.Columns.Add("Quantity", SAPbouiCOM.BoFieldsType.ft_Quantity);
                    oDataTable.Columns.Add("APCost", SAPbouiCOM.BoFieldsType.ft_Sum);
                    oDataTable.Columns.Add("NtBookVal", SAPbouiCOM.BoFieldsType.ft_Sum);
                    oDataTable.Columns.Add("DeprAmt", SAPbouiCOM.BoFieldsType.ft_Sum);
                    oDataTable.Columns.Add("CurMnthAmt", SAPbouiCOM.BoFieldsType.ft_Sum);
                    
                    for (int count = 0; count < oDataTable.Columns.Count; count++)
                    {
                        var column = oDataTable.Columns.Item(count);
                        string columnName = column.Name;

                        if (columnName == "LineNum")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = "#";
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind("ItemsMTR", columnName);
                            oColumn.AffectsFormMode = false;
                        }
                        else if (columnName == "CheckBox")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
                            oColumn.TitleObject.Caption = "";
                            oColumn.Editable = true;
                            oColumn.ValOff = "N";
                            oColumn.ValOn = "Y";
                            oColumn.DataBind.Bind("ItemsMTR", columnName);
                            oColumn.AffectsFormMode = false;
                        }                        
                        
                        else if (columnName == "Project")
                        {
                            oColumn = oColumns.Add("Project", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Project");
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind("ItemsMTR", columnName);
                            oLink = oColumn.ExtendedObject;
                            oLink.LinkedObjectType = "63";
                        }
                        else if (columnName == "ItemGrp")
                        {                            
                            oColumn = oColumns.Add("ItemGrp", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("ItmsGrpCod");
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind("ItemsMTR", columnName);
                            oLink = oColumn.ExtendedObject;
                            oLink.LinkedObjectType = "52";
                        }
                        else if (columnName == "ItemCode")
                        {
                            oColumn = oColumns.Add("ItemCode", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("ItemCode");
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind("ItemsMTR", columnName);
                            oLink = oColumn.ExtendedObject;
                            oLink.LinkedObjectType = "4";
                        }
                        else if (columnName == "DocEntry")
                        {
                            oColumn = oColumns.Add("DocEntry", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("DocEntry");
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind("ItemsMTR", columnName);
                        }
                        
                        else if (columnName == "DistNumber")
                        {
                            oColumn = oColumns.Add("DistNumber", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("DistNumber");
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind("ItemsMTR", columnName);
                            //oLink = oColumn.ExtendedObject;
                            //oLink.LinkedObjectType = "10000044";

                        }
                        else
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate(columnName);
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind("ItemsMTR", columnName);
                            oColumn.AffectsFormMode = false;
                        }

                    }
                    oMatrix.Clear();
                    oMatrix.LoadFromDataSource();
                    oMatrix.AutoResizeColumns();
                }

                resizeItems(oForm);
                oForm.Visible = true;
                oForm.Select();
            }
        }

        public static void matrixColumnSetLinkedObjectTypeInvoicesMTR(SAPbouiCOM.Form oForm, SAPbouiCOM.ItemEvent pVal, out string errorText)
        {
            errorText = null;

            try
            {
                if (pVal.ColUID == "DocEntry")
                {
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED & pVal.BeforeAction == true)
                    {
                        SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("ItemsMTR").Specific));

                        SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("ItemsMTR");
                        string docType = oDataTable.GetValue("DocType", pVal.Row - 1);

                        SAPbouiCOM.Column oColumn;

                        if (docType == "13")
                        {
                            oColumn = oMatrix.Columns.Item(pVal.ColUID);
                            SAPbouiCOM.LinkedButton oLink = oColumn.ExtendedObject;
                            oLink.LinkedObjectType = docType; //ARInvoice
                        }
                        if (docType == "60")
                        {
                            oColumn = oMatrix.Columns.Item(pVal.ColUID);
                            SAPbouiCOM.LinkedButton oLink = oColumn.ExtendedObject;
                            oLink.LinkedObjectType = docType; //ARInvoice
                        }
                    }
                }
                else
                {

                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
            finally
            {
                GC.Collect();
            }
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
                    resizeItems(oForm);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                {
                    SAPbouiCOM.IChooseFromListEvent oCFLEvento = null;
                    oCFLEvento = ((SAPbouiCOM.IChooseFromListEvent)(pVal));

                    chooseFromList(oForm, pVal.BeforeAction, oCFLEvento, out errorText);
                }

                if (pVal.ItemUID == "ItemsMTR")
                {
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED)
                    {
                        matrixColumnSetLinkedObjectTypeInvoicesMTR(oForm, pVal, out errorText);
                    }
                }

                if (pVal.ItemUID == "DeprMonth" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED & pVal.BeforeAction == false)
                {
                    oForm.Freeze(true);
                    itemPressed(oForm, pVal, out errorText);
                    oForm.Freeze(false);
                }
                if (pVal.ItemUID == "DeprMonth" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_VALIDATE && pVal.ItemChanged)
                {
                    oForm.Freeze(true);
                    itemPressed(oForm, pVal, out errorText);
                    oForm.Freeze(false);
                }
               

                if (pVal.ItemUID == "fillMTR" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.BeforeAction == false)
                {               
                    fillMTRItems(oForm);
                }

                if (pVal.ItemUID == "CreatDoc" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.BeforeAction == false)
                {
                    CreateDocuments(oForm);
                    fillMTRItems(oForm);
                }
                
            }
        }

        private static void itemPressed(SAPbouiCOM.Form oForm, SAPbouiCOM.ItemEvent pVal, out string errorText)
        {
            errorText = null;

            try
            {
                if (pVal.ItemUID == "DeprMonth")
                {
                    SAPbouiCOM.EditText oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("DeprMonth").Specific;
                    DateTime AccrMnth = DateTime.ParseExact(oEditText.Value, "yyyyMMdd", null);
                    
                    AccrMnth = new DateTime(AccrMnth.Year, AccrMnth.Month, 1);
                    AccrMnth =  AccrMnth.AddMonths(1).AddDays(-1);
                    oEditText.Value = AccrMnth.ToString("yyyyMMdd");
                }
            }
            catch (Exception ex)
            {
                int errCode;
                string errMsg;

                Program.oCompany.GetLastError(out errCode, out errMsg);
                errorText = BDOSResources.getTranslate("ErrorDescription") + " : " + errMsg + "! " + BDOSResources.getTranslate("Code") + " : " + errCode + "! " + BDOSResources.getTranslate("OtherInfo") + " : " + ex.Message;
            }
            finally
            {
                GC.Collect();
            }
        }


        private static int CreateDocuments(SAPbouiCOM.Form oForm)
        {
            string errorText = null;

            SAPbobsCOM.CompanyService oCompanyService = null;
            SAPbobsCOM.GeneralService oGeneralService = null;
            SAPbobsCOM.GeneralData oGeneralData = null;
            oCompanyService = Program.oCompany.GetCompanyService();
            oGeneralService = oCompanyService.GetGeneralService("UDO_F_BDOSDEPACR_D");
            oGeneralData = ((SAPbobsCOM.GeneralData)(oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)));

            DateTime AccrMnth = Convert.ToDateTime(DateTime.ParseExact(oForm.Items.Item("DeprMonth").Specific.Value, "yyyyMMdd", CultureInfo.InvariantCulture));

            oGeneralData.SetProperty("U_AccrMnth", AccrMnth);
            oGeneralData.SetProperty("U_DocDate", AccrMnth);

            SAPbobsCOM.GeneralDataCollection oChildren = null;

            SAPbouiCOM.DataTable DepreciationLines = oForm.DataSources.DataTables.Item("ItemsMTR");

            oChildren = oGeneralData.Child("BDOSDEPAC1");

            for (int i = 0; i < DepreciationLines.Rows.Count; i++)
            {
                string CheckBox = DepreciationLines.GetValue("CheckBox", i);
                if (CheckBox == "Y")
                {
                    double CurMnthAmt = DepreciationLines.GetValue("CurMnthAmt", i);
                    if (CurMnthAmt == 0)
                    {
                        SAPbobsCOM.GeneralData oChild = oChildren.Add();
                        oChild.SetProperty("U_ItemCode", DepreciationLines.GetValue("ItemCode", i));
                        oChild.SetProperty("U_DistNumber", DepreciationLines.GetValue("DistNumber", i));
                        oChild.SetProperty("U_Project", DepreciationLines.GetValue("Project", i));
                        oChild.SetProperty("U_Project", DepreciationLines.GetValue("Project", i));
                        oChild.SetProperty("U_Quantity", DepreciationLines.GetValue("Quantity", i));
                        oChild.SetProperty("U_DeprAmt", DepreciationLines.GetValue("DeprAmt", i));
                    }
                    else
                    {
                        string ItemCode = DepreciationLines.GetValue("ItemCode", i);
                        string DistNumber = DepreciationLines.GetValue("DistNumber", i);

                        Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("DocumentAlreadyCreatedForBatchNumber") + " " + DistNumber + " (" + ItemCode+")", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    }
                }
            }

            int DocumentCreated = 1;

            if (oChildren.Count==0)
            {
                return 0;
            }

            try
            {
                CommonFunctions.StartTransaction();

                var response = oGeneralService.Add(oGeneralData);
                int docEntry = response.GetProperty("DocEntry");

                if (docEntry > 0)
                {
                    DataTable JrnLinesDT = BDOSDepreciationAccrualDocument.createAdditionalEntries(null, oGeneralData, 0, AccrMnth, "", null, null);

                    BDOSDepreciationAccrualDocument.JrnEntry(docEntry.ToString(), docEntry.ToString(), AccrMnth, JrnLinesDT, "", out errorText);

                    if (errorText != null)
                    {
                        DocumentCreated = 0;
                        CommonFunctions.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                        Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("DocumentNotCreated") + ". " + BDOSResources.getTranslate("ReasonIs") + ": " + errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
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
                Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("DocumentNotCreated") + ". " + BDOSResources.getTranslate("ReasonIs") + ": " + Ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

            return DocumentCreated;

        }

        public static void resizeItems(SAPbouiCOM.Form oForm)
        {
            try
            {
                SAPbouiCOM.Item oMatrixItem = oForm.Items.Item("ItemsMTR");

                oMatrixItem.Height = oForm.Height - 220;
                oMatrixItem.Width = oForm.Width - 20;
            }
            catch
            {
            }
        }

        public static string BatchDepreciaionQuery(DateTime DeprMonth, string ItemCodes, string BatchNumbers, string WhsCode, bool isInvoice = false)
        {

            

            string query = @"select 
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
                             ""DepcAcc"".""DeprAmt"" as ""DeprAmt"",
                             ""DepcAcc"".""CurrDeprAmt"" as ""CurrDeprAmt"",
	                         ""DepcAcc"".""DeprQty"" as ""DeprQty"", 
	                         ""FinTable"".""NtBookVal"" as ""NtBookVal"",
	                         ""FinTable"".""Quantity"" as ""Quantity""

                             from (select
	                         ""OIVL"".""LocCode"","
                            +
                            (isInvoice ? @" ""OBVL"".""DocEntry"", ""OBVL"".""DocType"", " : "")
                            +
                             @"""OWHS"".""U_BDOSPrjCod"" as ""Project"",
                             ""OITM"".""ItmsGrpCod"" as ""ItemGrp"", 	                         
                             ""OBVL"".""ItemCode"",
	 	                     ""OITM"".""ItemName"",
                             ""OITB"".""U_BDOSUsLife"" as ""UseLife"",
	                         ""OBVL"".""DistNumber"",
                             ""OBTN"".""InDate"",
	                         ""OBTN"".""CostTotal"" / ""OBTN"".""Quantity"" as ""APCost"",
                             
	                         SUM( ""OBTN"".""CostTotal""/""OBTN"".""Quantity"" * ""OBVL"".""Quantity""*case when ""OBVL"".""TransValue"">0 then 1 else -1 end) as ""NtBookVal"",
	                         SUM(""OBVL"".""Quantity""*case when ""OBVL"".""TransValue"">0 then 1 else -1 end ) as ""Quantity"" 
                        from ""OBVL""
                        
                        inner join ""OBTN"" on ""OBTN"".""DistNumber"" = ""OBVL"".""DistNumber"" and ""OBTN"".""ItemCode"" = ""OBVL"".""ItemCode"" 
                        and #ItemFilter#
                        and #BatchFilter#
                        and  ADD_MONTHS(NEXT_DAY(LAST_DAY(""OBTN"".""InDate"")),-1)< ADD_MONTHS(NEXT_DAY(LAST_DAY('" + DeprMonth.ToString("yyyyMMdd") + @"')),-1)
                        inner join ""OITM"" on  ""OBVL"".""ItemCode"" = ""OITM"".""ItemCode""
                        inner join ""OITB"" on  ""OITB"".""ItmsGrpCod"" = ""OITM"".""ItmsGrpCod"" and ""OITB"".""U_BDOSFxAs""='Y'
                        inner join ""OIVL"" on ""OIVL"".""ItemCode"" = ""OBVL"".""ItemCode"" 
                        and ""OIVL"".""CreatedBy"" = ""OBVL"".""DocEntry"" 
                        and ""OIVL"".""TransType"" = ""OBVL"".""DocType"" 
                        and case when ""OIVL"".""ParentID"" = -1 
                        then ""OIVL"".""ParentID"" 
                        else ""OIVL"".""TransType"" 
                        END = ""OBVL"".""BaseType"" 
                        inner join ""OWHS"" on ""OWHS"".""WhsCode"" = ""OIVL"".""LocCode"" 
                        
                        group by ""OIVL"".""LocCode"","
                            +
                            (isInvoice ? @" ""OBVL"".""DocEntry"", ""OBVL"".""DocType"", " : "")
                            +
                             @"""OWHS"".""U_BDOSPrjCod"",
                            ""OITM"".""ItmsGrpCod"",
	                         ""OBVL"".""ItemCode"",
	                         ""OITM"".""ItemName"",
                            ""OITB"".""U_BDOSUsLife"",
	                         ""OBVL"".""DistNumber"",
                            ""OBTN"".""InDate"",
	                         ""OBTN"".""CostTotal"" / ""OBTN"".""Quantity"") as ""FinTable""
	                         left  join (select
                        ""@BDOSDEPAC1"".""U_ItemCode"",
                        ""@BDOSDEPAC1"".""U_DistNumber"",
                        SUM(case when ISNULL(""@BDOSDEPAC1"".""U_Quantity"",0)=0 then 0 else ""@BDOSDEPAC1"".""U_DeprAmt""/""@BDOSDEPAC1"".""U_Quantity"" end) as ""DeprAmt"",

                        SUM(case when ""@BDOSDEPACR"".""U_AccrMnth"" = '" + DeprMonth.ToString("yyyyMMdd") + @"' 
                        then ""@BDOSDEPAC1"".""U_DeprAmt""
	                    else 0
                        end) as ""CurrDeprAmt"",

	                    SUM(""@BDOSDEPAC1"".""U_Quantity"") as ""DeprQty""
                        from ""@BDOSDEPAC1"" 
                        inner join ""@BDOSDEPACR"" on ""@BDOSDEPACR"".""DocEntry"" = ""@BDOSDEPAC1"".""DocEntry"" and ""@BDOSDEPACR"".""Canceled"" = 'N'group by ""@BDOSDEPAC1"".""U_ItemCode"",""@BDOSDEPAC1"".""U_DistNumber"") as ""DepcAcc"" 
                        on ""DepcAcc"".""U_ItemCode"" = ""FinTable"".""ItemCode"" 
                        and ""DepcAcc"".""U_DistNumber"" = ""FinTable"".""DistNumber"""
                         +
                            (isInvoice ? @" where ""FinTable"".""DocType""= 13 or ""FinTable"".""DocType""= 60 " : "")
                            +

                        @"";

            if(ItemCodes=="")
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
            DateTime DeprMonth = Convert.ToDateTime(DateTime.ParseExact(oForm.Items.Item("DeprMonth").Specific.Value, "yyyyMMdd", CultureInfo.InvariantCulture));
            bool isInvoice = oForm.Items.Item("InvDepr").Specific.Selected;

            string query = BatchDepreciaionQuery(DeprMonth, "", "","", isInvoice);

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRecordSet.DoQuery(query);
            int rowIndex = 0;

            while (!oRecordSet.EoF)
            {
                DateTime InDateStart = oRecordSet.Fields.Item("InDate").Value;
                DateTime InDateEnd = InDateStart.AddMonths(oRecordSet.Fields.Item("UseLife").Value);

                if (DeprMonth > InDateEnd || DeprMonth < InDateStart)
                {
                    oRecordSet.MoveNext();
                    continue;
                }

                decimal CurrDeprAmt = Convert.ToDecimal(oRecordSet.Fields.Item("CurrDeprAmt").Value, CultureInfo.InvariantCulture);
                int monthsApart = 12 * (InDateEnd.Year - DeprMonth.Year) + (InDateEnd.Month - DeprMonth.Month) + 1;

                monthsApart = Math.Abs(monthsApart);
                                
                decimal AlrDeprAmt = 0;
                
                AlrDeprAmt = Convert.ToDecimal(oRecordSet.Fields.Item("DeprAmt").Value)  * Convert.ToDecimal(oRecordSet.Fields.Item("Quantity").Value);
                decimal NtBookVal = Convert.ToDecimal(oRecordSet.Fields.Item("APCost").Value * oRecordSet.Fields.Item("Quantity").Value) - AlrDeprAmt;
                decimal DeprAmt = NtBookVal / monthsApart;
                
                oDataTable.Rows.Add();
                oDataTable.SetValue("LineNum", rowIndex, rowIndex + 1);
                oDataTable.SetValue("CheckBox", rowIndex, "Y"); // 
                oDataTable.SetValue("Project", rowIndex, oRecordSet.Fields.Item("Project").Value);
                oDataTable.SetValue("ItemGrp", rowIndex, oRecordSet.Fields.Item("ItemGrp").Value);
                oDataTable.SetValue("ItemCode", rowIndex, oRecordSet.Fields.Item("ItemCode").Value);
                oDataTable.SetValue("ItemName", rowIndex, oRecordSet.Fields.Item("ItemName").Value);
                oDataTable.SetValue("DistNumber", rowIndex, oRecordSet.Fields.Item("DistNumber").Value);
                oDataTable.SetValue("UseLife", rowIndex, monthsApart);                
                oDataTable.SetValue("Quantity", rowIndex, oRecordSet.Fields.Item("Quantity").Value);
                oDataTable.SetValue("APCost", rowIndex, oRecordSet.Fields.Item("APCost").Value * oRecordSet.Fields.Item("Quantity").Value);
                oDataTable.SetValue("AlrDeprAmt", rowIndex, Convert.ToDouble(AlrDeprAmt));
                oDataTable.SetValue("NtBookVal", rowIndex, Convert.ToDouble(NtBookVal));
                oDataTable.SetValue("DeprAmt", rowIndex, Convert.ToDouble(DeprAmt));
                oDataTable.SetValue("CurMnthAmt", rowIndex, Convert.ToDouble(CurrDeprAmt));
                if(isInvoice)
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
            oMatrix.AutoResizeColumns();
            oForm.Update();
            oForm.Freeze(false);

        }

        private static void chooseFromList(SAPbouiCOM.Form oForm, bool BeforeAction, SAPbouiCOM.IChooseFromListEvent oCFLEvento, out string errorText)
        {
            errorText = null;

            string sCFL_ID = oCFLEvento.ChooseFromListUID;
            SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(sCFL_ID);

            SAPbouiCOM.DataTable oDataTable = null;
            oDataTable = oCFLEvento.SelectedObjects;

            if (BeforeAction == false)
            {
                if (oDataTable != null)
                {
                        
                }
            }
        }

        public static void addMenus(out string errorText)
        {
            errorText = null;

            SAPbouiCOM.MenuItem menuItem;
            SAPbouiCOM.MenuItem fatherMenuItem;
            SAPbouiCOM.MenuCreationParams oCreationPackage;

            try
            {
                fatherMenuItem = Program.uiApp.Menus.Item("1536");

                // Add a pop-up menu item
                oCreationPackage = (SAPbouiCOM.MenuCreationParams)Program.uiApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                oCreationPackage.Checked = false;
                oCreationPackage.Enabled = true;
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "BDOSDepAccrForm";
                oCreationPackage.String = BDOSResources.getTranslate("DepreciationAccruingWizard");
                oCreationPackage.Position = fatherMenuItem.SubMenus.Count - 1;

                menuItem = fatherMenuItem.SubMenus.AddEx(oCreationPackage);
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
        }

    }
}
