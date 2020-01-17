using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace BDO_Localisation_AddOn
{
    class BDOSInterestAccrualWizard
    {
        public static void createForm()
        {
            string errorText;
            int formHeight = Program.uiApp.Desktop.Width;
            int formWidth = Program.uiApp.Desktop.Width;

            //ფორმის აუცილებელი თვისებები
            Dictionary<string, object> formProperties = new Dictionary<string, object>();
            formProperties.Add("UniqueID", "BDOSInterestAccrualWizard");
            formProperties.Add("BorderStyle", SAPbouiCOM.BoFormBorderStyle.fbs_Sizable);
            formProperties.Add("Title", BDOSResources.getTranslate("InterestAccrualWizard"));
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

                    int left_s = 6;
                    int left_e = 180;
                    int height = 15;
                    int top = 10;
                    int width_s = 160;
                    int width_e = 140;

                    FormsB1.addChooseFromList(oForm, false, "3", "BankCFL");
                    formItems = new Dictionary<string, object>();
                    itemName = "BankCodeS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("Bank"));
                    formItems.Add("LinkTo", "BankCodeE");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "BankCodeE"; //10 characters
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
                    formItems.Add("ChooseFromListUID", "BankCFL");
                    formItems.Add("ChooseFromListAlias", "BankCode");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "BankCodeLB"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    formItems.Add("Left", left_e - 20);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("LinkTo", "BankCodeE");
                    formItems.Add("LinkedObjectType", "3"); //Bank Codes

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    top += height + 1;

                    DateTime docDate = DateTime.Today;
                    string docDateTxt = docDate.ToString("yyyyMMdd");

                    formItems = new Dictionary<string, object>();
                    itemName = "DocDateS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("PostingDate"));
                    formItems.Add("LinkTo", "DocDateE");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "DocDateE";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_DATE);
                    formItems.Add("Length", 1);
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Left", left_e);
                    formItems.Add("Width", width_e);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("ValueEx", docDateTxt);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    top += height + 1;

                    //----------------------------------------------------------------------------------------------------------

                    top = top + 2 * height + 1;

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
                        throw new Exception(errorText);
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
                        throw new Exception(errorText);
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
                    itemName = "LoanMTR"; //10 characters
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

                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("LoanMTR").Specific;
                    oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
                    SAPbouiCOM.Columns oColumns = oMatrix.Columns;
                    SAPbouiCOM.Column oColumn;

                    SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Add("LoanMTR");

                    oDataTable.Columns.Add("CheckBox", SAPbouiCOM.BoFieldsType.ft_Text, 1);
                    oDataTable.Columns.Add("LineNum", SAPbouiCOM.BoFieldsType.ft_Integer, 50);
                    oDataTable.Columns.Add("DocEntry", SAPbouiCOM.BoFieldsType.ft_Integer, 11);
                    oDataTable.Columns.Add("U_CrLnCode", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 50);
                    oDataTable.Columns.Add("U_CrLnName", SAPbouiCOM.BoFieldsType.ft_Text, 100);
                    oDataTable.Columns.Add("U_StartDate", SAPbouiCOM.BoFieldsType.ft_Date);
                    oDataTable.Columns.Add("U_LnCurrCode", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 3);
                    oDataTable.Columns.Add("U_ExchngRate", SAPbouiCOM.BoFieldsType.ft_Rate);
                    oDataTable.Columns.Add("U_CrLnAcct", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 15);
                    oDataTable.Columns.Add("U_IntrstRate", SAPbouiCOM.BoFieldsType.ft_Percent);
                    oDataTable.Columns.Add("U_AccrDate", SAPbouiCOM.BoFieldsType.ft_Date);
                    oDataTable.Columns.Add("U_AccrDays", SAPbouiCOM.BoFieldsType.ft_Integer);
                    oDataTable.Columns.Add("U_CrLnAmtLC", SAPbouiCOM.BoFieldsType.ft_Sum);
                    oDataTable.Columns.Add("U_CrLnAmtFC", SAPbouiCOM.BoFieldsType.ft_Sum);
                    oDataTable.Columns.Add("U_IntAmtLC", SAPbouiCOM.BoFieldsType.ft_Sum);
                    oDataTable.Columns.Add("U_IntAmtFC", SAPbouiCOM.BoFieldsType.ft_Sum);
                    oDataTable.Columns.Add("U_ExpnsAcct", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 15);
                    oDataTable.Columns.Add("U_IntPblAcct", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 15);

                    string UID = "LoanMTR";
                    SAPbouiCOM.LinkedButton oLink;

                    oColumn = oColumns.Add("LineNum", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = "#";
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "LineNum");

                    oColumn = oColumns.Add("CheckBox", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
                    oColumn.TitleObject.Caption = "";
                    oColumn.Editable = true;
                    oColumn.ValOff = "N";
                    oColumn.ValOn = "Y";
                    oColumn.DataBind.Bind(UID, "CheckBox");

                    oColumn = oColumns.Add("DocEntry", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("DocEntry") + " (" + BDOSResources.getTranslate("InterestAccrualDocument") + ")";
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "DocEntry");
                    oLink = oColumn.ExtendedObject;
                    oLink.LinkedObjectType = "UDO_F_BDOSINAC_D"; //Interest Accrual Document

                    FormsB1.addChooseFromList(oForm, false, "UDO_F_BDOSCRLN_D", "CreditLineCodeCFL");
                    oColumn = oColumns.Add("CrLnCode", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON); //Credit Line Code
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("CreditLine");
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "U_CrLnCode");
                    oColumn.ChooseFromListUID = "CreditLineCodeCFL";
                    oColumn.ChooseFromListAlias = "Code";
                    oLink = oColumn.ExtendedObject;
                    oLink.LinkedObjectType = "UDO_F_BDOSCRLN_D"; //Credit Line Master Data

                    oColumn = oColumns.Add("CrLnName", SAPbouiCOM.BoFormItemTypes.it_EDIT); //Credit Line Name
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("CreditLineCode");
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "U_CrLnName");

                    oColumn = oColumns.Add("StartDate", SAPbouiCOM.BoFormItemTypes.it_EDIT); //Starting Date
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("StartingDate");
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "U_StartDate");

                    FormsB1.addChooseFromList(oForm, false, "37", "CurrencyCFL");
                    oColumn = oColumns.Add("LnCurrCode", SAPbouiCOM.BoFormItemTypes.it_EDIT); //Loan Currency Code
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("Currency");
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "U_LnCurrCode");
                    oColumn.ChooseFromListUID = "CurrencyCFL";
                    oColumn.ChooseFromListAlias = "CurrCode"; //Currency Codes

                    oColumn = oColumns.Add("ExchngRate", SAPbouiCOM.BoFormItemTypes.it_EDIT); //Exchange Rate
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("ExchangeRate");
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "U_ExchngRate");

                    FormsB1.addChooseFromList(oForm, false, "1", "CrLnAcctCFL");
                    oColumn = oColumns.Add("CrLnAcct", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON); //Credit Line Account Code
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("CreditLineAccount");
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "U_CrLnAcct");
                    oColumn.ChooseFromListUID = "CrLnAcctCFL";
                    oColumn.ChooseFromListAlias = "AcctCode";
                    oLink = oColumn.ExtendedObject;
                    oLink.LinkedObjectType = "1"; //G/L Accounts

                    oColumn = oColumns.Add("IntrstRate", SAPbouiCOM.BoFormItemTypes.it_EDIT); //Interest Rate
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("InterestRate");
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "U_IntrstRate");

                    oColumn = oColumns.Add("AccrDate", SAPbouiCOM.BoFormItemTypes.it_EDIT); //Accrual Date Start
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("AccrualDateStart");
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "U_AccrDate");

                    oColumn = oColumns.Add("AccrDays", SAPbouiCOM.BoFormItemTypes.it_EDIT); //Accrual Days
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("QuantityOfAccrualDays");
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "U_AccrDays");

                    oColumn = oColumns.Add("CrLnAmtLC", SAPbouiCOM.BoFormItemTypes.it_EDIT); //Credit Line Balance LC
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("CreditLineBalanceLC");
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "U_CrLnAmtLC");

                    oColumn = oColumns.Add("CrLnAmtFC", SAPbouiCOM.BoFormItemTypes.it_EDIT); //Credit Line Balance FC
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("CreditLineBalanceFC");
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "U_CrLnAmtFC");

                    oColumn = oColumns.Add("IntAmtLC", SAPbouiCOM.BoFormItemTypes.it_EDIT); //Interest Amount LC
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("InterestAmountLC");
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "U_IntAmtLC");

                    oColumn = oColumns.Add("IntAmtFC", SAPbouiCOM.BoFormItemTypes.it_EDIT); //Interest Amount FC
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("InterestAmountFC");
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "U_IntAmtFC");

                    FormsB1.addChooseFromList(oForm, false, "1", "ExpnsAcctCFL");
                    oColumn = oColumns.Add("ExpnsAcct", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON); //Expense Account Code
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("ExpenseAccount");
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "U_ExpnsAcct");
                    oColumn.ChooseFromListUID = "ExpnsAcctCFL";
                    oColumn.ChooseFromListAlias = "AcctCode";
                    oLink = oColumn.ExtendedObject;
                    oLink.LinkedObjectType = "1"; //G/L Accounts

                    FormsB1.addChooseFromList(oForm, false, "1", "IntPblAcctCFL");
                    oColumn = oColumns.Add("IntPblAcct", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON); //Interest Payable Account Code
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("InterestPayableAccount");
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "U_IntPblAcct");
                    oColumn.ChooseFromListUID = "IntPblAcctCFL";
                    oColumn.ChooseFromListAlias = "AcctCode";
                    oLink = oColumn.ExtendedObject;
                    oLink.LinkedObjectType = "1"; //G/L Accounts
                }
                oForm.Visible = true;
                oForm.Select();
            }
        }

        public static void addMenus()
        {
            try
            {
                SAPbouiCOM.MenuItem fatherMenuItem = Program.uiApp.Menus.Item("1536");
                // Add a pop-up menu item
                SAPbouiCOM.MenuCreationParams oCreationPackage = (SAPbouiCOM.MenuCreationParams)Program.uiApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                oCreationPackage.Checked = false;
                oCreationPackage.Enabled = true;
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "BDOSInterestAccrualWizard";
                oCreationPackage.String = BDOSResources.getTranslate("InterestAccrualWizard");
                oCreationPackage.Position = fatherMenuItem.SubMenus.Count - 1;

                SAPbouiCOM.MenuItem menuItem = fatherMenuItem.SubMenus.AddEx(oCreationPackage);
            }
            catch
            {
                //Program.uiApp.MessageBox(ex.Message);
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
                    int answer = Program.uiApp.MessageBox(BDOSResources.getTranslate("CloseInterestAccrualWizard") + "?", 1, BDOSResources.getTranslate("Yes"), BDOSResources.getTranslate("No"), "");

                    if (answer != 1)
                        BubbleEvent = false;
                }

                else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK)
                {
                    if (!pVal.BeforeAction)
                    {
                        if ((pVal.ItemUID == "checkB" || pVal.ItemUID == "unCheckB"))
                        {
                            checkUncheckMTR(oForm, pVal.ItemUID);
                        }
                        else if (pVal.ItemUID == "fillB")
                        {
                            fillMTR(oForm);
                        }
                        else if (pVal.ItemUID == "createDocB")
                        {
                            int answer = Program.uiApp.MessageBox(BDOSResources.getTranslate("CreateInterestAccrualDocument") + "?", 1, BDOSResources.getTranslate("Yes"), BDOSResources.getTranslate("No"), "");

                            if (answer == 1)
                            {
                                createDocuments(oForm);
                            }
                            return;
                        }
                    }
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
            }
        }

        public static void resizeForm(SAPbouiCOM.Form oForm)
        {
            oForm.Freeze(true);
            try
            {
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("LoanMTR").Specific;
                int mtrWidth = oForm.ClientWidth;
                oForm.Items.Item("LoanMTR").Width = mtrWidth;
                oForm.Items.Item("LoanMTR").Height = oForm.ClientHeight - 25;
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

        public static void checkUncheckMTR(SAPbouiCOM.Form oForm, string checkOperation)
        {
            try
            {
                oForm.Freeze(true);

                SAPbouiCOM.CheckBox oCheckBox;
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("LoanMTR").Specific;

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
                GC.Collect();
            }
        }

        public static void chooseFromList(SAPbouiCOM.Form oForm, SAPbouiCOM.ItemEvent pVal, SAPbouiCOM.IChooseFromListEvent oCFLEvento)
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
                        if (oCFLEvento.ChooseFromListUID == "BankCFL")
                        {
                            string value = oDataTable.GetValue("BankCode", 0);
                            LanguageUtils.IgnoreErrors<string>(() => oForm.Items.Item("BankCodeE").Specific.Value = value);
                            //if (bankCodeOld != value)
                            //    clearMatrix(oForm);
                        }
                        //else if (oCFLEvento.ChooseFromListUID.StartsWith("Dimension"))
                        //{
                        //    string dimension = oDataTable.GetValue("OcrCode", 0);
                        //    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("LoanMTR").Specific;
                        //    LanguageUtils.IgnoreErrors<string>(() => oMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value = dimension);
                        //}
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

        public static void fillMTR(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("LoanMTR");
            oDataTable.Rows.Clear();
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string bankCode = oForm.DataSources.UserDataSources.Item("BankCodeE").ValueEx;
            string docDateStr = oForm.DataSources.UserDataSources.Item("DocDateE").ValueEx;

            if (string.IsNullOrEmpty(bankCode) || string.IsNullOrEmpty(docDateStr))
            {
                string errorText = BDOSResources.getTranslate("TheFollowingFieldsAreMandatory") +
                    " : \"" + oForm.Items.Item("BankCodeS").Specific.caption + "\", " +
                    "\"" + oForm.Items.Item("DocDateS").Specific.caption + "\"";

                Program.uiApp.SetStatusBarMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short);
                return;
            }

            DateTime docDate = Convert.ToDateTime(DateTime.ParseExact(docDateStr, "yyyyMMdd", CultureInfo.InvariantCulture));
            SAPbobsCOM.SBObob oSBOBob = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);

            StringBuilder query = new StringBuilder();
            query.Append("SELECT T3.*, \n");
            query.Append("T4.\"U_DocDate\" \n");
            query.Append("FROM \n");
            query.Append("(SELECT \n");
            query.Append("T0.\"Code\", \n");
            query.Append("T0.\"Name\", \n");
            query.Append("T0.\"U_StartDate\", \n");
            query.Append("T0.\"U_BankCode\", \n");
            query.Append("T0.\"U_CurrCode\", \n");
            query.Append("T0.\"U_CrLnAcct\", \n");
            query.Append("T0.\"U_IntrstRate\", \n");
            query.Append("T0.\"U_ExpnsAcct\", \n");
            query.Append("T0.\"U_IntPblAcct\", \n");
            query.Append("T1.\"U_StartDate\" AS \"U_EndDate\" \n");
            query.Append("FROM \"@BDOSCRLN\" AS T0 \n");
            query.Append("LEFT JOIN \"@BDOSCRLN\" AS T1 \n");
            query.Append("ON T0.\"Name\" = T1.\"Name\" AND T0.\"U_StartDate\" < T1.\"U_StartDate\" \n");
            query.Append("WHERE T0.\"U_BankCode\" = '" + bankCode + "' AND T0.\"U_StartDate\" <= '" + docDateStr + "' \n");
            query.Append("ORDER BY T0.\"Name\", T0.\"U_StartDate\") AS T3 \n");
            query.Append("LEFT JOIN \n");
            query.Append("(SELECT \n");
            query.Append("MAX(\"U_DocDate\") AS \"U_DocDate\", \n");
            query.Append("\"U_CrLnCode\", \n");
            query.Append("\"U_CrLnName\" \n");
            query.Append("FROM \"@BDOSINA1\" \n");
            query.Append("INNER JOIN \"@BDOSINAC\" ON \"@BDOSINAC\".\"DocEntry\" = \"@BDOSINA1\".\"DocEntry\" \n");
            query.Append("WHERE \"Canceled\" = 'N' AND \"U_BankCode\" = '" + bankCode + "' \n");
            query.Append("GROUP BY \n");
            query.Append("\"U_CrLnCode\", \n");
            query.Append("\"U_CrLnName\", \n");
            query.Append("\"U_BankCode\") AS T4 \n");
            query.Append("ON T3.\"Code\" = \"T4\".\"U_CrLnCode\" \n");
            query.Append("ORDER BY T3.\"Name\", T3.\"U_StartDate\"");

            oRecordSet.DoQuery(query.ToString());

            try
            {
                int rowIndex = 0;

                while (!oRecordSet.EoF)
                {
                    DateTime creditLineStartDate = oRecordSet.Fields.Item("U_StartDate").Value;
                    DateTime lastInterestAccrualDocDate = oRecordSet.Fields.Item("U_DocDate").Value;
                    DateTime creditLineEndDate = oRecordSet.Fields.Item("U_EndDate").Value;

                    if (lastInterestAccrualDocDate >= docDate)
                    {
                        oRecordSet.MoveNext();
                        continue;
                    }

                    string currency = oRecordSet.Fields.Item("U_CurrCode").Value;
                    decimal interestRate = Convert.ToDecimal(oRecordSet.Fields.Item("U_IntrstRate").Value, CultureInfo.InvariantCulture);
                    bool foreignCurrency = (!string.IsNullOrEmpty(currency) && currency != Program.LocalCurrency);
                    decimal currencyRate = foreignCurrency ? Convert.ToDecimal(oSBOBob.GetCurrencyRate(currency, docDate).Fields.Item("CurrencyRate").Value, CultureInfo.InvariantCulture) : decimal.Zero;

                    DateTime accrualStartDate;
                    DateTime accrualEndDate;
                    int accrualDays;

                    if (creditLineEndDate.ToString("yyyyMMdd") != "18991230" && creditLineEndDate <= docDate)
                        accrualEndDate = creditLineEndDate;
                    else
                        accrualEndDate = docDate.AddDays(1);

                    if (lastInterestAccrualDocDate.ToString("yyyyMMdd") != "18991230")
                        accrualStartDate = lastInterestAccrualDocDate.AddDays(1);
                    else
                        accrualStartDate = creditLineStartDate;

                    int numberOfDaysInYear = new DateTime(DateTime.Today.Year, 12, 31).DayOfYear;
                    accrualDays = (accrualEndDate - accrualStartDate).Days;
                    accrualDays = accrualDays == 0 ? 1 : accrualDays;

                    if (accrualDays <= 0)
                    {
                        oRecordSet.MoveNext();
                        continue;
                    }

                    decimal dayRate = interestRate / numberOfDaysInYear;
                    string creditLineAccountCode = oRecordSet.Fields.Item("U_CrLnAcct").Value;
                    string currencyForQuery = foreignCurrency ? " = '" + currency + "' " : " IS NULL ";

                    decimal creditLineBalanceFC = decimal.Zero;
                    decimal creditLineBalanceLC = decimal.Zero;
                    decimal interestAmountFC = decimal.Zero;
                    decimal interestAmountLC = decimal.Zero;

                    for (int k = 0; k < accrualDays; k++)
                    {
                        StringBuilder queryForBalances2 = new StringBuilder();
                        queryForBalances2.Append("SELECT \n");
                        queryForBalances2.Append("Sum(\"Credit\") - Sum(\"Debit\") AS \"U_CrLnAmtLC\", \n");
                        queryForBalances2.Append("Sum(\"FCCredit\") - Sum(\"FCDebit\") AS \"U_CrLnAmtFC\" \n");
                        queryForBalances2.Append("FROM \"JDT1\" \n");
                        queryForBalances2.Append("WHERE \"Account\" = '" + creditLineAccountCode + "' \n");
                        queryForBalances2.Append("AND \"FCCurrency\" " + currencyForQuery + " \n");
                        queryForBalances2.Append("AND \"RefDate\" <= '" + accrualStartDate.AddDays(k).ToString("yyyyMMdd") + "' \n");
                        queryForBalances2.Append("GROUP BY \n");
                        queryForBalances2.Append("\"Account\", \n");
                        queryForBalances2.Append("\"FCCurrency\"");

                        SAPbobsCOM.Recordset oRecordSetForBalances2 = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        oRecordSetForBalances2.DoQuery(queryForBalances2.ToString());

                        if (!oRecordSetForBalances2.EoF)
                        {
                            creditLineBalanceFC = foreignCurrency ? Convert.ToDecimal(oRecordSetForBalances2.Fields.Item("U_CrLnAmtFC").Value, CultureInfo.InvariantCulture) : decimal.Zero;
                            creditLineBalanceLC = foreignCurrency ? creditLineBalanceFC * currencyRate : Convert.ToDecimal(oRecordSetForBalances2.Fields.Item("U_CrLnAmtLC").Value, CultureInfo.InvariantCulture);

                            interestAmountFC += creditLineBalanceFC * dayRate;
                            interestAmountLC += creditLineBalanceLC * dayRate;

                            oRecordSetForBalances2.MoveNext();
                        }
                        Marshal.ReleaseComObject(oRecordSetForBalances2);
                    }

                    oDataTable.Rows.Add();
                    oDataTable.SetValue("LineNum", rowIndex, rowIndex + 1);
                    oDataTable.SetValue("CheckBox", rowIndex, "N");
                    oDataTable.SetValue("U_CrLnCode", rowIndex, oRecordSet.Fields.Item("Code").Value);
                    oDataTable.SetValue("U_CrLnName", rowIndex, oRecordSet.Fields.Item("Name").Value);
                    oDataTable.SetValue("U_StartDate", rowIndex, creditLineStartDate);
                    oDataTable.SetValue("U_LnCurrCode", rowIndex, currency);
                    oDataTable.SetValue("U_ExchngRate", rowIndex, Convert.ToDouble(currencyRate, CultureInfo.InvariantCulture));
                    oDataTable.SetValue("U_CrLnAcct", rowIndex, creditLineAccountCode);
                    oDataTable.SetValue("U_IntrstRate", rowIndex, oRecordSet.Fields.Item("U_IntrstRate").Value);
                    oDataTable.SetValue("U_AccrDate", rowIndex, accrualStartDate);
                    oDataTable.SetValue("U_AccrDays", rowIndex, accrualDays);
                    oDataTable.SetValue("U_CrLnAmtLC", rowIndex, Convert.ToDouble(creditLineBalanceLC, CultureInfo.InvariantCulture));
                    oDataTable.SetValue("U_CrLnAmtFC", rowIndex, Convert.ToDouble(creditLineBalanceFC, CultureInfo.InvariantCulture));
                    oDataTable.SetValue("U_IntAmtLC", rowIndex, Convert.ToDouble(interestAmountLC, CultureInfo.InvariantCulture));
                    oDataTable.SetValue("U_IntAmtFC", rowIndex, Convert.ToDouble(interestAmountFC, CultureInfo.InvariantCulture));
                    oDataTable.SetValue("U_ExpnsAcct", rowIndex, oRecordSet.Fields.Item("U_ExpnsAcct").Value);
                    oDataTable.SetValue("U_IntPblAcct", rowIndex, oRecordSet.Fields.Item("U_IntPblAcct").Value);
                    rowIndex++;

                    oRecordSet.MoveNext();
                }

                oForm.Freeze(true);
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("LoanMTR").Specific;
                oMatrix.Clear();
                oMatrix.LoadFromDataSource();
                oForm.Update();
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

        public static void createDocuments(SAPbouiCOM.Form oForm)
        {
            string errorText;

            string bankCode = oForm.DataSources.UserDataSources.Item("BankCodeE").ValueEx;
            string docDateStr = oForm.DataSources.UserDataSources.Item("DocDateE").ValueEx;

            if (string.IsNullOrEmpty(bankCode) || string.IsNullOrEmpty(docDateStr))
            {
                errorText = BDOSResources.getTranslate("TheFollowingFieldsAreMandatory") +
                    " : \"" + oForm.Items.Item("BankCodeS").Specific.caption + "\"" +
                    ", \"" + oForm.Items.Item("DocDateS").Specific.caption + "\"";

                Program.uiApp.SetStatusBarMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short);
                return;
            }

            DateTime docDate = Convert.ToDateTime(DateTime.ParseExact(docDateStr, "yyyyMMdd", CultureInfo.InvariantCulture));

            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("LoanMTR").Specific;
            oMatrix.FlushToDataSource();

            SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("LoanMTR");
            string checkBox;

            SAPbobsCOM.CompanyService oCompanyService = Program.oCompany.GetCompanyService();
            SAPbobsCOM.GeneralService oGeneralService = oCompanyService.GetGeneralService("UDO_F_BDOSINAC_D");
            SAPbobsCOM.GeneralData oGeneralData = (SAPbobsCOM.GeneralData)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);
            SAPbobsCOM.GeneralDataCollection oChildren = oGeneralData.Child("BDOSINA1");

            string prevCrLnName = null;
            string crLnName = null;
            int createdDocEntry;
            List<int> tableRowList = new List<int>();
            int transId;

            for (int i = 0; i < oDataTable.Rows.Count; i++)
            {
                checkBox = oDataTable.GetValue("CheckBox", i);

                if (checkBox == "Y" && oDataTable.GetValue("DocEntry", i) == 0)
                {
                    crLnName = oDataTable.GetValue("U_CrLnName", i);
                    if (prevCrLnName != null && prevCrLnName != crLnName)
                    {
                        oGeneralData.SetProperty("U_DocDate", docDate);
                        oGeneralData.SetProperty("U_BankCode", bankCode);

                        createdDocEntry = createInterestAccrualDocument(oGeneralService, oGeneralData, docDate, prevCrLnName, out transId);
                        if (createdDocEntry > 0 && transId > 0)
                            for (int j = 0; j < tableRowList.Count; j++)
                                oDataTable.SetValue("DocEntry", tableRowList[j], createdDocEntry);

                        tableRowList.Clear();
                        Marshal.ReleaseComObject(oChildren);
                        Marshal.ReleaseComObject(oGeneralData);
                        Marshal.ReleaseComObject(oGeneralService);
                        Marshal.ReleaseComObject(oCompanyService);

                        oCompanyService = Program.oCompany.GetCompanyService();
                        oGeneralService = oCompanyService.GetGeneralService("UDO_F_BDOSINAC_D");
                        oGeneralData = (SAPbobsCOM.GeneralData)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);
                        oChildren = oGeneralData.Child("BDOSINA1");
                    }
                    SAPbobsCOM.GeneralData oChild = oChildren.Add();
                    oChild.SetProperty("U_CrLnCode", oDataTable.GetValue("U_CrLnCode", i));
                    oChild.SetProperty("U_CrLnName", oDataTable.GetValue("U_CrLnName", i));
                    oChild.SetProperty("U_LnCurrCode", oDataTable.GetValue("U_LnCurrCode", i));
                    oChild.SetProperty("U_ExchngRate", oDataTable.GetValue("U_ExchngRate", i));
                    oChild.SetProperty("U_CrLnAcct", oDataTable.GetValue("U_CrLnAcct", i));
                    oChild.SetProperty("U_IntrstRate", oDataTable.GetValue("U_IntrstRate", i));
                    oChild.SetProperty("U_CrLnAmtLC", oDataTable.GetValue("U_CrLnAmtLC", i));
                    oChild.SetProperty("U_CrLnAmtFC", oDataTable.GetValue("U_CrLnAmtFC", i));
                    oChild.SetProperty("U_IntAmtLC", oDataTable.GetValue("U_IntAmtLC", i));
                    oChild.SetProperty("U_IntAmtFC", oDataTable.GetValue("U_IntAmtFC", i));
                    oChild.SetProperty("U_ExpnsAcct", oDataTable.GetValue("U_ExpnsAcct", i));
                    oChild.SetProperty("U_IntPblAcct", oDataTable.GetValue("U_IntPblAcct", i));

                    prevCrLnName = crLnName;
                    tableRowList.Add(i);
                }
            }
            oGeneralData.SetProperty("U_DocDate", docDate);
            oGeneralData.SetProperty("U_BankCode", bankCode);

            createdDocEntry = createInterestAccrualDocument(oGeneralService, oGeneralData, docDate, crLnName, out transId);
            if (createdDocEntry > 0 && transId > 0)
                for (int j = 0; j < tableRowList.Count; j++)
                    oDataTable.SetValue("DocEntry", tableRowList[j], createdDocEntry);

            Marshal.ReleaseComObject(oChildren);
            Marshal.ReleaseComObject(oGeneralData);
            Marshal.ReleaseComObject(oGeneralService);
            Marshal.ReleaseComObject(oCompanyService);

            oForm.Freeze(true);
            oMatrix.Clear();
            oMatrix.LoadFromDataSource();
            oForm.Update();
            oForm.Freeze(false);
        }

        static int createInterestAccrualDocument(SAPbobsCOM.GeneralService oGeneralService, SAPbobsCOM.GeneralData oGeneralData, DateTime docDate, string crLnName, out int transId)
        {
            int docEntry = 0;
            string errorText = null;
            transId = 0;

            try
            {
                CommonFunctions.StartTransaction();

                var response = oGeneralService.Add(oGeneralData);
                docEntry = response.GetProperty("DocEntry");

                if (docEntry > 0)
                {
                    DataTable jrnLinesDT = BDOSInterestAccrual.createAdditionalEntries(null, oGeneralData);
                    transId = BDOSInterestAccrual.jrnEntry(docEntry.ToString(), docEntry.ToString(), docDate, jrnLinesDT, out errorText);

                    if (errorText != null || transId == 0)
                    {
                        if (Program.oCompany.InTransaction)
                            CommonFunctions.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                    }
                    else
                    {
                        Marshal.ReleaseComObject(oGeneralData);

                        oGeneralData = (SAPbobsCOM.GeneralData)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);
                        //Get UDO record
                        SAPbobsCOM.GeneralDataParams oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                        oGeneralParams.SetProperty("DocEntry", docEntry);
                        oGeneralData = oGeneralService.GetByParams(oGeneralParams);
                        oGeneralData.SetProperty("U_TransId", transId);
                        oGeneralService.Update(oGeneralData);

                        Marshal.ReleaseComObject(oGeneralParams);

                        if (Program.oCompany.InTransaction)
                            CommonFunctions.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                    }
                }
                else
                {
                    if (Program.oCompany.InTransaction)
                        CommonFunctions.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                    int resultCode;
                    string errorMessage;
                    Program.oCompany.GetLastError(out resultCode, out errorMessage);
                    errorText = errorMessage;
                }

                return docEntry;
            }
            catch (Exception Ex)
            {
                if (Program.oCompany.InTransaction)
                    CommonFunctions.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                errorText = Ex.Message;
                return docEntry;
            }
            finally
            {
                if (docEntry > 0 && transId > 0)
                    Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("DocumentCreatedSuccesfully") + ": " + docEntry + ". " + BDOSResources.getTranslate("CreditLineCode") + ": \"" + crLnName + "\"", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                else
                    Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("DocumentNotCreated") + "! " + BDOSResources.getTranslate("CreditLineCode") + ": \"" + crLnName + "\". " + BDOSResources.getTranslate("ReasonIs") + ": " + errorText, SAPbouiCOM.BoMessageTime.bmt_Short);
            }
        }
    }
}
