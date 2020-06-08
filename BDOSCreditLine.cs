using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace BDO_Localisation_AddOn
{
    class BDOSCreditLine
    {
        public static void createMasterDataUDO(out string errorText)
        {
            string tableName = "BDOSCRLN";
            string description = "Credit Line Master Data";

            int result = UDO.addUserTable(tableName, description, SAPbobsCOM.BoUTBTableType.bott_MasterData, out errorText);

            if (result != 0)
            {
                return;
            }

            Dictionary<string, object> fieldskeysMap;

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "CurrCode");
            fieldskeysMap.Add("TableName", "BDOSCRLN");
            fieldskeysMap.Add("Description", "Currency Code");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 3);
            fieldskeysMap.Add("Mandatory", SAPbobsCOM.BoYesNoEnum.tYES);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BankCode");
            fieldskeysMap.Add("TableName", "BDOSCRLN");
            fieldskeysMap.Add("Description", "Bank Code");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 30);
            fieldskeysMap.Add("Mandatory", SAPbobsCOM.BoYesNoEnum.tYES);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "CrLnAcct");
            fieldskeysMap.Add("TableName", "BDOSCRLN");
            fieldskeysMap.Add("Description", "Credit Line Account Code");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 15);
            fieldskeysMap.Add("Mandatory", SAPbobsCOM.BoYesNoEnum.tYES);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "IntrstRate");
            fieldskeysMap.Add("TableName", "BDOSCRLN");
            fieldskeysMap.Add("Description", "Interest Rate");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Percentage);
            fieldskeysMap.Add("Mandatory", SAPbobsCOM.BoYesNoEnum.tYES);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "StartDate");
            fieldskeysMap.Add("TableName", "BDOSCRLN");
            fieldskeysMap.Add("Description", "Starting Date");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Date);
            fieldskeysMap.Add("Mandatory", SAPbobsCOM.BoYesNoEnum.tYES);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "ExpnsAcct");
            fieldskeysMap.Add("TableName", "BDOSCRLN");
            fieldskeysMap.Add("Description", "Expense Account Code");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 15);
            fieldskeysMap.Add("Mandatory", SAPbobsCOM.BoYesNoEnum.tYES);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "IntPblAcct");
            fieldskeysMap.Add("TableName", "BDOSCRLN");
            fieldskeysMap.Add("Description", "Interest Payable Account Code");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 15);
            fieldskeysMap.Add("Mandatory", SAPbobsCOM.BoYesNoEnum.tYES);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            Dictionary<string, string> validValues = new Dictionary<string, string>();
            validValues.Add("C", "Calendar Year");
            validValues.Add("F", "Fixed");

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "Type");
            fieldskeysMap.Add("TableName", "BDOSCRLN");
            fieldskeysMap.Add("Description", "Type");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 1);
            fieldskeysMap.Add("DefaultValue", "C");
            fieldskeysMap.Add("ValidValues", validValues); 

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "NbrOfDays");
            fieldskeysMap.Add("TableName", "BDOSCRLN");
            fieldskeysMap.Add("Description", "Number of Days");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Numeric);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Quantity);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            GC.Collect();
        }

        public static void registerUDO()
        {
            SAPbobsCOM.UserObjectsMD oUserObjectMD = null;
            SAPbobsCOM.UserObjectMD_FindColumns oUDOFind = null;
            SAPbobsCOM.UserObjectMD_FormColumns oUDOForm = null;
            SAPbobsCOM.UserObjectMD_EnhancedFormColumns oUDOEnhancedForm = null;
            oUserObjectMD = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);
            Marshal.ReleaseComObject(oUserObjectMD);
            oUserObjectMD = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);
            oUDOFind = oUserObjectMD.FindColumns;
            oUDOForm = oUserObjectMD.FormColumns;
            oUDOEnhancedForm = oUserObjectMD.EnhancedFormColumns;

            var retval = oUserObjectMD.GetByKey("UDO_F_BDOSCRLN_D");

            if (!retval)
            {
                oUserObjectMD.Code = "UDO_F_BDOSCRLN_D";
                oUserObjectMD.Name = "Credit Line Master Data";
                oUserObjectMD.TableName = "BDOSCRLN";
                oUserObjectMD.ObjectType = SAPbobsCOM.BoUDOObjType.boud_MasterData;
                oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tNO;
                oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tNO;
                oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO;
                oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tNO;
                oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tNO;
                oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.EnableEnhancedForm = SAPbobsCOM.BoYesNoEnum.tNO;

                //Find
                oUDOFind.ColumnAlias = "Code";
                oUDOFind.ColumnDescription = "Code";
                oUDOFind.Add();
                oUDOFind.ColumnAlias = "Name";
                oUDOFind.ColumnDescription = "Credit Line Code";
                oUDOFind.Add();
                oUDOFind.ColumnAlias = "U_CurrCode";
                oUDOFind.ColumnDescription = "Currency Code";
                oUDOFind.Add();
                oUDOFind.ColumnAlias = "U_BankCode";
                oUDOFind.ColumnDescription = "Bank Code";
                oUDOFind.Add();
                oUDOFind.ColumnAlias = "U_CrLnAcct";
                oUDOFind.ColumnDescription = "Credit Line Account Code";
                oUDOFind.Add();
                oUDOFind.ColumnAlias = "U_IntrstRate";
                oUDOFind.ColumnDescription = "Interest Rate";
                oUDOFind.Add();
                oUDOFind.ColumnAlias = "U_StartDate";
                oUDOFind.ColumnDescription = "Starting Date";
                oUDOFind.Add();
                oUDOFind.ColumnAlias = "U_ExpnsAcct";
                oUDOFind.ColumnDescription = "Expense Account Code";
                oUDOFind.Add();
                oUDOFind.ColumnAlias = "U_IntPblAcct";
                oUDOFind.ColumnDescription = "Interest Payable Account Code";
                oUDOFind.Add();
                oUDOFind.ColumnAlias = "U_Type";
                oUDOFind.ColumnDescription = "Type";
                oUDOFind.Add();
                oUDOFind.ColumnAlias = "U_NbrOfDays";
                oUDOFind.ColumnDescription = "Number of Days";
                oUDOFind.Add();

                for (int i = 0; i < oUDOFind.Count - 1; i++)
                {
                    oUDOFind.SetCurrentLine(i);

                    //Form
                    oUDOForm.FormColumnAlias = oUDOFind.ColumnAlias;
                    oUDOForm.FormColumnDescription = oUDOFind.ColumnDescription;
                    oUDOForm.Editable = SAPbobsCOM.BoYesNoEnum.tYES;
                    oUDOForm.Add();

                    //Enhanced
                    oUDOEnhancedForm.ColumnAlias = oUDOFind.ColumnAlias;
                    oUDOEnhancedForm.ColumnDescription = oUDOFind.ColumnDescription;
                    oUDOEnhancedForm.ColumnIsUsed = SAPbobsCOM.BoYesNoEnum.tYES;
                    oUDOEnhancedForm.ColumnNumber = i + 1;
                    oUDOEnhancedForm.ChildNumber = 1;
                    oUDOEnhancedForm.Editable = SAPbobsCOM.BoYesNoEnum.tYES;
                    oUDOEnhancedForm.Add();
                }

                if (oUserObjectMD.Add() != 0)
                {
                    Program.uiApp.MessageBox(Program.oCompany.GetLastErrorDescription());
                }
            }
            Marshal.ReleaseComObject(oUserObjectMD);
        }

        public static void updateUDO()
        {
            SAPbobsCOM.UserObjectsMD oUserObjectMD = null;
            SAPbobsCOM.UserObjectMD_FindColumns oUDOFind = null;
            SAPbobsCOM.UserObjectMD_FormColumns oUDOForm = null;
            SAPbobsCOM.UserObjectMD_EnhancedFormColumns oUDOEnhancedForm = null;
            oUserObjectMD = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);
            Marshal.ReleaseComObject(oUserObjectMD);
            oUserObjectMD = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);

            var retval = oUserObjectMD.GetByKey("UDO_F_BDOSCRLN_D");

            if (retval)
            {
                oUDOFind = oUserObjectMD.FindColumns;
                oUDOForm = oUserObjectMD.FormColumns;
                oUDOEnhancedForm = oUserObjectMD.EnhancedFormColumns;

                if (oUDOForm.Count < 11)
                {
                    int oldCount = oUDOFind.Count - 1;

                    oUDOFind.SetCurrentLine(oldCount);
                    oUDOForm.SetCurrentLine(oldCount);

                    //Find
                    oUDOFind.Add();
                    oUDOFind.ColumnAlias = "U_Type";
                    oUDOFind.ColumnDescription = "Type";
                    oUDOFind.Add();
                    oUDOFind.ColumnAlias = "U_NbrOfDays";
                    oUDOFind.ColumnDescription = "Number of Days";
                    oUDOFind.Add();

                    //Form
                    oUDOForm.Add();
                    oUDOForm.FormColumnAlias = "U_Type";
                    oUDOForm.FormColumnDescription = "Type";
                    oUDOForm.Editable = SAPbobsCOM.BoYesNoEnum.tYES;
                    oUDOForm.Add();
                    oUDOForm.FormColumnAlias = "U_NbrOfDays";
                    oUDOForm.FormColumnDescription = "Number of Days";
                    oUDOForm.Editable = SAPbobsCOM.BoYesNoEnum.tYES;
                    oUDOForm.Add();

                    if (oUserObjectMD.Update() != 0)
                    {
                        Program.uiApp.MessageBox(Program.oCompany.GetLastErrorDescription());
                    }
                }
            }
            Marshal.ReleaseComObject(oUserObjectMD);
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
                oCreationPackage.UniqueID = "UDO_F_BDOSCRLN_D";
                oCreationPackage.String = BDOSResources.getTranslate("CreditLineMasterData");
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

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
                {
                    if (pVal.BeforeAction)
                    {
                        Program.FORM_LOAD_FOR_ACTIVATE = true;
                        changeFormItems(oForm);
                        oForm.Title = BDOSResources.getTranslate("CreditLineMasterData");
                    }
                }

                else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE)
                {
                    if (!pVal.BeforeAction)
                    {
                        if (Program.FORM_LOAD_FOR_ACTIVATE)
                        {
                            Program.FORM_LOAD_FOR_ACTIVATE = false;
                            setVisibleFormItems(oForm);
                        }
                    }
                }

                else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                {
                    SAPbouiCOM.IChooseFromListEvent oCFLEvento = (SAPbouiCOM.IChooseFromListEvent)pVal;
                    chooseFromList(oForm, pVal, oCFLEvento);
                }
            }
        }

        public static void uiApp_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;

            SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(BusinessObjectInfo.FormTypeEx, Program.currentFormCount);

            if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE)
            {
                if (BusinessObjectInfo.BeforeAction)
                {
                    if (checkRemoving(oForm))
                    {
                        Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("RecordIsUsedInDocuments"));
                        BubbleEvent = false;
                    }
                }
            }
        }

        public static void changeFormItems(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("3").Specific;

            SAPbouiCOM.Columns oColumns = oMatrix.Columns;

            SAPbouiCOM.Column oColumn = oColumns.Item("Code");
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("CreditLine");

            oColumn = oColumns.Item("Name");
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("CreditLineCode");

            FormsB1.addChooseFromList(oForm, false, "37", "CurrencyCFL");
            oColumn = oColumns.Item("U_CurrCode");
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Currency");
            oColumn.ChooseFromListUID = "CurrencyCFL";
            oColumn.ChooseFromListAlias = "CurrCode";

            FormsB1.addChooseFromList(oForm, false, "3", "BankCFL");
            oColumn = oColumns.Item("U_BankCode");
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Bank");
            oColumn.ChooseFromListUID = "BankCFL";
            oColumn.ChooseFromListAlias = "BankCode";

            FormsB1.addChooseFromList(oForm, false, "1", "CrLnAcctCFL");
            oColumn = oColumns.Item("U_CrLnAcct");
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("CreditLineAccount");
            oColumn.ChooseFromListUID = "CrLnAcctCFL";
            oColumn.ChooseFromListAlias = "AcctCode";

            oColumn = oColumns.Item("U_IntrstRate");
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("InterestRate");

            oColumn = oColumns.Item("U_StartDate");
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("StartingDate");

            FormsB1.addChooseFromList(oForm, false, "1", "ExpnsAcctCFL");
            oColumn = oColumns.Item("U_ExpnsAcct");
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("ExpenseAccount");
            oColumn.ChooseFromListUID = "ExpnsAcctCFL";
            oColumn.ChooseFromListAlias = "AcctCode";

            FormsB1.addChooseFromList(oForm, false, "1", "IntPblAcctCFL");
            oColumn = oColumns.Item("U_IntPblAcct");
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("InterestPayableAccount");
            oColumn.ChooseFromListUID = "IntPblAcctCFL";
            oColumn.ChooseFromListAlias = "AcctCode";

            oColumn = oColumns.Item("U_Type");
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Type");
            oColumn.DisplayDesc = true;
            oColumn.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;

            oColumn = oColumns.Item("U_NbrOfDays");
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("NumberOfDays");
        }

        public static void setVisibleFormItems(SAPbouiCOM.Form oForm)
        {
            try
            {
                //SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("3").Specific));

                //SAPbouiCOM.Columns oColumns = oMatrix.Columns;
                //SAPbouiCOM.Column oColumn = oColumns.Item("Code");
                //oColumn.Editable = string.IsNullOrEmpty(oMatrix.Columns.Item("Code").Cells.Item(1).Specific.Value);

                //oMatrix.Columns.Item("Code").Cells.Item(1).Specific.Item.Enabled = false;


                //oForm.Items.Item("0_U_E").Enabled = (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE);
            }

            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }

            finally
            {
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
                    if (oCFLEvento.ChooseFromListUID == "CrLnAcctCFL" || oCFLEvento.ChooseFromListUID == "ExpnsAcctCFL" || oCFLEvento.ChooseFromListUID == "IntPblAcctCFL")
                    {
                        SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("3").Specific;
                        string currency = oMatrix.Columns.Item("U_CurrCode").Cells.Item(pVal.Row).Specific.Value;

                        SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID);
                        SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();

                        SAPbouiCOM.Condition oCon = oCons.Add();
                        oCon.Alias = "Postable"; //Active Account, (Title Account)
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = "Y";

                        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                        oCon = oCons.Add();
                        oCon.Alias = "FrozenFor"; //Inactive
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = "N";

                        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                        oCon = oCons.Add();
                        oCon.Alias = "ActCurr"; //Account Currency
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = string.IsNullOrEmpty(currency) ? "000" : currency;
                        oCon.BracketOpenNum = 1;

                        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR;

                        oCon = oCons.Add();
                        oCon.Alias = "ActCurr"; //Account Currency
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = string.IsNullOrEmpty(currency) ? "000" : "##";
                        oCon.BracketCloseNum = 1;

                        oCFL.SetConditions(oCons);
                    }
                }
                else
                {
                    SAPbouiCOM.DataTable oDataTable = oCFLEvento.SelectedObjects;

                    if (oDataTable != null)
                    {
                        string value;
                        switch (oCFLEvento.ChooseFromListUID)
                        {
                            case "CurrencyCFL":
                                value = oDataTable.GetValue("CurrCode", 0);
                                break;
                            case "BankCFL":
                                value = oDataTable.GetValue("BankCode", 0);
                                break;
                            case "CrLnAcctCFL":
                            case "ExpnsAcctCFL":
                            case "IntPblAcctCFL":
                                value = oDataTable.GetValue("AcctCode", 0);
                                break;
                            default:
                                value = string.Empty;
                                break;
                        }

                        SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("3").Specific;
                        LanguageUtils.IgnoreErrors<string>(() => oMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value = value);
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

        public static bool checkRemoving(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.DBDataSource oDBDataSource = oForm.DataSources.DBDataSources.Item("@BDOSCRLN");
            string code = oDBDataSource.GetValue("Code", 0).Trim();

            Dictionary<string, string> listTables = new Dictionary<string, string>();
            listTables.Add("@BDOSINA1", "U_CrLnCode");

            return CommonFunctions.codeIsUsed(listTables, code);
        }

        public static SAPbobsCOM.Recordset getInfo(string code)
        {
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (string.IsNullOrEmpty(code))
                    return null;
                string query = @"SELECT
	               ""Code"",
                   ""Name"",
                   ""U_BankCode"",
	               ""U_CurrCode"",
                   ""U_CrLnAcct"",
	               ""U_IntrstRate"",
                   ""U_StartDate"",
	               ""U_ExpnsAcct"",
	               ""U_IntPblAcct""
                FROM ""@BDOSCRLN""
                WHERE ""Code"" = '" + code + "'";

                oRecordSet.DoQuery(query);

                if (!oRecordSet.EoF)
                {
                    return oRecordSet;
                }
                return null;
            }
            catch (Exception ex)
            {
                Marshal.ReleaseComObject(oRecordSet);
                throw new Exception(ex.Message);
            }
        }
    }
}
