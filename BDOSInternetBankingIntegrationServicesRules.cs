using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace BDO_Localisation_AddOn
{
    static partial class BDOSInternetBankingIntegrationServicesRules
    {
        public static void createMasterDataUDO( out string errorText)
        {
            errorText = null;
            string tableName = "BDOSINTR";
            string description = "Banking Integration Rules";

            SAPbobsCOM.UserObjectsMD oUserObjectsMD = null;
            oUserObjectsMD = (SAPbobsCOM.UserObjectsMD)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectsMD);

            int result = UDO.addUserTable( tableName, description, SAPbobsCOM.BoUTBTableType.bott_MasterData, out errorText);

            if (result != 0)
            {
                return;
            }

            Dictionary<string, object> fieldskeysMap;
            Dictionary<string, string> listValidValuesDict;

            listValidValuesDict = new Dictionary<string, string>();

            listValidValuesDict.Add("TransferToOwnAccount", "Transfer Between Own Accounts"); //!Transfer Between Own Accounts
            listValidValuesDict.Add("CurrencyExchange", "Currency Exchange"); //!Currency exchange
            listValidValuesDict.Add("TransferFromBP", "Transfer from BP"); //Transfer From BP
            listValidValuesDict.Add("TransferToBP", "Transfer to BP"); //Transfer To BP
            listValidValuesDict.Add("TreasuryTransfer", "Treasury Transfers"); //!Treasury transfers
            listValidValuesDict.Add("BankCharge", "BankCharge"); //Bank Charge
            listValidValuesDict.Add("OtherIncomes", "Other Incomes"); //Other expenses
            listValidValuesDict.Add("OtherExpenses", "Other Expenses"); //Other expenses
            listValidValuesDict.Add("Salary", "Salary"); //Other expenses

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "OpType");
            fieldskeysMap.Add("TableName", "BDOSINTR");
            fieldskeysMap.Add("Description", "Operation Type");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);
            fieldskeysMap.Add("ValidValues", listValidValuesDict);
            fieldskeysMap.Add("Mandatory", SAPbobsCOM.BoYesNoEnum.tYES);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "AcctCode");
            fieldskeysMap.Add("TableName", "BDOSINTR");
            fieldskeysMap.Add("Description", "Account Code");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 15);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "AcctName");
            fieldskeysMap.Add("TableName", "BDOSINTR");
            fieldskeysMap.Add("Description", "Account Name");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 100);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "CFWId");
            fieldskeysMap.Add("TableName", "BDOSINTR");
            fieldskeysMap.Add("Description", "Cash Flow Line Item ID");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 11);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "CFWName");
            fieldskeysMap.Add("TableName", "BDOSINTR");
            fieldskeysMap.Add("Description", "Cash Flow Line Item Name");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 100);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            if (CommonFunctions.IsDevelopment())
            {
                fieldskeysMap = new Dictionary<string, object>();
                fieldskeysMap.Add("Name", "BCFWId");
                fieldskeysMap.Add("TableName", "BDOSINTR");
                fieldskeysMap.Add("Description", "Budget Cash Flow ID");
                fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
                fieldskeysMap.Add("EditSize", 11);

                UDO.addUserTableFields(fieldskeysMap, out errorText);

                fieldskeysMap = new Dictionary<string, object>();
                fieldskeysMap.Add("Name", "BCFWName");
                fieldskeysMap.Add("TableName", "BDOSINTR");
                fieldskeysMap.Add("Description", "Budget Cash Flow Name");
                fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
                fieldskeysMap.Add("EditSize", 100);

                UDO.addUserTableFields(fieldskeysMap, out errorText);
            }

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "TresrCode");
            fieldskeysMap.Add("TableName", "BDOSINTR");
            fieldskeysMap.Add("Description", "Treasury Code");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            List<string> oColumnAlias = new List<string>();
            oColumnAlias.Add("OpType");
            oColumnAlias.Add("TresrCode");
            UDO.AddUserKey( "BDOSINTR", "OpTypeKey", oColumnAlias, out errorText);

            GC.Collect();
        }

        public static void registerUDO( out string errorText)
        {
            errorText = null;
            string code = "UDO_F_BDOSINTR_D"; //20 characters (must include at least one alphabetical character).
            Dictionary<string, object> formProperties;

            formProperties = new Dictionary<string, object>();
            formProperties.Add("Name", "Internet Banking Integration Services Rules"); //100 characters
            formProperties.Add("TableName", "BDOSINTR");
            formProperties.Add("ObjectType", SAPbobsCOM.BoUDOObjType.boud_MasterData);
            formProperties.Add("CanCancel", SAPbobsCOM.BoYesNoEnum.tYES);
            formProperties.Add("CanClose", SAPbobsCOM.BoYesNoEnum.tYES);
            formProperties.Add("CanCreateDefaultForm", SAPbobsCOM.BoYesNoEnum.tYES);
            formProperties.Add("CanDelete", SAPbobsCOM.BoYesNoEnum.tYES);
            formProperties.Add("CanFind", SAPbobsCOM.BoYesNoEnum.tYES);
            formProperties.Add("CanYearTransfer", SAPbobsCOM.BoYesNoEnum.tYES);
            formProperties.Add("EnableEnhancedForm", SAPbobsCOM.BoYesNoEnum.tNO);

            List<Dictionary<string, object>> listFindColumns = new List<Dictionary<string, object>>();
            List<Dictionary<string, object>> listFormColumns = new List<Dictionary<string, object>>();

            Dictionary<string, object> fieldskeysMap;

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "Code");
            fieldskeysMap.Add("ColumnDescription", "Code"); //30 characters
            listFindColumns.Add(fieldskeysMap);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "Name");
            fieldskeysMap.Add("ColumnDescription", "Name"); //30 characters
            listFindColumns.Add(fieldskeysMap);

            formProperties.Add("FindColumns", listFindColumns);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Editable", SAPbobsCOM.BoYesNoEnum.tYES);
            fieldskeysMap.Add("FormColumnAlias", "Code");
            fieldskeysMap.Add("FormColumnDescription", "Code"); //30 characters
            fieldskeysMap.Add("SonNumber", 0); //30 characters            
            listFormColumns.Add(fieldskeysMap);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Editable", SAPbobsCOM.BoYesNoEnum.tYES);
            fieldskeysMap.Add("FormColumnAlias", "U_OpType");
            fieldskeysMap.Add("FormColumnDescription", "Operation Type"); //30 characters
            fieldskeysMap.Add("SonNumber", 0); //30 characters
            listFormColumns.Add(fieldskeysMap);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Editable", SAPbobsCOM.BoYesNoEnum.tYES);
            fieldskeysMap.Add("FormColumnAlias", "U_AcctCode");
            fieldskeysMap.Add("FormColumnDescription", "Account Code"); //30 characters
            fieldskeysMap.Add("SonNumber", 0); //30 characters
            listFormColumns.Add(fieldskeysMap);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Editable", SAPbobsCOM.BoYesNoEnum.tYES);
            fieldskeysMap.Add("FormColumnAlias", "U_AcctName");
            fieldskeysMap.Add("FormColumnDescription", "Account Name"); //30 characters
            fieldskeysMap.Add("SonNumber", 0); //30 characters
            listFormColumns.Add(fieldskeysMap);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Editable", SAPbobsCOM.BoYesNoEnum.tYES);
            fieldskeysMap.Add("FormColumnAlias", "U_CFWId");
            fieldskeysMap.Add("FormColumnDescription", "Cash Flow Line Item ID"); //30 characters
            fieldskeysMap.Add("SonNumber", 0); //30 characters
            listFormColumns.Add(fieldskeysMap);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Editable", SAPbobsCOM.BoYesNoEnum.tYES);
            fieldskeysMap.Add("FormColumnAlias", "U_CFWName");
            fieldskeysMap.Add("FormColumnDescription", "Cash Flow Line Item Name"); //30 characters
            fieldskeysMap.Add("SonNumber", 0); //30 characters
            listFormColumns.Add(fieldskeysMap);

            if (CommonFunctions.IsDevelopment())
            {
                fieldskeysMap = new Dictionary<string, object>();
                fieldskeysMap.Add("Editable", SAPbobsCOM.BoYesNoEnum.tYES);
                fieldskeysMap.Add("FormColumnAlias", "U_BCFWId");
                fieldskeysMap.Add("FormColumnDescription", "Budget Cash Flow ID"); //30 characters
                fieldskeysMap.Add("SonNumber", 0); //30 characters
                listFormColumns.Add(fieldskeysMap);

                fieldskeysMap = new Dictionary<string, object>();
                fieldskeysMap.Add("Editable", SAPbobsCOM.BoYesNoEnum.tYES);
                fieldskeysMap.Add("FormColumnAlias", "U_BCFWName");
                fieldskeysMap.Add("FormColumnDescription", "Budget Cash Flow Name"); //30 characters
                fieldskeysMap.Add("SonNumber", 0); //30 characters
                listFormColumns.Add(fieldskeysMap);
            }

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Editable", SAPbobsCOM.BoYesNoEnum.tYES);
            fieldskeysMap.Add("FormColumnAlias", "U_TresrCode");
            fieldskeysMap.Add("FormColumnDescription", "Treasury Code"); //30 characters
            fieldskeysMap.Add("SonNumber", 0); //30 characters
            listFormColumns.Add(fieldskeysMap);

            formProperties.Add("FormColumns", listFormColumns);

            UDO.registerUDO( code, formProperties, out errorText);

            GC.Collect();
        }

        public static void updateUDO()
        {
            string code = "UDO_F_BDOSINTR_D";

            List<string> listFindColumns = new List<string>(); //FindColumns

            if (CommonFunctions.IsDevelopment())
            {
            listFindColumns.Add("U_BCFWId");
            listFindColumns.Add("U_BCFWName");
            }

            if (listFindColumns.Count == 0)
                return;

            string queryFindColumns = @"SELECT ""ColAlias""
                                FROM ""UDO3"" 
                                WHERE ""Code"" = '" + code + "'";
            for (int i = 0; i < listFindColumns.Count(); i++)
            {
                string conTxt = (i == 0 ? " AND ( " : " OR ");
                queryFindColumns = queryFindColumns + conTxt + @" ""ColAlias"" = '" + listFindColumns[i] + "'";
            }
            queryFindColumns = queryFindColumns + " )";

            SAPbobsCOM.Recordset oRecordSetFindColumns = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRecordSetFindColumns.DoQuery(queryFindColumns);

            if (oRecordSetFindColumns.RecordCount != listFindColumns.Count)
            {
                Marshal.ReleaseComObject(oRecordSetFindColumns);
                oRecordSetFindColumns = null;

                GC.WaitForPendingFinalizers();

                string errorText = "";
                registerUDO(out errorText);
            }
            else
            {
                Marshal.ReleaseComObject(oRecordSetFindColumns);
                oRecordSetFindColumns = null;

                GC.WaitForPendingFinalizers();
            }
        }

        public static void changeFormItems( SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            bool multiSelection = false;
            string objectType = "1"; //oChartOfAccounts
            string uniqueID_lf_GLAccountCFL = "GLAccount_CFL";
            FormsB1.addChooseFromList( oForm, multiSelection, objectType, uniqueID_lf_GLAccountCFL);

            multiSelection = false;
            objectType = "242"; //CashFlowLineItem
            string uniqueID_lf_CashFlowLineItemCFL = "CashFlowLineItem_CFL";
            FormsB1.addChooseFromList( oForm, multiSelection, objectType, uniqueID_lf_CashFlowLineItemCFL);

            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("3").Specific));

            SAPbouiCOM.Columns oColumns = oMatrix.Columns;
            SAPbouiCOM.Column oColumn;

            oColumn = oColumns.Item("Code");
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Code");

            oColumn = oColumns.Item("U_OpType");
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("OperationType");
            oColumn.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
            SAPbouiCOM.ValidValues oValidValues = oColumn.ValidValues;

            foreach (SAPbouiCOM.ValidValue oValidValue in oValidValues)
            {
                oColumn.ValidValues.Add(oValidValue.Value, BDOSResources.getTranslate(oValidValue.Value));
            }

            oColumn = oColumns.Item("U_AcctCode");
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("AccountCode");
            oColumn.ChooseFromListUID = uniqueID_lf_GLAccountCFL;
            oColumn.ChooseFromListAlias = "AcctCode";

            oColumn = oColumns.Item("U_AcctName");
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("AccountName");

            oColumn = oColumns.Item("U_CFWId");
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("CashFlowLineItemID");
            oColumn.ChooseFromListUID = uniqueID_lf_CashFlowLineItemCFL;
            oColumn.ChooseFromListAlias = "CFWId";

            oColumn = oColumns.Item("U_CFWName");
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("CashFlowLineItemName");

            if (CommonFunctions.IsDevelopment())
            {
                string uniqueID_lf_Budg_CFL = "Budg_CFL";
                multiSelection = false;
                objectType = "UDO_F_BDOSBUCFW_D";
                FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_lf_Budg_CFL);

                oColumn = oColumns.Item("U_BCFWId");
                oColumn.TitleObject.Caption = BDOSResources.getTranslate("BudgetCashFlow");
                oColumn.ChooseFromListUID = uniqueID_lf_Budg_CFL;
                oColumn.ChooseFromListAlias = "Code";

                oColumn = oColumns.Item("U_BCFWName");
                oColumn.TitleObject.Caption = BDOSResources.getTranslate("Name");

                oColumns.Item("U_BCFWId").Visible = true;
                oColumns.Item("U_BCFWName").Visible = true;

            }
            else
            {
                try
                {
                    oColumns.Item("U_BCFWId").Visible = false;
                }
                catch
                { }

                try
                {
                    oColumns.Item("U_BCFWName").Visible = false;
                }
                catch
                { }
            }

            oColumn = oColumns.Item("U_TresrCode");
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("TreasuryCode");
        }

        public static void addMenus()
        {
            SAPbouiCOM.MenuItem menuItem;
            SAPbouiCOM.MenuItem fatherMenuItem;
            SAPbouiCOM.MenuCreationParams oCreationPackage;

            try
            {
                fatherMenuItem = Program.uiApp.Menus.Item("11264");

                // Add a pop-up menu item
                oCreationPackage = (SAPbouiCOM.MenuCreationParams)Program.uiApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                oCreationPackage.Checked = false;
                oCreationPackage.Enabled = true;
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "UDO_F_BDOSINTR_D";
                oCreationPackage.String = BDOSResources.getTranslate("InternetBankingIntegrationServicesRules");
                oCreationPackage.Position = -1;

                menuItem = fatherMenuItem.SubMenus.AddEx(oCreationPackage);
            }
            catch 
            {
               
            }
        }

        public static void setSizeForm( SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            try
            {
                oForm.ClientHeight = Program.uiApp.Desktop.Width / 6;
                //oForm.ClientWidth = Program.uiApp.Desktop.Width / 3;

                oForm.Height = Program.uiApp.Desktop.Width / 4;
                //oForm.ClientWidth = Program.uiApp.Desktop.Width / 2;

                oForm.Left = (Program.uiApp.Desktop.Width - oForm.Width) / 2;
                oForm.Top = (Program.uiApp.Desktop.Height - oForm.Height) / 3;
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

        public static void resizeForm( SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            SAPbouiCOM.Item oItem = null;

            try
            {
                oItem = oForm.Items.Item("1");
                oItem.Top = oForm.ClientHeight - 25;

                oItem = oForm.Items.Item("2");
                oItem.Top = oForm.ClientHeight - 25;
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

        public static void chooseFromList(SAPbouiCOM.Form oForm, SAPbouiCOM.IChooseFromListEvent oCFLEvento, SAPbouiCOM.ItemEvent pVal, out string errorText)
        {
            errorText = null;

            try
            {
                string sCFL_ID = oCFLEvento.ChooseFromListUID;
                SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(sCFL_ID);

                if (pVal.BeforeAction == true)
                {
                    if (sCFL_ID == "GLAccount_CFL")
                    {
                        if (pVal.ItemUID == "3" && pVal.ColUID == "U_AcctCode")
                        {
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
                            oCon.Alias = "LocManTran"; //Lock Manual Transaction (Control Account)
                            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                            oCon.CondVal = "N";

                            oCFL.SetConditions(oCons);
                        }
                    }
                    else if (sCFL_ID == "CashFlowLineItem_CFL")
                    {
                        SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();

                        SAPbouiCOM.Condition oCon = oCons.Add();
                        oCon.Alias = "Postable"; //Active Account, (Title Account)
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = "Y";


                        oCFL.SetConditions(oCons);
                    }
                }
                else
                {
                    SAPbouiCOM.DataTable oDataTable = null;
                    oDataTable = oCFLEvento.SelectedObjects;

                    if (oDataTable != null)
                    {
                        if (sCFL_ID == "GLAccount_CFL")
                        {
                            if (pVal.ItemUID == "3" && pVal.ColUID == "U_AcctCode")
                            {
                                string acctCode = Convert.ToString(oDataTable.GetValue("AcctCode", 0));
                                string acctName = Convert.ToString(oDataTable.GetValue("AcctName", 0));

                                SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("3").Specific));
                                SAPbouiCOM.CellPosition cellPos = oMatrix.GetCellFocus();
                                if (cellPos == null)
                                {
                                    return;
                                }
                                SAPbouiCOM.EditText oEditText;
                                try
                                {
                                    oEditText = oMatrix.Columns.Item("U_AcctCode").Cells.Item(cellPos.rowIndex).Specific;
                                    oEditText.Value = acctCode;
                                }
                                catch { }

                                try
                                {
                                    oEditText = oMatrix.Columns.Item("U_AcctName").Cells.Item(cellPos.rowIndex).Specific;
                                    oEditText.Value = acctName;
                                }
                                catch { }
                            }
                        }
                        else if (sCFL_ID == "CashFlowLineItem_CFL")
                        {
                            if (pVal.ItemUID == "3" && pVal.ColUID == "U_CFWId")
                            {
                                string CFWId = Convert.ToString(oDataTable.GetValue("CFWId", 0));
                                string CFWName = Convert.ToString(oDataTable.GetValue("CFWName", 0));
                                SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("3").Specific));
                                SAPbouiCOM.CellPosition cellPos = oMatrix.GetCellFocus();
                                if (cellPos == null)
                                {
                                    return;
                                }

                                SAPbouiCOM.EditText oEditText;

                                try
                                {
                                    oEditText = oMatrix.Columns.Item("U_CFWId").Cells.Item(cellPos.rowIndex).Specific;
                                    oEditText.Value = CFWId;
                                }
                                catch { }

                                try
                                {
                                    oEditText = oMatrix.Columns.Item("U_CFWName").Cells.Item(cellPos.rowIndex).Specific;
                                    oEditText.Value = CFWName;
                                }
                                catch { }
                            }
                        }
                        //Budg_CFL
                        else if (sCFL_ID == "Budg_CFL")
                        {
                            if (pVal.ItemUID == "3" && pVal.ColUID == "U_BCFWId")
                            {
                                string BCFWId = Convert.ToString(oDataTable.GetValue("Code", 0));
                                string BCFWName = Convert.ToString(oDataTable.GetValue("Name", 0));
                                SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("3").Specific));
                                SAPbouiCOM.CellPosition cellPos = oMatrix.GetCellFocus();
                                if (cellPos == null)
                                {
                                    return;
                                }

                                SAPbouiCOM.EditText oEditText;

                                try
                                {
                                    oEditText = oMatrix.Columns.Item("U_BCFWId").Cells.Item(cellPos.rowIndex).Specific;
                                    oEditText.Value = BCFWId;
                                }
                                catch { }

                                try
                                {
                                    oEditText = oMatrix.Columns.Item("U_BCFWName").Cells.Item(cellPos.rowIndex).Specific;
                                    oEditText.Value = BCFWName;
                                }
                                catch { }
                            }
                        }
                    }
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

        public static void validate(SAPbouiCOM.Form oForm, SAPbouiCOM.ItemEvent pVal, out string errorText)
        {
            errorText = null;

            try
            {
                if (pVal.ItemUID == "3" && pVal.ColUID == "U_CFWId")
                {
                    SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("3").Specific));
                    SAPbouiCOM.CellPosition cellPos = oMatrix.GetCellFocus();

                    if (cellPos == null)
                    {
                        return;
                    }

                    SAPbouiCOM.EditText oEditText;
                    oEditText = oMatrix.Columns.Item("U_CFWId").Cells.Item(cellPos.rowIndex).Specific;

                    if (oEditText.Value == "")
                    {
                        try
                        {
                            oEditText.Value = "";
                        }
                        catch { }

                        try
                        {
                            oEditText = oMatrix.Columns.Item("U_CFWName").Cells.Item(cellPos.rowIndex).Specific;
                            oEditText.Value = "";
                        }
                        catch { }
                    }
                }
                else if (pVal.ItemUID == "3" && pVal.ColUID == "U_AcctCode")
                {
                    SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("3").Specific));
                    SAPbouiCOM.CellPosition cellPos = oMatrix.GetCellFocus();

                    if (cellPos == null)
                    {
                        return;
                    }

                    SAPbouiCOM.EditText oEditText;
                    oEditText = oMatrix.Columns.Item("U_AcctCode").Cells.Item(cellPos.rowIndex).Specific;

                    if (oEditText.Value == "")
                    {
                        try
                        {
                            oEditText.Value = "";
                        }
                        catch { }

                        try
                        {
                            oEditText = oMatrix.Columns.Item("U_AcctName").Cells.Item(cellPos.rowIndex).Specific;
                            oEditText.Value = "";
                        }
                        catch { }
                    }
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

        public static SAPbobsCOM.Recordset getRules( OperationTypeFromIntBank opType, string treasuryCode)
        {
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                if (opType == OperationTypeFromIntBank.None)
                {
                    return null;
                }

                string additionalFieldsQueryStr = "";
                if (CommonFunctions.IsDevelopment())
                {
                    additionalFieldsQueryStr = additionalFieldsQueryStr + @",
                         ""U_BCFWId"",
                         ""U_BCFWName""";
                }

                string query = @"SELECT
                    	 ""Code"",
                     	 ""U_OpType"",
                    	 ""U_AcctCode"",
                    	 ""U_AcctName"",
                    	 ""U_CFWId"",
                    	 ""U_CFWName"",
                    	 ""U_TresrCode""" + additionalFieldsQueryStr + 
                    @" FROM ""@BDOSINTR"" WHERE ""U_OpType"" = '" + opType + "'";

                if (string.IsNullOrEmpty(treasuryCode) == false && opType != OperationTypeFromIntBank.BankCharge)
                {
                    query = query + @" AND ""U_TresrCode"" = '" + treasuryCode + "'";
                }

                oRecordSet.DoQuery(query);
                if (!oRecordSet.EoF)
                {
                    return oRecordSet;
                }

                Marshal.FinalReleaseComObject(oRecordSet);
                oRecordSet = null;
                return null;
            }
            catch 
            {
                return null;
            }
        }

        public static void uiApp_ItemEvent(  string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD & pVal.BeforeAction == false)
                {
                    changeFormItems( oForm, out errorText);
                    oForm.Title = BDOSResources.getTranslate("InternetBankingIntegrationServicesRules");
                    Program.FORM_LOAD_FOR_VISIBLE = true;
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE & pVal.BeforeAction == false & Program.FORM_LOAD_FOR_VISIBLE == true)
                {
                    oForm.Freeze(true);
                    setSizeForm( oForm, out errorText);
                    oForm.Title = BDOSResources.getTranslate("InternetBankingIntegrationServicesRules");
                    oForm.Freeze(false);
                    Program.FORM_LOAD_FOR_VISIBLE = false;
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_RESIZE & pVal.BeforeAction == false)
                {
                    oForm.Freeze(true);
                    resizeForm( oForm, out errorText);
                    oForm.Freeze(false);
                }

                if (pVal.ItemUID == "3" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                {
                    SAPbouiCOM.IChooseFromListEvent oCFLEvento = null;
                    oCFLEvento = ((SAPbouiCOM.IChooseFromListEvent)(pVal));

                    chooseFromList(oForm, oCFLEvento, pVal, out errorText);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_VALIDATE && pVal.ItemChanged && pVal.BeforeAction == false)
                {
                    validate(oForm, pVal, out errorText);
                }
            }
        }
    }

    public enum OperationTypeFromIntBank
    {
        None,
        TransferToOwnAccount,
        CurrencyExchange,
        TransferFromBP,
        TransferToBP,
        TreasuryTransfer,
        ReturnToCustomer,
        ReturnFromSupplier,
        BankCharge,
        OtherIncomes,
        OtherExpenses,
        Salary,
        WithoutSalary,
        TreasuryTransferPaymentOrderIoBP
    };
}
