﻿using System;
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

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BankCode");
            fieldskeysMap.Add("TableName", "BDOSCRLN");
            fieldskeysMap.Add("Description", "Bank Code");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 30);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "CrLnAcct");
            fieldskeysMap.Add("TableName", "BDOSCRLN");
            fieldskeysMap.Add("Description", "Credit Line Account Code");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 15);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "IntrstRate");
            fieldskeysMap.Add("TableName", "BDOSCRLN");
            fieldskeysMap.Add("Description", "Interest Rate");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Percentage);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "StartDate");
            fieldskeysMap.Add("TableName", "BDOSCRLN");
            fieldskeysMap.Add("Description", "Starting Date");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Date);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "ExpnsAcct");
            fieldskeysMap.Add("TableName", "BDOSCRLN");
            fieldskeysMap.Add("Description", "Expense Account Code");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 15);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "IntPblAcct");
            fieldskeysMap.Add("TableName", "BDOSCRLN");
            fieldskeysMap.Add("Description", "Interest Payable Account Code");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 15);

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
            oUserObjectMD = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD) as SAPbobsCOM.UserObjectsMD;
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
                oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES;
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

                if (!retval)
                {
                    if (oUserObjectMD.Add() != 0)
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

            //if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE)
            //{
            //    if (BusinessObjectInfo.BeforeAction)
            //    {
            //        if (checkRemoving(oForm))
            //        {
            //            Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("RecordIsUsedInDocuments"));
            //            Program.uiApp.MessageBox(BDOSResources.getTranslate("OperationUnsuccesfullSeeLog"));
            //            BubbleEvent = false;
            //        }
            //    }
            //}
        }

        public static void changeFormItems(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("3").Specific;

            SAPbouiCOM.Columns oColumns = oMatrix.Columns;

            SAPbouiCOM.Column oColumn = oColumns.Item("Code");
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Code");

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

        //public static bool checkRemoving(SAPbouiCOM.Form oForm)
        //{
        //    SAPbouiCOM.DBDataSource DocDBSourceTAXP = oForm.DataSources.DBDataSources.Item(0);
        //    string code = DocDBSourceTAXP.GetValue("Code", 0).Trim();

        //    Dictionary<string, string> listTables = new Dictionary<string, string>();
        //    listTables.Add("@BDOSFUN1", "U_CrtrCode");

        //    return CommonFunctions.codeIsUsed(listTables, code);
        //}
    }
}
