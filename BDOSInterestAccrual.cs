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
    class BDOSInterestAccrual
    {
        public static string bankCodeOld;
        public static string crLnCodeOld;

        public static void createDocumentUDO(out string errorText)
        {
            string tableName = "BDOSINAC";
            string description = "Interest Accrual Document";

            int result = UDO.addUserTable(tableName, description, SAPbobsCOM.BoUTBTableType.bott_Document, out errorText);

            if (result != 0)
            {
                return;
            }

            Dictionary<string, object> fieldskeysMap;

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "DocDate");
            fieldskeysMap.Add("TableName", "BDOSINAC");
            fieldskeysMap.Add("Description", "Posting Date");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Date);
            fieldskeysMap.Add("EditSize", 15);
            fieldskeysMap.Add("Mandatory", SAPbobsCOM.BoYesNoEnum.tYES);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BankCode");
            fieldskeysMap.Add("TableName", "BDOSINAC");
            fieldskeysMap.Add("Description", "Bank Code");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 30);
            fieldskeysMap.Add("Mandatory", SAPbobsCOM.BoYesNoEnum.tYES);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "TransId");
            fieldskeysMap.Add("TableName", "BDOSINAC");
            fieldskeysMap.Add("Description", "Transaction Number");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Numeric);
            fieldskeysMap.Add("EditSize", 11);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            tableName = "BDOSINA1";
            description = "Interest Accrual Child1";

            result = UDO.addUserTable(tableName, description, SAPbobsCOM.BoUTBTableType.bott_DocumentLines, out errorText);

            if (result != 0)
            {
                return;
            }

            //Credit Line Master Data
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "CrLnCode");
            fieldskeysMap.Add("TableName", "BDOSINA1");
            fieldskeysMap.Add("Description", "Credit Line Code");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);
            fieldskeysMap.Add("Mandatory", SAPbobsCOM.BoYesNoEnum.tYES);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            //Credit Line Master Data
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "CrLnName");
            fieldskeysMap.Add("TableName", "BDOSINA1");
            fieldskeysMap.Add("Description", "Credit Line Name");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 100);
            fieldskeysMap.Add("Mandatory", SAPbobsCOM.BoYesNoEnum.tYES);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "LnCurrCode");
            fieldskeysMap.Add("TableName", "BDOSINA1");
            fieldskeysMap.Add("Description", "Loan Currency Code");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 3);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "ExchngRate");
            fieldskeysMap.Add("TableName", "BDOSINA1");
            fieldskeysMap.Add("Description", "Exchange Rate");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Rate);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "CrLnAcct");
            fieldskeysMap.Add("TableName", "BDOSINA1");
            fieldskeysMap.Add("Description", "Credit Line Account Code");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 15);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "IntrstRate");
            fieldskeysMap.Add("TableName", "BDOSINA1");
            fieldskeysMap.Add("Description", "Interest Rate");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Percentage);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "CrLnAmtLC");
            fieldskeysMap.Add("TableName", "BDOSINA1");
            fieldskeysMap.Add("Description", "Credit Line Balance LC");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "CrLnAmtFC");
            fieldskeysMap.Add("TableName", "BDOSINA1");
            fieldskeysMap.Add("Description", "Credit Line Balance FC");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "IntAmtLC");
            fieldskeysMap.Add("TableName", "BDOSINA1");
            fieldskeysMap.Add("Description", "Interest Amount LC");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "IntAmtFC");
            fieldskeysMap.Add("TableName", "BDOSINA1");
            fieldskeysMap.Add("Description", "Interest Amount FC");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "ExpnsAcct");
            fieldskeysMap.Add("TableName", "BDOSINA1");
            fieldskeysMap.Add("Description", "Expense Account Code");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 15);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "IntPblAcct");
            fieldskeysMap.Add("TableName", "BDOSINA1");
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
            SAPbobsCOM.IUserObjectMD_ChildTables oUDOChildTables = null;
            GC.Collect();
            oUserObjectMD = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);
            Marshal.ReleaseComObject(oUserObjectMD);
            oUserObjectMD = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD) as SAPbobsCOM.UserObjectsMD;
            oUDOFind = oUserObjectMD.FindColumns;
            oUDOForm = oUserObjectMD.FormColumns;
            oUDOChildTables = oUserObjectMD.ChildTables;

            var retval = oUserObjectMD.GetByKey("UDO_F_BDOSINAC_D");

            if (!retval)
            {
                oUserObjectMD.Code = "UDO_F_BDOSINAC_D";
                oUserObjectMD.Name = "Interest Accrual Document";
                oUserObjectMD.TableName = "BDOSINAC";
                oUserObjectMD.ObjectType = SAPbobsCOM.BoUDOObjType.boud_Document;
                oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO;
                oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES;

                //Find
                oUDOFind.ColumnAlias = "DocEntry";
                oUDOFind.ColumnDescription = "Internal Number";
                oUDOFind.Add();
                oUDOFind.ColumnAlias = "DocNum";
                oUDOFind.ColumnDescription = "Document Number";
                oUDOFind.Add();
                oUDOFind.ColumnAlias = "CreateDate";
                oUDOFind.ColumnDescription = "Create Date";
                oUDOFind.Add();
                oUDOFind.ColumnAlias = "UpdateDate";
                oUDOFind.ColumnDescription = "Update Date";
                oUDOFind.Add();
                oUDOFind.ColumnAlias = "Status";
                oUDOFind.ColumnDescription = "Status";
                oUDOFind.Add();
                oUDOFind.ColumnAlias = "Canceled";
                oUDOFind.ColumnDescription = "Canceled";
                oUDOFind.Add();
                oUDOFind.ColumnAlias = "U_DocDate";
                oUDOFind.ColumnDescription = "Posting Date";
                oUDOFind.Add();
                oUDOFind.ColumnAlias = "U_BankCode";
                oUDOFind.ColumnDescription = "Bank Code";
                oUDOFind.Add();
                oUDOFind.ColumnAlias = "U_TransId";
                oUDOFind.ColumnDescription = "Transaction Number";
                oUDOFind.Add();
                oUDOFind.ColumnAlias = "Remark";
                oUDOFind.ColumnDescription = "Remark";
                oUDOFind.Add();

                //Form
                oUDOForm.FormColumnAlias = "DocEntry";
                oUDOForm.FormColumnDescription = "Internal Number";
                oUDOForm.Add();

                oUDOChildTables.Add();
                oUDOChildTables.SetCurrentLine(oUDOChildTables.Count - 1);
                oUDOChildTables.TableName = "BDOSINA1";
                oUDOChildTables.ObjectName = "BDOSINA1";

                if (!retval)
                {
                    if (oUserObjectMD.Add() != 0)
                    {
                        Program.uiApp.MessageBox(Program.oCompany.GetLastErrorDescription());
                    }
                }
            }
            Marshal.ReleaseComObject(oUDOForm);
            Marshal.ReleaseComObject(oUDOFind);
            Marshal.ReleaseComObject(oUDOChildTables);
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
                oCreationPackage.UniqueID = "UDO_F_BDOSINAC_D";
                oCreationPackage.String = BDOSResources.getTranslate("InterestAccrualDocument");
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

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD && pVal.BeforeAction)
                {
                    createFormItems(oForm);
                    Program.FORM_LOAD_FOR_VISIBLE = true;
                    Program.FORM_LOAD_FOR_ACTIVATE = true;
                }

                else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE && !pVal.BeforeAction)
                {
                    if (Program.FORM_LOAD_FOR_VISIBLE)
                    {
                        setSizeForm(oForm);
                        oForm.Freeze(true);
                        oForm.Title = BDOSResources.getTranslate("InterestAccrualDocument");
                        oForm.Freeze(false);
                        Program.FORM_LOAD_FOR_VISIBLE = false;
                        setVisibleFormItems(oForm);
                    }
                }

                else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_RESIZE && !pVal.BeforeAction)
                {
                    resizeForm(oForm);
                }

                else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE && !pVal.BeforeAction)
                {
                    if (Program.FORM_LOAD_FOR_ACTIVATE)
                    {
                        oForm.Freeze(true);
                        try
                        {
                            SAPbouiCOM.StaticText staticText = oForm.Items.Item("0_U_S").Specific;
                            staticText.Caption = BDOSResources.getTranslate("DocEntry");

                            Program.FORM_LOAD_FOR_ACTIVATE = false;
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

                else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                {
                    SAPbouiCOM.IChooseFromListEvent oCFLEvento = (SAPbouiCOM.IChooseFromListEvent)pVal;
                    chooseFromList(oForm, pVal, oCFLEvento, ref BubbleEvent);
                }

                else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK)
                {
                    if (!pVal.BeforeAction)
                    {
                        if (pVal.FormMode == 3)
                        {
                            if (pVal.ItemUID == "addMTRB")
                                addMatrixRow(oForm);
                            else if (pVal.ItemUID == "delMTRB")
                                deleteMatrixRow(oForm);
                        }
                    }
                }

                else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_GOT_FOCUS)
                {
                    if (!pVal.BeforeAction)
                    {
                        if (pVal.ItemUID == "BankCodeE")
                        {
                            bankCodeOld = oForm.DataSources.DBDataSources.Item("@BDOSINAC").GetValue("U_BankCode", 0);
                        }
                        else if (pVal.ItemUID == "LoanMTR" && pVal.ColUID == "CrLnCode")
                        {
                            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(pVal.ItemUID).Specific;
                            crLnCodeOld = oMatrix.GetCellSpecific(pVal.ColUID, pVal.Row).Value;
                        }
                    }
                }

                else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_LOST_FOCUS)
                {
                    if (!pVal.BeforeAction)
                    {
                        if (pVal.ItemUID == "BankCodeE")
                        {
                            string bankCode = oForm.DataSources.DBDataSources.Item("@BDOSINAC").GetValue("U_BankCode", 0);
                            if (bankCode != bankCodeOld && !string.IsNullOrEmpty(bankCodeOld) && string.IsNullOrEmpty(bankCode))
                            {
                                clearMatrix(oForm);
                                bankCodeOld = null;
                            }
                        }
                        else if (pVal.ItemUID == "LoanMTR" && pVal.ColUID == "CrLnCode")
                        {
                            oForm.Freeze(true);
                            try
                            {
                                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(pVal.ItemUID).Specific;
                                string crLnCode = oMatrix.GetCellSpecific(pVal.ColUID, pVal.Row).Value;
                                if (crLnCode != crLnCodeOld && !string.IsNullOrEmpty(crLnCodeOld) && string.IsNullOrEmpty(crLnCode))
                                {
                                    updateRowByCreditLine(oForm, null, pVal.Row - 1);
                                    updateExchangeRateRow(oForm, pVal.Row);
                                    crLnCodeOld = null;
                                }
                            }
                            catch (Exception ex)
                            {
                                crLnCodeOld = null;
                                throw new Exception(ex.Message);
                            }
                            finally
                            {
                                oForm.Freeze(false);
                            }
                        }
                    }
                }

                else if (pVal.ItemChanged)
                {
                    if (pVal.ItemUID == "DocDateE" && !pVal.BeforeAction)
                        updateExchangeRateRow(oForm);
                }
            }
        }

        public static void uiApp_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;

            SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(BusinessObjectInfo.FormTypeEx, Program.currentFormCount);

            if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD && !BusinessObjectInfo.BeforeAction)
            {
                setVisibleFormItems(oForm);
            }

            else if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD || BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE)
            {
                if (BusinessObjectInfo.BeforeAction)
                {
                    if (oForm.DataSources.DBDataSources.Item("@BDOSINAC").GetValue("Canceled", 0) == "N" && !Program.cancellationTrans)
                    {
                        SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("LoanMTR").Specific;
                        oMatrix.FlushToDataSource();
                        if (oMatrix.RowCount > 1)
                        {
                            //checkDuplicatesInDBDataSources
                            string errorText;
                            SAPbouiCOM.DBDataSource oDBDataSourceMTR = oForm.DataSources.DBDataSources.Item("@BDOSINA1");
                            Dictionary<string, SAPbouiCOM.DBDataSource> oKeysDictionary = new Dictionary<string, SAPbouiCOM.DBDataSource>();
                            oKeysDictionary.Add("U_CrLnCode", oDBDataSourceMTR);
                            List<string> creditLineList = CommonFunctions.checkDuplicatesInDBDataSources(oDBDataSourceMTR, oKeysDictionary, out errorText);
                            if (!string.IsNullOrEmpty(errorText))
                            {
                                Program.uiApp.SetStatusBarMessage(errorText + " " + BDOSResources.getTranslate("CreditLine") + ": " + string.Join(",", creditLineList), SAPbouiCOM.BoMessageTime.bmt_Short);
                                BubbleEvent = false;
                                return;
                            }
                        }
                    }
                }
                else if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE && !BusinessObjectInfo.BeforeAction && BusinessObjectInfo.ActionSuccess)
                {
                    if (Program.cancellationTrans && Program.canceledDocEntry != 0)
                    {
                        cancellation(oForm, Program.canceledDocEntry);
                    }
                }
            }           
        }

        public static void createFormItems(SAPbouiCOM.Form oForm)
        {
            string errorText;

            oForm.AutoManaged = true;

            Dictionary<string, object> formItems;
            string itemName;
            int left_s = 6;
            int left_e = 160;
            int height = 15;
            int top = 5;
            int width_s = 139;
            int width_e = 140;

            int left_s2 = 300;
            int left_e2 = left_s2 + 121;
            int top2 = 5;

            top += (height + 1);

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
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSINAC");
            formItems.Add("Alias", "U_BankCode");
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
            oForm.Items.Item(itemName).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False); //OK mode
            oForm.Items.Item(itemName).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 8, SAPbouiCOM.BoModeVisualBehavior.mvb_False); //View mode

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

            top = top + height + 1;

            FormsB1.addChooseFromList(oForm, false, "30", "JournalEntryCFL");
            formItems = new Dictionary<string, object>();
            itemName = "TransIdS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("TransactionNo"));
            formItems.Add("LinkTo", "TransIdE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }

            formItems = new Dictionary<string, object>();
            itemName = "TransIdE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSINAC");
            formItems.Add("Alias", "U_TransId");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            //formItems.Add("ChooseFromListUID", "JournalEntryCFL");
            //formItems.Add("ChooseFromListAlias", "TransId");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }
            oForm.Items.Item(itemName).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False); //All modes
            oForm.Items.Item(itemName).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True); //Find mode

            formItems = new Dictionary<string, object>();
            itemName = "TransIdLB"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            formItems.Add("Left", left_e - 20);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("LinkTo", "TransIdE");
            formItems.Add("LinkedObjectType", "30"); //Journal Entry

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "No.S"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s2);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top2);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Number"));
            formItems.Add("LinkTo", "DocNumE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }

            formItems = new Dictionary<string, object>();
            itemName = "DocNumE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSINAC");
            formItems.Add("Alias", "DocNum");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e2);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top2);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }
            oForm.Items.Item(itemName).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False); //All modes
            oForm.Items.Item(itemName).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True); //Find mode

            top2 += height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "StatusS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s2);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top2);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Status"));
            formItems.Add("LinkTo", "StatusC");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }

            formItems = new Dictionary<string, object>();
            itemName = "StatusC"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSINAC");
            formItems.Add("Alias", "Status");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            formItems.Add("Left", left_e2);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top2);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
            formItems.Add("DisplayDesc", true);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }
            oForm.Items.Item(itemName).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False); //All modes
            oForm.Items.Item(itemName).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True); //Find mode

            top2 += height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "CanceledS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s2);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top2);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Canceled"));
            formItems.Add("LinkTo", "CanceledC");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }

            formItems = new Dictionary<string, object>();
            itemName = "CanceledC"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSINAC");
            formItems.Add("Alias", "Canceled");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            formItems.Add("Left", left_e2);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top2);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
            formItems.Add("DisplayDesc", true);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }
            oForm.Items.Item(itemName).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False); //All modes
            oForm.Items.Item(itemName).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True); //Find mode

            top2 += height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "CreateDatS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s2);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top2);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("CreateDate"));
            formItems.Add("LinkTo", "CreateDatE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }

            formItems = new Dictionary<string, object>();
            itemName = "CreateDatE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSINAC");
            formItems.Add("Alias", "CreateDate");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e2);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top2);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }
            oForm.Items.Item(itemName).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False); //All modes
            oForm.Items.Item(itemName).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True); //Find mode

            top2 += height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "UpdateDatS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s2);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top2);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("UpdateDate"));
            formItems.Add("LinkTo", "UpdateDatE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }

            formItems = new Dictionary<string, object>();
            itemName = "UpdateDatE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSINAC");
            formItems.Add("Alias", "UpdateDate");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e2);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top2);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }
            oForm.Items.Item(itemName).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False); //All modes
            oForm.Items.Item(itemName).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True); //Find mode

            top2 += height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "DocDateS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s2);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top2);
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
            itemName = "DocDateE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSINAC");
            formItems.Add("Alias", "U_DocDate");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e2);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top2);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }
            oForm.Items.Item(itemName).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False); //OK mode
            oForm.Items.Item(itemName).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 8, SAPbouiCOM.BoModeVisualBehavior.mvb_False); //View mode

            top += (3 * height + 1);

            formItems = new Dictionary<string, object>();
            itemName = "addMTRB"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            formItems.Add("Left", left_s);
            formItems.Add("Width", 70);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Add"));
            formItems.Add("Visible", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }
            oForm.Items.Item(itemName).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False); //OK mode
            oForm.Items.Item(itemName).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 8, SAPbouiCOM.BoModeVisualBehavior.mvb_False); //View mode

            formItems = new Dictionary<string, object>();
            itemName = "delMTRB"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            formItems.Add("Left", left_s + 70 + 1);
            formItems.Add("Width", 70);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Delete"));
            formItems.Add("Visible", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }
            oForm.Items.Item(itemName).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False); //OK mode
            oForm.Items.Item(itemName).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 8, SAPbouiCOM.BoModeVisualBehavior.mvb_False); //View mode

            top += height + 1;

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

            oForm.DataSources.DBDataSources.Add("@BDOSINA1");

            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("LoanMTR").Specific;
            oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
            SAPbouiCOM.Columns oColumns = oMatrix.Columns;

            SAPbouiCOM.Column oColumn = oColumns.Add("LineID", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "#";
            oColumn.Editable = false;
            oColumn.DataBind.SetBound(true, "@BDOSINA1", "LineId");

            FormsB1.addChooseFromList(oForm, false, "UDO_F_BDOSCRLN_D", "CreditLineCodeCFL");
            oColumn = oColumns.Add("CrLnCode", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON); //Credit Line Code
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("CreditLine");
            oColumn.DataBind.SetBound(true, "@BDOSINA1", "U_CrLnCode");
            oColumn.ChooseFromListUID = "CreditLineCodeCFL";
            oColumn.ChooseFromListAlias = "Code";
            SAPbouiCOM.LinkedButton oLink = oColumn.ExtendedObject;
            oLink.LinkedObjectType = "UDO_F_BDOSCRLN_D"; //Credit Line Master Data

            oColumn = oColumns.Add("CrLnName", SAPbouiCOM.BoFormItemTypes.it_EDIT); //Credit Line Name
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("CreditLineCode");
            oColumn.Editable = false;
            oColumn.DataBind.SetBound(true, "@BDOSINA1", "U_CrLnName");

            FormsB1.addChooseFromList(oForm, false, "37", "CurrencyCFL");
            oColumn = oColumns.Add("LnCurrCode", SAPbouiCOM.BoFormItemTypes.it_EDIT); //Loan Currency Code
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Currency");
            oColumn.Editable = false;
            oColumn.DataBind.SetBound(true, "@BDOSINA1", "U_LnCurrCode");
            oColumn.ChooseFromListUID = "CurrencyCFL";
            oColumn.ChooseFromListAlias = "CurrCode"; //Currency Codes

            oColumn = oColumns.Add("ExchngRate", SAPbouiCOM.BoFormItemTypes.it_EDIT); //Exchange Rate
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("ExchangeRate");
            oColumn.Editable = false;
            oColumn.DataBind.SetBound(true, "@BDOSINA1", "U_ExchngRate");

            FormsB1.addChooseFromList(oForm, false, "1", "CrLnAcctCFL");
            oColumn = oColumns.Add("CrLnAcct", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON); //Credit Line Account Code
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("CreditLineAccount");
            oColumn.Editable = false;
            oColumn.DataBind.SetBound(true, "@BDOSINA1", "U_CrLnAcct");
            oColumn.ChooseFromListUID = "CrLnAcctCFL";
            oColumn.ChooseFromListAlias = "AcctCode";
            oLink = oColumn.ExtendedObject;
            oLink.LinkedObjectType = "1"; //G/L Accounts

            oColumn = oColumns.Add("IntrstRate", SAPbouiCOM.BoFormItemTypes.it_EDIT); //Interest Rate
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("InterestRate");
            oColumn.Editable = false;
            oColumn.DataBind.SetBound(true, "@BDOSINA1", "U_IntrstRate");

            oColumn = oColumns.Add("CrLnAmtLC", SAPbouiCOM.BoFormItemTypes.it_EDIT); //Credit Line Balance LC
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("CreditLineBalanceLC");
            oColumn.Editable = false;
            oColumn.DataBind.SetBound(true, "@BDOSINA1", "U_CrLnAmtLC");

            oColumn = oColumns.Add("CrLnAmtFC", SAPbouiCOM.BoFormItemTypes.it_EDIT); //Credit Line Balance FC
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("CreditLineBalanceFC");
            oColumn.Editable = false;
            oColumn.DataBind.SetBound(true, "@BDOSINA1", "U_CrLnAmtFC");

            oColumn = oColumns.Add("IntAmtLC", SAPbouiCOM.BoFormItemTypes.it_EDIT); //Interest Amount LC
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("InterestAmountLC");
            oColumn.Editable = false;
            oColumn.DataBind.SetBound(true, "@BDOSINA1", "U_IntAmtLC");

            oColumn = oColumns.Add("IntAmtFC", SAPbouiCOM.BoFormItemTypes.it_EDIT); //Interest Amount FC
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("InterestAmountFC");
            oColumn.Editable = false;
            oColumn.DataBind.SetBound(true, "@BDOSINA1", "U_IntAmtFC");

            FormsB1.addChooseFromList(oForm, false, "1", "ExpnsAcctCFL");
            oColumn = oColumns.Add("ExpnsAcct", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON); //Expense Account Code
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("ExpenseAccount");
            oColumn.Editable = false;
            oColumn.DataBind.SetBound(true, "@BDOSINA1", "U_ExpnsAcct");
            oColumn.ChooseFromListUID = "ExpnsAcctCFL";
            oColumn.ChooseFromListAlias = "AcctCode";
            oLink = oColumn.ExtendedObject;
            oLink.LinkedObjectType = "1"; //G/L Accounts

            FormsB1.addChooseFromList(oForm, false, "1", "IntPblAcctCFL");
            oColumn = oColumns.Add("IntPblAcct", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON); //Interest Payable Account Code
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("InterestPayableAccount");
            oColumn.Editable = false;
            oColumn.DataBind.SetBound(true, "@BDOSINA1", "U_IntPblAcct");
            oColumn.ChooseFromListUID = "IntPblAcctCFL";
            oColumn.ChooseFromListAlias = "AcctCode";
            oLink = oColumn.ExtendedObject;
            oLink.LinkedObjectType = "1"; //G/L Accounts

            oForm.Items.Item(itemName).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False); //OK mode
            oForm.Items.Item(itemName).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 8, SAPbouiCOM.BoModeVisualBehavior.mvb_False); //View mode

            top = top + oForm.Items.Item("LoanMTR").Height + 10;

            formItems = new Dictionary<string, object>();
            itemName = "CreatorS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Creator"));
            formItems.Add("LinkTo", "CreatorE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }

            formItems = new Dictionary<string, object>();
            itemName = "CreatorE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSINAC");
            formItems.Add("Alias", "Creator");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
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
            oForm.Items.Item(itemName).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False); //All modes
            oForm.Items.Item(itemName).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True); //Find mode

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "RemarksS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Remarks"));
            formItems.Add("LinkTo", "RemarksE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }

            formItems = new Dictionary<string, object>();
            itemName = "RemarksE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSINAC");
            formItems.Add("Alias", "Remark");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EXTEDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e * 3);
            formItems.Add("Top", top);
            formItems.Add("Height", 3 * height);
            formItems.Add("UID", itemName);
            formItems.Add("ScrollBars", SAPbouiCOM.BoScrollBars.sb_Vertical);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }

            GC.Collect();
        }

        public static void chooseFromList(SAPbouiCOM.Form oForm, SAPbouiCOM.ItemEvent pVal, SAPbouiCOM.IChooseFromListEvent oCFLEvento, ref bool bubbleEvent)
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
                    else if (oCFLEvento.ChooseFromListUID == "CreditLineCodeCFL")
                    {
                        SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID);
                        SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();
                        string bankCode = oForm.DataSources.DBDataSources.Item("@BDOSINAC").GetValue("U_BankCode", 0);
                        if (!string.IsNullOrEmpty(bankCode))
                        {
                            SAPbouiCOM.Condition oCon = oCons.Add();
                            oCon.Alias = "U_BankCode";
                            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                            oCon.CondVal = bankCode;
                            oCFL.SetConditions(oCons);
                        }
                        else
                        {
                            bubbleEvent = false;
                            Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("BeforeChoosingCreditLineCodeFillBankFirst") + "!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                        }
                    }
                }
                else
                {
                    SAPbouiCOM.DataTable oDataTable = oCFLEvento.SelectedObjects;

                    if (oDataTable != null)
                    {
                        SAPbouiCOM.DBDataSource oDBDataSource = oForm.DataSources.DBDataSources.Item("@BDOSINAC");

                        if (oCFLEvento.ChooseFromListUID == "BankCFL")
                        {
                            string value = oDataTable.GetValue("BankCode", 0);
                            LanguageUtils.IgnoreErrors<string>(() => oForm.Items.Item("BankCodeE").Specific.Value = value);
                            if (bankCodeOld != value)
                                clearMatrix(oForm);
                        }
                        else if (oCFLEvento.ChooseFromListUID == "JournalEntryCFL")
                        {
                            string value = oDataTable.GetValue("TransId", 0).ToString();
                            LanguageUtils.IgnoreErrors<string>(() => oForm.Items.Item("TransIdE").Specific.Value = value);
                        }
                        else
                        {
                            string value;
                            switch (oCFLEvento.ChooseFromListUID)
                            {
                                case "CreditLineCodeCFL":
                                    value = oDataTable.GetValue("Code", 0);
                                    break;
                                case "CurrencyCFL":
                                    value = oDataTable.GetValue("CurrCode", 0);
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

                            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("LoanMTR").Specific;
                            LanguageUtils.IgnoreErrors<string>(() => oMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value = value);

                            if (oCFLEvento.ChooseFromListUID == "CreditLineCodeCFL")
                            {
                                updateRowByCreditLine(oForm, value, pVal.Row - 1);
                                updateExchangeRateRow(oForm, pVal.Row);
                            }
                        }

                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
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

        public static void setSizeForm(SAPbouiCOM.Form oForm)
        {
            oForm.Freeze(true);
            try
            {
                oForm.ClientHeight = Program.uiApp.Desktop.Height / 4;
                //oForm.Height = Program.uiApp.Desktop.Width / 4;
                oForm.Left = (Program.uiApp.Desktop.Width - oForm.Width) / 2;
                oForm.Top = (Program.uiApp.Desktop.Height - oForm.Height) / 3;
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

        public static void resizeForm(SAPbouiCOM.Form oForm)
        {
            oForm.Freeze(true);
            try
            {
                int left_e = 160;
                oForm.Items.Item("0_U_E").Left = left_e;
                oForm.Items.Item("0_U_E").Width = 140;

                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("LoanMTR").Specific;
                int mtrWidth = oForm.ClientWidth;
                oForm.Items.Item("LoanMTR").Width = mtrWidth;
                oForm.Items.Item("LoanMTR").Height = oForm.ClientHeight / 2;
                int columnsCount = oMatrix.Columns.Count - 1;
                oMatrix.Columns.Item("LineID").Width = 19;
                mtrWidth -= 19;
                mtrWidth /= columnsCount;
                foreach (SAPbouiCOM.Column column in oMatrix.Columns)
                {
                    if (column.UniqueID == "LineID")
                        continue;
                    column.Width = mtrWidth;
                }

                int height = 15;
                int top = oForm.Items.Item("LoanMTR").Top - height - 1;
                oForm.Items.Item("addMTRB").Top = top;
                oForm.Items.Item("delMTRB").Top = top;
                top = oForm.Items.Item("LoanMTR").Top + oForm.Items.Item("LoanMTR").Height + 10;
                oForm.Items.Item("CreatorS").Top = top;
                oForm.Items.Item("CreatorE").Top = top;
                top += height + 1;
                oForm.Items.Item("RemarksS").Top = top;
                oForm.Items.Item("RemarksE").Top = top;

                //ღილაკები
                int topTemp1 = oForm.Items.Item("RemarksE").Top + height * 2 + 1;
                int topTemp2 = oForm.ClientHeight - 25;
                //ღილაკები
                top = topTemp2 > topTemp1 ? topTemp2 : topTemp1;

                oForm.Items.Item("1").Top = top;
                oForm.Items.Item("2").Top = top;
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

        public static void setVisibleFormItems(SAPbouiCOM.Form oForm)
        {
            oForm.Freeze(true);
            try
            {
                //bool isFixed = oForm.DataSources.DBDataSources.Item("@BDOSINAC").GetValue("U_Fixed", 0) == "Y";
                //oForm.Items.Item("PerKmS").Visible = isFixed;
                //oForm.Items.Item("PerKmE").Visible = isFixed;
                //oForm.Items.Item("PerHrS").Visible = isFixed;
                //oForm.Items.Item("PerHrE").Visible = isFixed;
                //oForm.Items.Item("addMTRB").Visible = !isFixed;
                //oForm.Items.Item("delMTRB").Visible = !isFixed;
                //oForm.Items.Item("LoanMTR").Visible = !isFixed;
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

        public static void addMatrixRow(SAPbouiCOM.Form oForm)
        {
            oForm.Freeze(true);
            try
            {
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("LoanMTR").Specific;
                oMatrix.FlushToDataSource();

                SAPbouiCOM.DBDataSource oDBDataSourceMTR = oForm.DataSources.DBDataSources.Item("@BDOSINA1");
                if (!string.IsNullOrEmpty(oDBDataSourceMTR.GetValue("U_CrLnCode", oDBDataSourceMTR.Size - 1)))
                    oDBDataSourceMTR.InsertRecord(oDBDataSourceMTR.Size);
                oDBDataSourceMTR.SetValue("LineId", oDBDataSourceMTR.Size - 1, oDBDataSourceMTR.Size.ToString());

                oMatrix.LoadFromDataSource();

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
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

        public static void deleteMatrixRow(SAPbouiCOM.Form oForm)
        {
            oForm.Freeze(true);
            try
            {
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("LoanMTR").Specific;
                oMatrix.FlushToDataSource();
                int firstRow = 0;
                int row = 0;
                int deletedRowCount = 0;

                while (row != -1)
                {
                    row = oMatrix.GetNextSelectedRow(firstRow, SAPbouiCOM.BoOrderType.ot_RowOrder);
                    if (row > -1)
                    {
                        deletedRowCount++;
                        oForm.DataSources.DBDataSources.Item("@BDOSINA1").RemoveRecord(row - deletedRowCount);
                        firstRow = row;
                    }
                }

                SAPbouiCOM.DBDataSource oDBDataSourceMTR = oForm.DataSources.DBDataSources.Item("@BDOSINA1");
                int rowCount = oDBDataSourceMTR.Size;

                for (int i = 1; i <= rowCount; i++)
                {
                    string crLnCode = oDBDataSourceMTR.GetValue("U_CrLnCode", i - 1);
                    if (!string.IsNullOrEmpty(crLnCode))
                        oDBDataSourceMTR.SetValue("LineId", i - 1, i.ToString());
                }
                oMatrix.LoadFromDataSource();

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE && deletedRowCount > 0)
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                GC.Collect();
                oForm.Freeze(false);
            }
        }

        public static void clearMatrix(SAPbouiCOM.Form oForm)
        {
            oForm.Freeze(true);
            try
            {
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("LoanMTR").Specific;
                oMatrix.FlushToDataSource();
                SAPbouiCOM.DBDataSource oDBDataSourceMTR = oForm.DataSources.DBDataSources.Item("@BDOSINA1");
                oDBDataSourceMTR.Clear();
                oMatrix.LoadFromDataSource();
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

        private static void updateRowByCreditLine(SAPbouiCOM.Form oForm, string crLnCode, int i)
        {
            try
            {
                oForm.Freeze(true);

                SAPbouiCOM.DBDataSource oDBDataSource = oForm.DataSources.DBDataSources.Item("@BDOSINAC");
                string docEntryStr = oDBDataSource.GetValue("DocEntry", 0);

                int docEntry = 0;

                if (!string.IsNullOrEmpty(docEntryStr))
                    docEntry = Convert.ToInt32(oDBDataSource.GetValue("DocEntry", 0));

                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("LoanMTR").Specific;
                oMatrix.FlushToDataSource();

                SAPbouiCOM.DBDataSource oDBDataSourceMTR = oForm.DataSources.DBDataSources.Item("@BDOSINA1");

                SAPbobsCOM.Recordset oRecordSet = BDOSCreditLine.getInfo(crLnCode);
                var CreditLineObj = new
                {
                    Code = oRecordSet != null ? oRecordSet.Fields.Item("Code").Value : "",
                    U_CrLnName = oRecordSet != null ? oRecordSet.Fields.Item("Name").Value : "",
                    U_BankCode = oRecordSet != null ? oRecordSet.Fields.Item("U_BankCode").Value : "",
                    U_CurrCode = oRecordSet != null ? oRecordSet.Fields.Item("U_CurrCode").Value : "",
                    U_CrLnAcct = oRecordSet != null ? oRecordSet.Fields.Item("U_CrLnAcct").Value : "",
                    U_IntrstRate = oRecordSet != null ? Convert.ToDecimal(oRecordSet.Fields.Item("U_IntrstRate").Value, CultureInfo.InvariantCulture) : decimal.Zero,
                    U_StartDate = oRecordSet != null ? oRecordSet.Fields.Item("U_StartDate").Value : "",
                    U_ExpnsAcct = oRecordSet != null ? oRecordSet.Fields.Item("U_ExpnsAcct").Value : "",
                    U_IntPblAcct = oRecordSet != null ? oRecordSet.Fields.Item("U_IntPblAcct").Value : ""
                };

                if (oRecordSet != null)
                    Marshal.ReleaseComObject(oRecordSet);

                oDBDataSourceMTR.SetValue("U_CrLnName", i, CreditLineObj.U_CrLnName);
                oDBDataSourceMTR.SetValue("U_LnCurrCode", i, CreditLineObj.U_CurrCode);
                oDBDataSourceMTR.SetValue("U_CrLnAcct", i, CreditLineObj.U_CrLnAcct);
                oDBDataSourceMTR.SetValue("U_IntrstRate", i, FormsB1.ConvertDecimalToString(CreditLineObj.U_IntrstRate));
                oDBDataSourceMTR.SetValue("U_ExpnsAcct", i, CreditLineObj.U_ExpnsAcct);
                oDBDataSourceMTR.SetValue("U_IntPblAcct", i, CreditLineObj.U_IntPblAcct);

                oMatrix.LoadFromDataSource();
            }
            catch (Exception ex)
            {
                Program.uiApp.StatusBar.SetSystemMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        private static void updateExchangeRateRow(SAPbouiCOM.Form oForm, int rowIndex = 0)
        {
            try
            {
                oForm.Freeze(true);

                string docDateS = oForm.DataSources.DBDataSources.Item("@BDOSINAC").GetValue("U_DocDate", 0);
                if (string.IsNullOrEmpty(docDateS))
                    return;

                DateTime date = Convert.ToDateTime(DateTime.ParseExact(docDateS, "yyyyMMdd", CultureInfo.InvariantCulture));
                SAPbobsCOM.SBObob oSBOBob = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);

                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("LoanMTR").Specific;
                oMatrix.FlushToDataSource();

                int rowCount = rowIndex == 0 ? oMatrix.RowCount : rowIndex;
                int i = rowIndex == 0 ? 1 : rowIndex;

                SAPbouiCOM.DBDataSource oDBDataSourceMTR = oForm.DataSources.DBDataSources.Item("@BDOSINA1");

                for (; i <= rowCount; i++)
                {
                    string currency = oDBDataSourceMTR.GetValue("U_LnCurrCode", i - 1);
                    if (!string.IsNullOrEmpty(currency) && currency != Program.LocalCurrency)
                        oDBDataSourceMTR.SetValue("U_ExchngRate", i - 1, oSBOBob.GetCurrencyRate(currency, date).Fields.Item("CurrencyRate").Value);
                    else
                        oDBDataSourceMTR.SetValue("U_ExchngRate", i - 1, FormsB1.ConvertDecimalToString(decimal.Zero));
                }
                oMatrix.LoadFromDataSource();
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

        public static DataTable createAdditionalEntries(SAPbouiCOM.Form oForm, SAPbobsCOM.GeneralData oGeneralData)
        {
            try
            {
                DataTable jeLines = JournalEntry.JournalEntryTable();
                SAPbobsCOM.GeneralDataCollection oChild = null;
                SAPbouiCOM.DBDataSource oDBDataSourceTable = null;
                DataTable accountTable = CommonFunctions.GetOACTTable();

                int jeCount;

                if (oForm == null)
                {
                    oChild = oGeneralData.Child("BDOSINA1");
                    jeCount = oChild.Count;
                }
                else
                {
                    oDBDataSourceTable = oForm.DataSources.DBDataSources.Item("@BDOSINA1");
                    jeCount = oDBDataSourceTable.Size;
                }

                for (int i = 0; i < jeCount; i++)
                {
                    decimal interestAmountLC = Convert.ToDecimal(CommonFunctions.getChildOrDbDataSourceValue(oDBDataSourceTable, oChild, null, "U_IntAmtLC", i), CultureInfo.InvariantCulture);
                    decimal interestAmountFC = Convert.ToDecimal(CommonFunctions.getChildOrDbDataSourceValue(oDBDataSourceTable, oChild, null, "U_IntAmtFC", i), CultureInfo.InvariantCulture);

                    string expenseAccount = ((string)CommonFunctions.getChildOrDbDataSourceValue(oDBDataSourceTable, oChild, null, "U_ExpnsAcct", i)).Trim();
                    string interestPayableAccount = ((string)CommonFunctions.getChildOrDbDataSourceValue(oDBDataSourceTable, oChild, null, "U_IntPblAcct", i)).Trim();
                    string currencyCode = ((string)CommonFunctions.getChildOrDbDataSourceValue(oDBDataSourceTable, oChild, null, "U_LnCurrCode", i)).Trim();

                    if (currencyCode == Program.LocalCurrency)
                    {
                        currencyCode = string.Empty;
                        interestAmountFC = decimal.Zero;
                    }

                    JournalEntry.AddJournalEntryRow(accountTable, jeLines, "Full", expenseAccount, interestPayableAccount, interestAmountLC, interestAmountFC, currencyCode, "", "", "", "", "", "", "", "");
                }
                return jeLines;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        public static int jrnEntry(string docEntry, string docNum, DateTime docDate, DataTable jrnLinesDT, out string errorText)
        {
            int transId = 0;
            try
            {
                JournalEntry.JrnEntry(docEntry, "UDO_F_BDOSINAC_D", "Interest Accrual Document: " + docNum, docDate, jrnLinesDT, out errorText);

                if (string.IsNullOrEmpty(errorText))
                {
                    string Ref1 = docEntry.ToString();
                    string Ref2 = "UDO_F_BDOSINAC_D";

                    SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    string query = "SELECT " +
                                    "\"TransId\" " +
                                    "FROM \"OJDT\"  " +
                                    "WHERE \"StornoToTr\" IS NULL " +
                                    "AND \"Ref1\" = '" + Ref1 + "' " +
                                    "AND \"Ref2\" = '" + Ref2 + "' ";
                    oRecordSet.DoQuery(query);

                    if (!oRecordSet.EoF)
                        transId = Convert.ToInt32(oRecordSet.Fields.Item("TransId").Value);
                    Marshal.ReleaseComObject(oRecordSet);
                }
                return transId;
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
                return transId;
            }
        }

        public static void cancellation(SAPbouiCOM.Form oForm, int docEntry)
        {
            string errorText;
            JournalEntry.cancellation(oForm, docEntry, "UDO_F_BDOSINAC_D", out errorText);
            Program.canceledDocEntry = 0;
            if (!string.IsNullOrEmpty(errorText))
            {
                throw new Exception(errorText);
            }
        }
    }
}
