﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace BDO_Localisation_AddOn
{
    static partial class BDO_ProfitTaxBaseType
    {
        public static void createMasterDataUDO( out string errorText)
        {
            errorText = null;
            string tableName = "BDO_PTBT";
            string description = "Profit Tax Base Type";

            SAPbobsCOM.UserObjectsMD oUserObjectsMD = null;
            oUserObjectsMD = (SAPbobsCOM.UserObjectsMD)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectsMD);

            int result = UDO.addUserTable( tableName, description, SAPbobsCOM.BoUTBTableType.bott_MasterData, out errorText);

            if (result != 0)
            {
                return;
            }

            GC.Collect();
        }

        public static void CreateBaseTypes( out string errorText)
        {
            errorText = null;

            SAPbobsCOM.CompanyService oCompanyService = null;
            SAPbobsCOM.GeneralService oGeneralService = null;
            SAPbobsCOM.GeneralData oGeneralData = null;
            oCompanyService = Program.oCompany.GetCompanyService();
            oGeneralService = oCompanyService.GetGeneralService("UDO_F_BDO_PTBT_D");
            oGeneralData = ((SAPbobsCOM.GeneralData)(oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)));

            oGeneralData.SetProperty("Code", "1");
            oGeneralData.SetProperty("Name", BDOSResources.getTranslate("ProfitSharing"));
            try
            {
                var response = oGeneralService.Add(oGeneralData);
            }
            catch
            {
            }

            oGeneralData.SetProperty("Code", "2");
            oGeneralData.SetProperty("Name", BDOSResources.getTranslate("ExpensesOfNonEconomicExpenses"));
            try
            {
                var response = oGeneralService.Add(oGeneralData);
            }
            catch
            {
            }

            oGeneralData.SetProperty("Code", "3");
            oGeneralData.SetProperty("Name", BDOSResources.getTranslate("PaymentsForNonEconomicActivities"));
            try
            {
                var response = oGeneralService.Add(oGeneralData);
            }
            catch
            {
            }

            oGeneralData.SetProperty("Code", "4");
            oGeneralData.SetProperty("Name", BDOSResources.getTranslate("GoodsAndServicesDeliveredFreeOfCharge"));
            try
            {
                var response = oGeneralService.Add(oGeneralData);
            }
            catch
            {
            }

            oGeneralData.SetProperty("Code", "5");
            oGeneralData.SetProperty("Name", BDOSResources.getTranslate("RepresentationalExpensesAbovePredefinedLevel"));
            try
            {
                var response = oGeneralService.Add(oGeneralData);
            }
            catch
            {
            }

            oGeneralData.SetProperty("Code", "6");
            oGeneralData.SetProperty("Name", BDOSResources.getTranslate("ProfitTaxExempt"));
            try
            {
                var response = oGeneralService.Add(oGeneralData);
            }
            catch
            {
            }
        }

        public static void registerUDO( out string errorText)
        {
            errorText = null;
            string code = "UDO_F_BDO_PTBT_D"; //20 characters (must include at least one alphabetical character).
            Dictionary<string, object> formProperties;

            formProperties = new Dictionary<string, object>();
            formProperties.Add("Name", "Profit Tax Base Type"); //100 characters
            formProperties.Add("TableName", "BDO_PTBT");
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
            fieldskeysMap.Add("FormColumnAlias", "Name");
            fieldskeysMap.Add("FormColumnDescription", "Name"); //30 characters
            fieldskeysMap.Add("SonNumber", 0); //30 characters
            listFormColumns.Add(fieldskeysMap);

            formProperties.Add("FormColumns", listFormColumns);

            UDO.registerUDO( code, formProperties, out errorText);
                       
            GC.Collect();

            try
            {
                CreateBaseTypes( out errorText);
            }
            catch (Exception ex)
            {
                string exM = ex.Message;
            }
        }

        public static void addMenus()
        {
            SAPbouiCOM.MenuItem menuItem;
            SAPbouiCOM.MenuItem fatherMenuItem;
            SAPbouiCOM.MenuCreationParams oCreationPackage;

            try
            {
                fatherMenuItem = Program.uiApp.Menus.Item("15616");
                // Add a pop-up menu item
                oCreationPackage = (SAPbouiCOM.MenuCreationParams)Program.uiApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                oCreationPackage.Checked = false;
                oCreationPackage.Enabled = true;
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "UDO_F_BDO_PTBT_D";
                oCreationPackage.String = BDOSResources.getTranslate("ProfitTaxBaseTypeMasterData");
                oCreationPackage.Position = -1;

                menuItem = fatherMenuItem.SubMenus.AddEx(oCreationPackage);
            }
            catch 
            {

            }
        }

        public static void createFormItems(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            GC.Collect();
        }

        public static void setVisibleFormItems(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
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

        public static void uiApp_ItemEvent(  string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
            SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD & pVal.BeforeAction == false)
                {
                    createFormItems(oForm, out errorText);
                    setVisibleFormItems(oForm, out errorText);
                    Program.FORM_LOAD_FOR_VISIBLE = true;
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE & pVal.BeforeAction == false & Program.FORM_LOAD_FOR_VISIBLE == true)
                {
                    oForm.Freeze(true);
                    setSizeForm( oForm, out errorText);
                    oForm.Title = BDOSResources.getTranslate("ProfitTaxBaseTypeMasterData");
                    oForm.Freeze(false);
                    Program.FORM_LOAD_FOR_VISIBLE = false;
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_RESIZE & pVal.BeforeAction == false)
                {
                    oForm.Freeze(true);
                    resizeForm( oForm, out errorText);
                    oForm.Freeze(false);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_VALIDATE & pVal.BeforeAction == true & pVal.InnerEvent == true)
                {
                    if (Program.removeRecordRow != 0 & Program.removeRecordTrans == true)
                    {
                        if (checkRemoving( oForm, Program.removeRecordRow, out errorText) == true)
                        {
                            Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("RecordIsUsedInDocuments"));
                            Program.uiApp.MessageBox(BDOSResources.getTranslate("OperationUnsuccesfullSeeLog"));
                            BubbleEvent = false;
                        }
                        Program.removeRecordRow = 0;
                        Program.removeRecordTrans = false;
                    }
                }
            }
        }

        public static bool checkRemoving( SAPbouiCOM.Form oForm, int removeRecordRow, out string errorText)
        {
            errorText = null;

            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("3").Specific));
            SAPbouiCOM.CellPosition cellPos = oMatrix.GetCellFocus();
            SAPbouiCOM.EditText oEditText = oMatrix.Columns.Item("Code").Cells.Item(removeRecordRow).Specific;
            string code = oEditText.Value.Trim();

            Dictionary<string, string> listTables = new Dictionary<string, string>();
            listTables.Add("@BDO_PTBS", "U_BaseType"); //Profit Tax Base

            return CommonFunctions.codeIsUsed( listTables, code);
        }
    }
}
