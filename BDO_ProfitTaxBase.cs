﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace BDO_Localisation_AddOn
{
    static partial class BDO_ProfitTaxBase
    {
        public static void createMasterDataUDO( out string errorText)
        {
            errorText = null;
            string tableName = "BDO_PTBS";
            string description = "Profit Tax Base";

            SAPbobsCOM.UserObjectsMD oUserObjectsMD = null;
            oUserObjectsMD = (SAPbobsCOM.UserObjectsMD)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectsMD);

            int result = UDO.addUserTable( tableName, description, SAPbobsCOM.BoUTBTableType.bott_MasterData, out errorText);

            if (result != 0)
            {
                return;
            }

            Dictionary<string, object> fieldskeysMap;

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BaseType");
            fieldskeysMap.Add("TableName", "BDO_PTBS");
            fieldskeysMap.Add("Description", BDOSResources.getTranslate("BaseType") + " " + BDOSResources.getTranslate("Code"));
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);
            fieldskeysMap.Add("Mandatory", SAPbobsCOM.BoYesNoEnum.tYES);
            fieldskeysMap.Add("DefaultValue", "1");

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BsTpDscr");
            fieldskeysMap.Add("TableName", "BDO_PTBS");
            fieldskeysMap.Add("Description", BDOSResources.getTranslate("BaseType") + " " + BDOSResources.getTranslate("Name"));
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 100);
            fieldskeysMap.Add("Mandatory", SAPbobsCOM.BoYesNoEnum.tYES);
            fieldskeysMap.Add("DefaultValue", BDOSResources.getTranslate("ProfitSharing"));

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            GC.Collect();
        }

        public static void registerUDO( out string errorText)
        {
            errorText = null;
            string code = "UDO_F_BDO_PTBS_D"; //20 characters (must include at least one alphabetical character).
            Dictionary<string, object> formProperties;

            formProperties = new Dictionary<string, object>();
            formProperties.Add("Name", "Profit Tax Base"); //100 characters
            formProperties.Add("TableName", "BDO_PTBS");
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
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "U_BaseType");
            fieldskeysMap.Add("ColumnDescription", "BaseType" + " " + "Code"); //
            listFindColumns.Add(fieldskeysMap);
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "U_BsTpDscr");
            fieldskeysMap.Add("ColumnDescription", "BaseType" + " " + "Name"); //
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
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Editable", SAPbobsCOM.BoYesNoEnum.tYES);
            fieldskeysMap.Add("FormColumnAlias", "U_BaseType");
            fieldskeysMap.Add("FormColumnDescription", "BaseType" + " " + "Code"); //
            fieldskeysMap.Add("SonNumber", 0); //30 characters
            listFormColumns.Add(fieldskeysMap);
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Editable", SAPbobsCOM.BoYesNoEnum.tYES);
            fieldskeysMap.Add("FormColumnAlias", "U_BsTpDscr");
            fieldskeysMap.Add("FormColumnDescription", "BaseType" + " " + "Name"); //
            fieldskeysMap.Add("SonNumber", 0); //30 characters
            listFormColumns.Add(fieldskeysMap);

            formProperties.Add("FormColumns", listFormColumns);

            UDO.registerUDO( code, formProperties, out errorText);

            GC.Collect();
        }

        public static void addMenus( out string errorText)
        {
            errorText = null;

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
                oCreationPackage.UniqueID = "UDO_F_BDO_PTBS_D";
                oCreationPackage.String = BDOSResources.getTranslate("ProfitTaxBaseMasterData");
                oCreationPackage.Position = -1;

                menuItem = fatherMenuItem.SubMenus.AddEx(oCreationPackage);
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
        }

        public static void createFormItems( SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            GC.Collect();
        }

        public static void setSizeForm( SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            try
            {
                oForm.ClientHeight = Program.uiApp.Desktop.Width / 6;

                oForm.Height = Program.uiApp.Desktop.Width / 4;

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

        public static void chooseFromList(SAPbouiCOM.Form oForm, SAPbouiCOM.IChooseFromListEvent oCFLEvento, out string errorText)
        {
            errorText = null;

            SAPbouiCOM.DataTable oDataTableSelectedObjects = oCFLEvento.SelectedObjects;
            string BsTypeCode = Convert.ToString(oDataTableSelectedObjects.GetValue("Code", 0));
            string BsTypeName = Convert.ToString(oDataTableSelectedObjects.GetValue("Name", 0));

            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("3").Specific));

            SAPbouiCOM.EditText BsTypeCodeEdit = oMatrix.Columns.Item("U_BaseType").Cells.Item(oCFLEvento.Row).Specific;

            SAPbouiCOM.Column oColumnName = oMatrix.Columns.Item("U_BsTpDscr");
            oColumnName.Editable = true;
            SAPbouiCOM.EditText BsTypeNameEdit = oColumnName.Cells.Item(oCFLEvento.Row).Specific;

            try
            {   

                oMatrix.Columns.Item("U_BsTpDscr").Cells.Item(oCFLEvento.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                BsTypeCodeEdit.Value = BsTypeCode;
                
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                }

            }
            catch (Exception ex)
            {
                errorText = ex.Message;
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                }
            }
            finally
            {
                GC.Collect();
            }

            try
            {
               
                oMatrix.Columns.Item("U_BaseType").Cells.Item(oCFLEvento.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                BsTypeNameEdit.Value = BsTypeName;

                oMatrix.Columns.Item("U_BaseType").Cells.Item(oCFLEvento.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                oColumnName.Editable = false;

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                }

            }
            catch (Exception ex)
            {
                errorText = ex.Message;
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                }
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
                    MatrixLink( oForm, out errorText);

                    createFormItems( oForm, out errorText);
                    Program.FORM_LOAD_FOR_VISIBLE = true;

                }
                
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                {
                    if (pVal.BeforeAction == false)
                    {
                        SAPbouiCOM.IChooseFromListEvent oCFLEvento = null;
                        oCFLEvento = ((SAPbouiCOM.IChooseFromListEvent)(pVal));

                        chooseFromList(oForm, oCFLEvento, out errorText);
                    }
                }
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE & pVal.BeforeAction == false & Program.FORM_LOAD_FOR_VISIBLE == true)
                {
                    oForm.Freeze(true);
                    setSizeForm( oForm, out errorText);
                    oForm.Title = BDOSResources.getTranslate("ProfitTaxBaseMasterData");
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

        public static void MatrixLink(  SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            bool multiSelection = false;
            string objectType = "UDO_F_BDO_PTBT_D"; 
            FormsB1.addChooseFromList( oForm, multiSelection, objectType, "CFL_BaseType");
            
            SAPbouiCOM.ChooseFromList oCFL;

            oCFL = oForm.ChooseFromLists.Item("CFL_BaseType");

            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("3").Specific));

            SAPbouiCOM.Column oColumn;

            oColumn = oMatrix.Columns.Item("U_BaseType");
            oColumn.ChooseFromListUID = "CFL_BaseType";
            oColumn.ChooseFromListAlias = "Code";

            oColumn = oMatrix.Columns.Item("U_BsTpDscr");
            oColumn.Editable = false;
        }

        public static bool checkRemoving( SAPbouiCOM.Form oForm, int removeRecordRow, out string errorText)
        {
            errorText = null;

                        SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("3").Specific));
            SAPbouiCOM.CellPosition cellPos = oMatrix.GetCellFocus();
            SAPbouiCOM.EditText oEditText = oMatrix.Columns.Item("Code").Cells.Item(removeRecordRow).Specific;
            string code = oEditText.Value.Trim();

            Dictionary<string, string> listTables = new Dictionary<string, string>();
            listTables.Add("@BDO_TAXP", "U_prBase"); //Profit Tax Accrual
            listTables.Add("OIGE", "U_prBase"); //Goods Isshue
            listTables.Add("OPCH", "U_prBase"); //AP Invoice
            listTables.Add("ODPO", "U_prBase"); // AP Down Payment Request
            listTables.Add("OVPM", "U_prBase"); // Outgoing Payment

            return CommonFunctions.codeIsUsed( listTables, code);
        }

    }
}
