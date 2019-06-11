using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BDO_Localisation_AddOn
{
    class BDOSItemCategories
    {
        public static void createMasterDataUDO( out string errorText)
        {
            //Item Categories
            string tableName = "BDOSITMCTG";
            string description = "Item Categories";

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
            fieldskeysMap.Add("Name", "FatherID");
            fieldskeysMap.Add("TableName", "BDOSITMCTG");
            fieldskeysMap.Add("Description", "Father ID");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("Mandatory", SAPbobsCOM.BoYesNoEnum.tNO);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "FatherN");
            fieldskeysMap.Add("TableName", "BDOSITMCTG");
            fieldskeysMap.Add("Description", "Father Name");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("Mandatory", SAPbobsCOM.BoYesNoEnum.tNO);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "Level");
            fieldskeysMap.Add("TableName", "BDOSITMCTG");
            fieldskeysMap.Add("Description", "Level");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("Mandatory", SAPbobsCOM.BoYesNoEnum.tYES);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            GC.Collect();

        }

        public static void registerUDO( out string errorText)
        {
            errorText = null;
            string code = "UDO_F_BDOSITMCTG_D"; //20 characters (must include at least one alphabetical character).
            Dictionary<string, object> formProperties;

            formProperties = new Dictionary<string, object>();
            formProperties.Add("Name", "Item Categories"); //100 characters
            formProperties.Add("TableName", "BDOSITMCTG");
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
            fieldskeysMap.Add("ColumnAlias", "U_FatherID");
            fieldskeysMap.Add("ColumnDescription", "Father ID"); //
            listFindColumns.Add(fieldskeysMap);
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "U_FatherN");
            fieldskeysMap.Add("ColumnDescription", "Father Name"); //
            listFindColumns.Add(fieldskeysMap);
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "U_Level");
            fieldskeysMap.Add("ColumnDescription", "Level"); //
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
            fieldskeysMap.Add("FormColumnAlias", "U_FatherID");
            fieldskeysMap.Add("FormColumnDescription", "Father ID"); //
            fieldskeysMap.Add("SonNumber", 0); //30 characters
            listFormColumns.Add(fieldskeysMap);
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Editable", SAPbobsCOM.BoYesNoEnum.tYES);
            fieldskeysMap.Add("FormColumnAlias", "U_FatherN");
            fieldskeysMap.Add("FormColumnDescription", "Father Name"); //
            fieldskeysMap.Add("SonNumber", 0); //30 characters
            listFormColumns.Add(fieldskeysMap);
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Editable", SAPbobsCOM.BoYesNoEnum.tYES);
            fieldskeysMap.Add("FormColumnAlias", "U_Level");
            fieldskeysMap.Add("FormColumnDescription", "Level"); //
            fieldskeysMap.Add("SonNumber", 0); //30 characters
            listFormColumns.Add(fieldskeysMap);

            formProperties.Add("FormColumns", listFormColumns);

            UDO.registerUDO( code, formProperties, out errorText);

            GC.Collect();
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
                    changeFormItems( oForm, out errorText);
                    oForm.Title = BDOSResources.getTranslate("ItemCategoriesMasterData");
                    Program.FORM_LOAD_FOR_VISIBLE = true;

                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                {                   
                    SAPbouiCOM.ChooseFromListEvent oCFLEvento = (SAPbouiCOM.ChooseFromListEvent)pVal;
                    chooseFromList( oForm, oCFLEvento, pVal.ItemUID, pVal.BeforeAction, pVal.Row, out errorText);   
                    if (pVal.Before_Action)
                    {
                        if (string.IsNullOrEmpty(errorText) == false)
                        {
                            Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("LevelIsNotFill"));
                            Program.uiApp.MessageBox(BDOSResources.getTranslate("OperationUnsuccesfullSeeLog"));
                            BubbleEvent = false;
                        }
                    }
                }

                if ((pVal.ColUID == "U_FatherID" || pVal.ColUID == "U_Level") && pVal.EventType == SAPbouiCOM.BoEventTypes.et_KEY_DOWN && pVal.BeforeAction == false)
                {
                    oForm.Freeze(true);
                    clearItems( oForm, pVal.Row, out errorText);
                    oForm.Freeze(false);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE & pVal.BeforeAction == false & Program.FORM_LOAD_FOR_VISIBLE == true)
                {
                    oForm.Freeze(true);
                    oForm.Title = BDOSResources.getTranslate("ItemCategoriesMasterData");
                    oForm.Freeze(false);
                    Program.FORM_LOAD_FOR_VISIBLE = false;
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

        public static void changeFormItems( SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("3").Specific;

            SAPbouiCOM.Column oColumn = oMatrix.Columns.Item("Code");
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Code");

            oColumn = oMatrix.Columns.Item("Name");
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Name");

            oColumn = oMatrix.Columns.Item("U_FatherID");
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Father") + " " + BDOSResources.getTranslate("ID");

            oColumn = oMatrix.Columns.Item("U_FatherN");
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Father") + " " + BDOSResources.getTranslate("Name");

            oColumn = oMatrix.Columns.Item("U_Level");
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Level");

            GC.Collect();
        }

        public static void MatrixLink(  SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            //UOM CODE
            SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams;
            SAPbouiCOM.ChooseFromList oCFL;
            SAPbouiCOM.ChooseFromListCollection oCFLs = oForm.ChooseFromLists;
            oCFLCreationParams = Program.uiApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);

            oCFLCreationParams.MultiSelection = false;
            oCFLCreationParams.ObjectType = "UDO_F_BDOSITMCTG_D";
            oCFLCreationParams.UniqueID = "CFL_Father";
            oCFL = oCFLs.Add(oCFLCreationParams);

            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("3").Specific));

            SAPbouiCOM.Column  oColumn = oMatrix.Columns.Item("U_FatherID");
            oColumn.ChooseFromListUID = "CFL_Father";
            oColumn.ChooseFromListAlias = "Code";

            oColumn = oMatrix.Columns.Item("U_FatherN");
            oColumn.Editable = false;
        }

        public static void chooseFromList( SAPbouiCOM.Form oForm, SAPbouiCOM.ChooseFromListEvent oCFLEvento, string ItemUID, bool BeforeAction, int row, out string errorText)
        {
            errorText = null;

            if (BeforeAction == false)
            {
                SAPbouiCOM.DataTable oDataTableSelectedObjects = oCFLEvento.SelectedObjects;
                string stCode = Convert.ToString(oDataTableSelectedObjects.GetValue("Code", 0));
                string stName = Convert.ToString(oDataTableSelectedObjects.GetValue("Name", 0));

                SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("3").Specific));

                SAPbouiCOM.EditText CodeEdit = oMatrix.Columns.Item("U_FatherID").Cells.Item(row).Specific;

                SAPbouiCOM.Column oColumnName = oMatrix.Columns.Item("U_FatherN");
                oColumnName.Editable = true;
                SAPbouiCOM.EditText NameEdit = oColumnName.Cells.Item(oCFLEvento.Row).Specific;

                try
                {
                    oMatrix.Columns.Item("U_FatherN").Cells.Item(oCFLEvento.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    CodeEdit.Value = stCode;
                }
                catch (Exception ex)
                {
                    errorText = ex.Message;
                }

                try
                {
                    oMatrix.Columns.Item("U_FatherID").Cells.Item(row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    NameEdit.Value = stName;

                    oMatrix.Columns.Item("U_FatherID").Cells.Item(row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    oColumnName.Editable = false;
                }
                catch (Exception ex)
                {
                    errorText = ex.Message;
                }

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                }
                GC.Collect();
            }
            else
            {                
                string sCFL_ID = oCFLEvento.ChooseFromListUID;
                SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(sCFL_ID);
                SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();

                SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("3").Specific;
                SAPbouiCOM.EditText oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item("U_Level").Cells.Item(row).Specific;
                if (string.IsNullOrEmpty(oEditText.Value))
                {
                    errorText = "LevelIsNotFill";
                }
                else
                {
                    int level = Convert.ToInt32(oEditText.Value) - 1;
                    SAPbouiCOM.Condition oCon = oCons.Add();
                    oCon.Alias = "U_Level";
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCon.CondVal = level.ToString();
                    oCon.Relationship =  SAPbouiCOM.BoConditionRelationship.cr_NONE;

                    oCFL.SetConditions(oCons);
                }
            }
        }

        public static void clearItems( SAPbouiCOM.Form oForm, int row, out string errorText)
        {
            errorText = null;

            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("3").Specific));

            SAPbouiCOM.EditText CodeEdit = oMatrix.Columns.Item("U_FatherID").Cells.Item(row).Specific;
            SAPbouiCOM.EditText LevelEdit = oMatrix.Columns.Item("U_Level").Cells.Item(row).Specific;

            SAPbouiCOM.Column oColumnName = oMatrix.Columns.Item("U_FatherN");
            SAPbouiCOM.EditText NameEdit = oColumnName.Cells.Item(row).Specific;

            if (string.IsNullOrEmpty(LevelEdit.Value) && string.IsNullOrEmpty(CodeEdit.Value) == false ||
                string.IsNullOrEmpty(CodeEdit.Value) && string.IsNullOrEmpty(NameEdit.Value) == false)
            {
                oColumnName.Editable = true;

                string stCode = CodeEdit.Value;

                if (string.IsNullOrEmpty(LevelEdit.Value))
                {
                    stCode = "";
                    try
                    {
                        oMatrix.Columns.Item("U_FatherN").Cells.Item(row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        CodeEdit.Value = "";
                    }
                    catch (Exception ex)
                    {
                        errorText = ex.Message;
                    }
                }

                if (string.IsNullOrEmpty(stCode))
                {
                    try
                    {
                        oMatrix.Columns.Item("U_FatherID").Cells.Item(row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        NameEdit.Value = "";

                        oMatrix.Columns.Item("U_FatherID").Cells.Item(row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        oColumnName.Editable = false;
                    }
                    catch (Exception ex)
                    {
                        errorText = ex.Message;
                    }
                }

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                }
                GC.Collect();
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
            for (int i = 1; i <= 10; i++)
            {
                listTables.Add("OITM", "U_BDOSCtg" + i.ToString());
            }
            return CommonFunctions.codeIsUsed( listTables, code);
        }

        public static void addMenus( out string errorText)
        {
            errorText = null;

            SAPbouiCOM.MenuItem menuItem;
            SAPbouiCOM.MenuItem fatherMenuItem;
            SAPbouiCOM.MenuCreationParams oCreationPackage;

            try
            {
                fatherMenuItem = Program.uiApp.Menus.Item("15872");
                // Add a pop-up menu item
                oCreationPackage = (SAPbouiCOM.MenuCreationParams)Program.uiApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                oCreationPackage.Checked = false;
                oCreationPackage.Enabled = true;
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "UDO_F_BDOSITMCTG_D";
                oCreationPackage.String = BDOSResources.getTranslate("ItemCategories");
                oCreationPackage.Position = -1;

                menuItem = fatherMenuItem.SubMenus.AddEx(oCreationPackage);
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }

        }

    }
}
