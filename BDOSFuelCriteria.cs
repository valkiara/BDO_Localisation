using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BDO_Localisation_AddOn
{
    class BDOSFuelCriteria
    {
        public static void createMasterDataUDO(out string errorText)
        {
            string tableName = "BDOSFUCR";
            string description = "Fuel Criteria";

            int result = UDO.addUserTable(tableName, description, SAPbobsCOM.BoUTBTableType.bott_MasterData, out errorText);

            if (result != 0)
            {
                return;
            }

            Dictionary<string, object> fieldskeysMap;

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "Value");
            fieldskeysMap.Add("TableName", "BDOSFUCR");
            fieldskeysMap.Add("Description", "Value");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 100);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "Percentage");
            fieldskeysMap.Add("TableName", "BDOSFUCR");
            fieldskeysMap.Add("Description", "Percentage");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Percentage);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            GC.Collect();
        }

        public static void registerUDO()
        {
            SAPbobsCOM.UserObjectsMD oUserObjectMD = null;
            SAPbobsCOM.UserObjectMD_FindColumns oUDOFind = null;
            SAPbobsCOM.UserObjectMD_FormColumns oUDOForm = null;
            SAPbobsCOM.UserObjectMD_EnhancedFormColumns oUDOEnhancedForm = null;
            GC.Collect();
            oUserObjectMD = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD) as SAPbobsCOM.UserObjectsMD;
            oUDOFind = oUserObjectMD.FindColumns;
            oUDOForm = oUserObjectMD.FormColumns;
            oUDOEnhancedForm = oUserObjectMD.EnhancedFormColumns;

            var retval = oUserObjectMD.GetByKey("UDO_F_BDOSFUCR_D");

            if (!retval)
            {
                oUserObjectMD.Code = "UDO_F_BDOSFUCR_D";
                oUserObjectMD.Name = "Fuel Criteria";
                oUserObjectMD.TableName = "BDOSFUCR";
                oUserObjectMD.ObjectType = SAPbobsCOM.BoUDOObjType.boud_MasterData;
                oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO;
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
                oUDOFind.ColumnDescription = "Name";
                oUDOFind.Add();
                oUDOFind.ColumnAlias = "U_Value";
                oUDOFind.ColumnDescription = "Value";
                oUDOFind.Add();
                oUDOFind.ColumnAlias = "U_Percentage";
                oUDOFind.ColumnDescription = "Percentage";
                oUDOFind.Add();

                //Form
                oUDOForm.FormColumnAlias = "Code";
                oUDOForm.FormColumnDescription = "Code";
                oUDOForm.Editable = SAPbobsCOM.BoYesNoEnum.tYES;
                oUDOForm.Add();

                oUDOForm.FormColumnAlias = "Name";
                oUDOForm.FormColumnDescription = "Name";
                oUDOForm.Editable = SAPbobsCOM.BoYesNoEnum.tYES;
                oUDOForm.Add();

                oUDOForm.FormColumnAlias = "U_Value";
                oUDOForm.FormColumnDescription = "Value";
                oUDOForm.Editable = SAPbobsCOM.BoYesNoEnum.tYES;
                oUDOForm.Add();

                oUDOForm.FormColumnAlias = "U_Percentage";
                oUDOForm.FormColumnDescription = "Percentage";
                oUDOForm.Editable = SAPbobsCOM.BoYesNoEnum.tYES;
                oUDOForm.Add();

                //Enhanced
                oUDOEnhancedForm.ColumnAlias = "Code";
                oUDOEnhancedForm.ColumnDescription = "Code";
                oUDOEnhancedForm.ColumnIsUsed = SAPbobsCOM.BoYesNoEnum.tYES;
                oUDOEnhancedForm.ColumnNumber = 1;
                oUDOEnhancedForm.ChildNumber = 1;
                oUDOEnhancedForm.Editable = SAPbobsCOM.BoYesNoEnum.tYES;
                oUDOEnhancedForm.Add();

                oUDOEnhancedForm.ColumnAlias = "U_Value";
                oUDOEnhancedForm.ColumnDescription = "Value";
                oUDOEnhancedForm.ColumnIsUsed = SAPbobsCOM.BoYesNoEnum.tYES;
                oUDOEnhancedForm.ColumnNumber = 2;
                oUDOEnhancedForm.ChildNumber = 1;
                oUDOEnhancedForm.Editable = SAPbobsCOM.BoYesNoEnum.tYES;
                oUDOEnhancedForm.Add();

                oUDOEnhancedForm.ColumnAlias = "U_Percentage";
                oUDOEnhancedForm.ColumnDescription = "Percentage";
                oUDOEnhancedForm.ColumnIsUsed = SAPbobsCOM.BoYesNoEnum.tYES;
                oUDOEnhancedForm.ColumnNumber = 3;
                oUDOEnhancedForm.ChildNumber = 1;
                oUDOEnhancedForm.Editable = SAPbobsCOM.BoYesNoEnum.tYES;
                oUDOEnhancedForm.Add();

                if (!retval)
                {
                    if ((oUserObjectMD.Add() != 0))
                    {
                        Program.uiApp.MessageBox(Program.oCompany.GetLastErrorDescription());
                    }
                }
            }
        }

        public static void addMenus()
        {
            string enableFuelMng = (string)CommonFunctions.getOADM("U_BDOSEnbFlM");

            if (enableFuelMng == "Y")
            {
                try
                {
                    SAPbouiCOM.MenuItem fatherMenuItem = Program.uiApp.Menus.Item("3072");
                    // Add a pop-up menu item
                    SAPbouiCOM.MenuCreationParams oCreationPackage = (SAPbouiCOM.MenuCreationParams)Program.uiApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                    oCreationPackage.Checked = false;
                    oCreationPackage.Enabled = true;
                    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                    oCreationPackage.UniqueID = "UDO_F_BDOSFUCR_D";
                    oCreationPackage.String = BDOSResources.getTranslate("FuelCriteria");
                    oCreationPackage.Position = fatherMenuItem.SubMenus.Count - 1;

                    SAPbouiCOM.MenuItem menuItem = fatherMenuItem.SubMenus.AddEx(oCreationPackage);
                }
                catch
                {
                    //Program.uiApp.MessageBox(ex.Message);
                }
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
                    Program.FORM_LOAD_FOR_ACTIVATE = true;
                    changeFormItems(oForm);
                    oForm.Title = BDOSResources.getTranslate("FuelCriteria");
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE && !pVal.BeforeAction)
                {
                    if (Program.FORM_LOAD_FOR_ACTIVATE)
                    {
                        Program.FORM_LOAD_FOR_ACTIVATE = false;
                        setVisibleFormItems(oForm);
                    }
                }
            }
        }

        public static void changeFormItems(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("3").Specific));

            SAPbouiCOM.Columns oColumns = oMatrix.Columns;

            SAPbouiCOM.Column oColumn = oColumns.Item("Code");
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Code");
            oColumn = oColumns.Item("Name");
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Name");
            oColumn = oColumns.Item("U_Value");
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Value");
            oColumn = oColumns.Item("U_Percentage");
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Percentage");
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
    }
}
