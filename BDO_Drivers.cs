using System;
using System.Collections.Generic;

namespace BDO_Localisation_AddOn
{
    static partial class BDO_Drivers
    {
        public static void createMasterDataUDO(out string errorText)
        {
            errorText = null;
            string tableName = "BDO_DRVS";
            string description = "Drivers";

            SAPbobsCOM.UserObjectsMD oUserObjectsMD = null;
            oUserObjectsMD = (SAPbobsCOM.UserObjectsMD)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectsMD);

            int result = UDO.addUserTable(tableName, description, SAPbobsCOM.BoUTBTableType.bott_MasterData, out errorText);

            if (result != 0)
            {
                return;
            }

            Dictionary<string, object> fieldskeysMap;

            fieldskeysMap = new Dictionary<string, object>
            {
                { "Name", "firstName" },
                { "TableName", "BDO_DRVS" },
                { "Description", "First Name" },
                { "Type", SAPbobsCOM.BoFieldTypes.db_Alpha },
                { "EditSize", 50 }
            };

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>
            {
                { "Name", "lastName" },
                { "TableName", "BDO_DRVS" },
                { "Description", "Last Name" },
                { "Type", SAPbobsCOM.BoFieldTypes.db_Alpha },
                { "EditSize", 50 }
            };

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "tin");
            fieldskeysMap.Add("TableName", "BDO_DRVS");
            fieldskeysMap.Add("Description", "Driver TIN");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 11);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "notRsdnt");
            fieldskeysMap.Add("TableName", "BDO_DRVS");
            fieldskeysMap.Add("Description", "Driver Not Resident");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 1);
            fieldskeysMap.Add("DefaultValue", "N");

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            GC.Collect();
        }

        public static void registerUDO(out string errorText)
        {
            errorText = null;
            string code = "UDO_F_BDO_DRVS_D"; //20 characters (must include at least one alphabetical character).
            Dictionary<string, object> formProperties;

            formProperties = new Dictionary<string, object>();
            formProperties.Add("Name", "Drivers"); //100 characters
            formProperties.Add("TableName", "BDO_DRVS");
            formProperties.Add("ObjectType", SAPbobsCOM.BoUDOObjType.boud_MasterData);
            formProperties.Add("CanCancel", SAPbobsCOM.BoYesNoEnum.tYES);
            formProperties.Add("CanClose", SAPbobsCOM.BoYesNoEnum.tYES);
            formProperties.Add("CanCreateDefaultForm", SAPbobsCOM.BoYesNoEnum.tYES);
            formProperties.Add("CanDelete", SAPbobsCOM.BoYesNoEnum.tNO);
            formProperties.Add("CanFind", SAPbobsCOM.BoYesNoEnum.tYES);
            formProperties.Add("CanYearTransfer", SAPbobsCOM.BoYesNoEnum.tYES);

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
            fieldskeysMap.Add("ColumnAlias", "U_firstName");
            fieldskeysMap.Add("ColumnDescription", "First Name"); //30 characters
            listFindColumns.Add(fieldskeysMap);
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "U_lastName");
            fieldskeysMap.Add("ColumnDescription", "Last Name"); //30 characters
            listFindColumns.Add(fieldskeysMap);
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "U_tin");
            fieldskeysMap.Add("ColumnDescription", "Driver TIN"); //30 characters
            listFindColumns.Add(fieldskeysMap);
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "U_notRsdnt");
            fieldskeysMap.Add("ColumnDescription", "Driver Not Resident"); //30 characters
            listFindColumns.Add(fieldskeysMap);

            formProperties.Add("FindColumns", listFindColumns);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Editable", SAPbobsCOM.BoYesNoEnum.tYES);
            fieldskeysMap.Add("FormColumnAlias", "Code");
            fieldskeysMap.Add("FormColumnDescription", "Code"); //30 characters
            listFormColumns.Add(fieldskeysMap);

            formProperties.Add("FormColumns", listFormColumns);

            UDO.registerUDO(code, formProperties, out errorText);

            GC.Collect();
        }

        public static void addMenus()
        {
            SAPbouiCOM.MenuItem menuItem;
            SAPbouiCOM.MenuItem fatherMenuItem;
            SAPbouiCOM.MenuCreationParams oCreationPackage;

            try
            {
                fatherMenuItem = Program.uiApp.Menus.Item("43544");
                // Add a pop-up menu item
                oCreationPackage = (SAPbouiCOM.MenuCreationParams)Program.uiApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                oCreationPackage.Checked = false;
                oCreationPackage.Enabled = true;
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "UDO_F_BDO_DRVS_D";
                oCreationPackage.String = BDOSResources.getTranslate("DriverMasterData");
                oCreationPackage.Position = fatherMenuItem.SubMenus.Count - 1;

                menuItem = fatherMenuItem.SubMenus.AddEx(oCreationPackage);
            }
            catch
            {

            }
        }

        public static void createFormItems(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            Dictionary<string, object> formItems;

            string itemName = "";

            int left_s = 6;
            int left_e = 127;
            int height = 15;
            int top = 6;
            int width_s = 121;
            int width_e = 148;

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "13_U_S"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("FirstName"));
            formItems.Add("LinkTo", "13_U_E");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "13_U_E"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_DRVS");
            formItems.Add("Alias", "U_firstName");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "14_U_S"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("LastName"));
            formItems.Add("LinkTo", "14_U_E");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "14_U_E"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_DRVS");
            formItems.Add("Alias", "U_lastName");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "15_U_S"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("DriverTin"));
            formItems.Add("LinkTo", "15_U_E");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "15_U_E"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_DRVS");
            formItems.Add("Alias", "U_tin");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            formItems.Add("Left", left_e + width_e + 1);
            formItems.Add("Width", 18);
            formItems.Add("Top", top);
            formItems.Add("Height", 18);
            formItems.Add("Image", "15886_MENU"); //"15886_MENU_CHECKED" //"WS_TOPSEARCH_PICKER"); //WS_TOPSEARCH_PICKER "PNG_1536_MENU" //WS_COCKPIT_SWITCH_UI_MENU_ITEM
            formItems.Add("UID", "BDO_TinBtn"); //

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "16_U_CH"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDO_DRVS");
            formItems.Add("Alias", "U_notRsdnt");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s + 45);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Caption", BDOSResources.getTranslate("DriverNotResident"));
            formItems.Add("ValOff", "N");
            formItems.Add("ValOn", "Y");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            GC.Collect();
        }

        public static void setSizeForm(SAPbouiCOM.Form oForm, out string errorText)
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

        public static void resizeForm(SAPbouiCOM.Form oForm, out string errorText)
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

        //public static void formDataLoad(  SAPbouiCOM.Form oForm, out string errorText)
        //{
        //    errorText = null;

        //    BDOSResources.getTranslate("DriverMasterData");
        //}

        public static void uiApp_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD & pVal.BeforeAction == false)
                {
                    BDO_Drivers.createFormItems(oForm, out errorText);
                    Program.FORM_LOAD_FOR_VISIBLE = true;
                }

                if (pVal.ItemUID == "BDO_TinBtn")
                {
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK & pVal.BeforeAction == false)
                    {
                        SAPbouiCOM.Button oButton = (SAPbouiCOM.Button)oForm.Items.Item("BDO_TinBtn").Specific;
                        oButton.Image = "15886_MENU_CHECKED";
                        oForm.Freeze(true);
                        BDO_TinBtn_OnClick(oForm, out errorText);
                        oForm.Freeze(false);
                        oButton.Image = "15886_MENU";
                        if (errorText != null)
                        {
                            Program.uiApp.SetStatusBarMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                            return;
                        }
                    }
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE & pVal.BeforeAction == false & Program.FORM_LOAD_FOR_VISIBLE == true)
                {
                    oForm.Freeze(true);
                    BDO_Drivers.setSizeForm(oForm, out errorText);
                    oForm.Title = BDOSResources.getTranslate("DriverMasterData");
                    oForm.Freeze(false);
                    Program.FORM_LOAD_FOR_VISIBLE = false;
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_RESIZE & pVal.BeforeAction == false)
                {
                    oForm.Freeze(true);
                    BDO_Drivers.resizeForm(oForm, out errorText);
                    oForm.Freeze(false);
                }


            }
        }

        public static void BDO_TinBtn_OnClick(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            string tin = oForm.DataSources.DBDataSources.Item(0).GetValue("U_tin", 0).Trim();

            if (tin == "")
            {
                errorText = BDOSResources.getTranslate("TINNotFill");
                return;
            }

            Dictionary<string, string> rsSettings = CompanyDetails.getRSSettings(out errorText);
            if (errorText != null)
            {
                return;
            }

            string name = BDO_Waybills.getInitFromTIN(tin, out errorText);

            if (name != "")
            {
                string firstLetter = "";
                string seccondLetter = "";

                string firstName = oForm.DataSources.DBDataSources.Item(0).GetValue("U_firstName", 0).Trim();
                string lastName = oForm.DataSources.DBDataSources.Item(0).GetValue("U_lastName", 0).Trim();
                int changeCardName = 1;

                if (firstName != "")
                {
                    changeCardName = Program.uiApp.MessageBox(BDOSResources.getTranslate("NameMismatchFromEnregDoYouWantEditDriver"), 1, BDOSResources.getTranslate("Yes"), BDOSResources.getTranslate("No"), "");
                }

                if (changeCardName == 1)
                {

                    firstLetter = name.Substring(0, name.IndexOf("."));
                    seccondLetter = name.Substring(name.IndexOf(".") + 2);

                    SAPbouiCOM.EditText ofirstName = ((SAPbouiCOM.EditText)(oForm.Items.Item("13_U_E").Specific));
                    ofirstName.Value = firstLetter;

                    SAPbouiCOM.EditText olastName = ((SAPbouiCOM.EditText)(oForm.Items.Item("14_U_E").Specific));
                    olastName.Value = seccondLetter;

                }

            }
            else
            {
                errorText = BDOSResources.getTranslate("NotRecognizeObjectByTINInEnreg");
                return;
            }
        }

        public static void uiApp_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(BusinessObjectInfo.FormTypeEx, Program.currentFormCount);

            if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE)
            {
                if (BusinessObjectInfo.BeforeAction == true)  //& pVal.InnerEvent == true)
                {
                    if (checkRemoving(oForm, out errorText) == true)
                    {
                        Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("RecordIsUsedInDocuments"));
                        Program.uiApp.MessageBox(BDOSResources.getTranslate("OperationUnsuccesfullSeeLog"));
                        BubbleEvent = false;
                    }
                }
            }
        }


        public static bool checkRemoving(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            SAPbouiCOM.DBDataSource DocDBSourceTAXP = oForm.DataSources.DBDataSources.Item(0);
            string code = DocDBSourceTAXP.GetValue("Code", 0).Trim();

            Dictionary<string, string> listTables = new Dictionary<string, string>();
            listTables.Add("@BDO_VECL", "U_drvCode"); //Vehicles
            listTables.Add("@BDO_WBLD", "U_drvCode"); //Waybills

            return CommonFunctions.codeIsUsed(listTables, code);
        }

    }
}
