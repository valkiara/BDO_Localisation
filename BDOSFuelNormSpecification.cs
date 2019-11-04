using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace BDO_Localisation_AddOn
{
    class BDOSFuelNormSpecification
    {
        public static void createMasterDataUDO(out string errorText)
        {
            string tableName = "BDOSFUNR";
            string description = "Specification of Fuel Norm";

            int result = UDO.addUserTable(tableName, description, SAPbobsCOM.BoUTBTableType.bott_MasterData, out errorText);

            if (result != 0)
            {
                return;
            }

            Dictionary<string, object> fieldskeysMap;

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "Active");
            fieldskeysMap.Add("TableName", "BDOSFUNR");
            fieldskeysMap.Add("Description", "Active");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 1);
            fieldskeysMap.Add("DefaultValue", "Y");

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "PrjCode");
            fieldskeysMap.Add("TableName", "BDOSFUNR");
            fieldskeysMap.Add("Description", "Project Code");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 20);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "Fixed");
            fieldskeysMap.Add("TableName", "BDOSFUNR");
            fieldskeysMap.Add("Description", "Fixed");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 1);
            fieldskeysMap.Add("DefaultValue", "Y");

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "PerKm");
            fieldskeysMap.Add("TableName", "BDOSFUNR");
            fieldskeysMap.Add("Description", "Per 100 km");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "PerHr");
            fieldskeysMap.Add("TableName", "BDOSFUNR");
            fieldskeysMap.Add("Description", "Per Hour");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            tableName = "BDOSFUN1";
            description = "Specif of Fuel Norm Child1";

            result = UDO.addUserTable(tableName, description, SAPbobsCOM.BoUTBTableType.bott_MasterDataLines, out errorText);

            if (result != 0)
            {
                return;
            }

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "CrtrCode");
            fieldskeysMap.Add("TableName", "BDOSFUN1");
            fieldskeysMap.Add("Description", "Criteria Code");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "CrtrName");
            fieldskeysMap.Add("TableName", "BDOSFUN1");
            fieldskeysMap.Add("Description", "Criteria Name");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 100);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "CrtrValue");
            fieldskeysMap.Add("TableName", "BDOSFUN1");
            fieldskeysMap.Add("Description", "Criteria Value");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 100);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "CrtrPr");
            fieldskeysMap.Add("TableName", "BDOSFUN1");
            fieldskeysMap.Add("Description", "Criteria Percentage");
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
            SAPbobsCOM.IUserObjectMD_ChildTables oUDOChildTables = null;
            GC.Collect();
            oUserObjectMD = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);
            Marshal.ReleaseComObject(oUserObjectMD);
            oUserObjectMD = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD) as SAPbobsCOM.UserObjectsMD;
            oUDOFind = oUserObjectMD.FindColumns;
            oUDOForm = oUserObjectMD.FormColumns;
            oUDOChildTables = oUserObjectMD.ChildTables;

            var retval = oUserObjectMD.GetByKey("UDO_F_BDOSFUNR_D");

            if (!retval)
            {
                oUserObjectMD.Code = "UDO_F_BDOSFUNR_D";
                oUserObjectMD.Name = "Specification of Fuel Norm";
                oUserObjectMD.TableName = "BDOSFUNR";
                oUserObjectMD.ObjectType = SAPbobsCOM.BoUDOObjType.boud_MasterData;
                oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.LogTableName = "ABDOSFUNR";

                //Find
                oUDOFind.ColumnAlias = "Code";
                oUDOFind.ColumnDescription = "Code";
                oUDOFind.Add();
                oUDOFind.ColumnAlias = "Name";
                oUDOFind.ColumnDescription = "Name";
                oUDOFind.Add();
                oUDOFind.ColumnAlias = "U_PrjCode";
                oUDOFind.ColumnDescription = "Project Code";
                oUDOFind.Add();
                oUDOFind.ColumnAlias = "U_Active";
                oUDOFind.ColumnDescription = "Active";
                oUDOFind.Add();
                oUDOFind.ColumnAlias = "U_Fixed";
                oUDOFind.ColumnDescription = "Fixed";
                oUDOFind.Add();
                oUDOFind.ColumnAlias = "U_PerKm";
                oUDOFind.ColumnDescription = "Per 100 km";
                oUDOFind.Add();
                oUDOFind.ColumnAlias = "U_PerHr";
                oUDOFind.ColumnDescription = "Per Hour";
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

                oUDOChildTables.Add();
                oUDOChildTables.SetCurrentLine(oUDOChildTables.Count - 1);
                oUDOChildTables.TableName = "BDOSFUN1";
                oUDOChildTables.ObjectName = "BDOSFUN1";
                oUDOChildTables.LogTableName = "ABDOSFUN1";

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
                    oCreationPackage.UniqueID = "UDO_F_BDOSFUNR_D";
                    oCreationPackage.String = BDOSResources.getTranslate("SpecificationOfFuelNorm");
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
                    createFormItems(oForm);
                    Program.FORM_LOAD_FOR_VISIBLE = true;
                    Program.FORM_LOAD_FOR_ACTIVATE = true;
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE && !pVal.BeforeAction)
                {
                    if (Program.FORM_LOAD_FOR_VISIBLE)
                    {
                        setSizeForm(oForm);
                        oForm.Freeze(true);
                        oForm.Title = BDOSResources.getTranslate("SpecificationOfFuelNorm");
                        oForm.Freeze(false);
                        Program.FORM_LOAD_FOR_VISIBLE = false;
                        setVisibleFormItems(oForm);
                    }
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_RESIZE && !pVal.BeforeAction)
                {
                    resizeForm(oForm);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE && !pVal.BeforeAction)
                {
                    if (Program.FORM_LOAD_FOR_ACTIVATE)
                    {
                        oForm.Freeze(true);
                        try
                        {
                            SAPbouiCOM.StaticText staticText = oForm.Items.Item("0_U_S").Specific;
                            staticText.Caption = BDOSResources.getTranslate("Code");
                            staticText = oForm.Items.Item("1_U_S").Specific;
                            staticText.Caption = BDOSResources.getTranslate("Name");

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

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                {
                    SAPbouiCOM.IChooseFromListEvent oCFLEvento = ((SAPbouiCOM.IChooseFromListEvent)(pVal));
                    chooseFromList(oForm, pVal, oCFLEvento);

                    //if (pVal.ItemUID == "CrtrMTR" && !pVal.BeforeAction)
                    //    addMatrixRow(oForm);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && !pVal.BeforeAction)
                {
                    if (!pVal.InnerEvent && (pVal.ItemUID == "CriteriaOB" || pVal.ItemUID == "FixedOB"))
                    {
                        if (oForm.DataSources.UserDataSources.Item("CriteriaOB").ValueEx == "1")
                            oForm.DataSources.DBDataSources.Item("@BDOSFUNR").SetValue("U_Fixed", 0, "N");
                        else if (oForm.DataSources.UserDataSources.Item("CriteriaOB").ValueEx == "2")
                            oForm.DataSources.DBDataSources.Item("@BDOSFUNR").SetValue("U_Fixed", 0, "Y");

                        oForm.Items.Item("PrjCodeE").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        setVisibleFormItems(oForm);

                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                        }
                    }
                }

                if (pVal.ItemUID == "addMTRB" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && !pVal.BeforeAction)
                {
                    addMatrixRow(oForm);
                }

                if (pVal.ItemUID == "delMTRB" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && !pVal.BeforeAction)
                {
                    deleteMatrixRow(oForm);
                }
            }
        }

        public static void uiApp_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;

            SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(BusinessObjectInfo.FormTypeEx, Program.currentFormCount);

            if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD && !BusinessObjectInfo.BeforeAction)
            {
                bool isFixed = oForm.DataSources.DBDataSources.Item("@BDOSFUNR").GetValue("U_Fixed", 0) == "Y";
                if (isFixed)
                    oForm.Items.Item("FixedOB").Specific.Selected = isFixed;
                else
                    oForm.Items.Item("CriteriaOB").Specific.Selected = true;

                setVisibleFormItems(oForm);

                oForm.Items.Item("0_U_E").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                oForm.Items.Item("1_U_E").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                oForm.Items.Item("0_U_E").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True);
                oForm.Items.Item("0_U_E").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True);
                oForm.Items.Item("1_U_E").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True);
                oForm.Items.Item("1_U_E").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True);
            }
            else if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE)
            {
                if (BusinessObjectInfo.BeforeAction)
                {
                    if (checkRemoving(oForm))
                    {
                        Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("RecordIsUsedInDocuments"));
                        Program.uiApp.MessageBox(BDOSResources.getTranslate("OperationUnsuccesfullSeeLog"));
                        BubbleEvent = false;
                    }
                }
            }
            else if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD || BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE)
            {
                if (BusinessObjectInfo.BeforeAction)
                {
                    SAPbouiCOM.DBDataSource oDBDataSource = oForm.DataSources.DBDataSources.Item("@BDOSFUNR");
                    bool isFixed = oForm.DataSources.DBDataSources.Item("@BDOSFUNR").GetValue("U_Fixed", 0) == "Y";

                    if (isFixed)
                    {
                        SAPbouiCOM.DBDataSource oDBDataSourceMTR = oForm.DataSources.DBDataSources.Item("@BDOSFUN1");
                        oDBDataSourceMTR.Clear();
                    }
                    else
                    {
                        oDBDataSource.SetValue("U_PerKm", 0, "0");
                        oDBDataSource.SetValue("U_PerHr", 0, "0");

                        //checkDuplicatesInDBDataSources
                        SAPbouiCOM.DBDataSource oDBDataSourceMTR = oForm.DataSources.DBDataSources.Item("@BDOSFUN1");

                        Dictionary<string, SAPbouiCOM.DBDataSource> oKeysDictionary = new Dictionary<string, SAPbouiCOM.DBDataSource>();
                        oKeysDictionary.Add("U_CrtrCode", oDBDataSourceMTR);
                        string errorText;
                        List<string> crtrCodeList = CommonFunctions.checkDuplicatesInDBDataSources(oDBDataSourceMTR, oKeysDictionary, out errorText);
                        if (string.IsNullOrEmpty(errorText) == false)
                        {
                            Program.uiApp.SetStatusBarMessage(errorText + " " + BDOSResources.getTranslate("Code") + ": " + string.Join(",", crtrCodeList), SAPbouiCOM.BoMessageTime.bmt_Short, true);
                            BubbleEvent = false;
                        }
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
            int left_e = 127;
            int height = 15;
            int top = 6;
            int width_s = 121;
            int width_e = 148;

            top = top + 2 * height;

            FormsB1.addChooseFromList(oForm, false, "63", "ProjectCodeCFL"); //Project Codes
            FormsB1.addChooseFromList(oForm, false, "UDO_F_BDOSFUCR_D", "FuelCriteriaCodeCFL"); //Fuel Criteria

            formItems = new Dictionary<string, object>();
            itemName = "PrjCodeS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Project"));
            formItems.Add("LinkTo", "PrjCodeE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }

            formItems = new Dictionary<string, object>();
            itemName = "PrjCodeE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSFUNR");
            formItems.Add("Alias", "U_PrjCode");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("ChooseFromListUID", "ProjectCodeCFL");
            formItems.Add("ChooseFromListAlias", "PrjCode");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }

            formItems = new Dictionary<string, object>();
            itemName = "PrjLB"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            formItems.Add("Left", left_e - 20);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("LinkTo", "PrjCodeE");
            formItems.Add("LinkedObjectType", "63"); //Project Codes

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "ActiveCH"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSFUNR");
            formItems.Add("Alias", "U_Active");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            formItems.Add("Width", width_e);
            formItems.Add("Left", left_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Active"));
            formItems.Add("ValOff", "N");
            formItems.Add("ValOn", "Y");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "CriteriaOB"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "UserDataSources");
            formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
            formItems.Add("Length", 1);
            formItems.Add("TableName", "");
            formItems.Add("Alias", itemName);
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_OPTION_BUTTON);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("FuelCriteria"));
            formItems.Add("ValOn", "Y");
            formItems.Add("ValOff", "N");
            formItems.Add("AffectsFormMode", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }

            formItems = new Dictionary<string, object>();
            itemName = "FixedOB"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "UserDataSources");
            formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
            formItems.Add("Length", 1);
            formItems.Add("TableName", "");
            formItems.Add("Alias", itemName);
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_OPTION_BUTTON);
            formItems.Add("Left", left_s + width_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Fixed"));
            formItems.Add("GroupWith", "CriteriaOB");
            formItems.Add("ValOn", "Y");
            formItems.Add("ValOff", "N");
            formItems.Add("Selected", true);
            formItems.Add("AffectsFormMode", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }

            top = top + 3 * height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "PerKmS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("PerKm"));
            formItems.Add("LinkTo", "PerKmE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }

            formItems = new Dictionary<string, object>();
            itemName = "PerKmE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSFUNR");
            formItems.Add("Alias", "U_PerKm");
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

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "PerHrS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("PerHr"));
            formItems.Add("LinkTo", "PerHrE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }

            formItems = new Dictionary<string, object>();
            itemName = "PerHrE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSFUNR");
            formItems.Add("Alias", "U_PerHr");
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

            top = top - height - 1;

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

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "CrtrMTR"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
            formItems.Add("Left", left_s);
            formItems.Add("Height", 150);
            formItems.Add("Top", top);
            formItems.Add("UID", itemName);
            formItems.Add("Visible", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }

            SAPbouiCOM.DBDataSource oDBDataSource = oForm.DataSources.DBDataSources.Add("@BDOSFUN1");

            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("CrtrMTR").Specific));
            oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
            SAPbouiCOM.Columns oColumns = oMatrix.Columns;

            SAPbouiCOM.Column oColumn = oColumns.Add("LineID", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "#";
            oColumn.Editable = false;

            oColumn = oColumns.Item("LineID");
            oColumn.DataBind.SetBound(true, "@BDOSFUN1", "LineId");

            oColumn = oColumns.Add("CrtrCode", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Code");

            oColumn = oColumns.Item("CrtrCode");
            oColumn.DataBind.SetBound(true, "@BDOSFUN1", "U_CrtrCode");
            oColumn.ChooseFromListUID = "FuelCriteriaCodeCFL";
            oColumn.ChooseFromListAlias = "Code";
            SAPbouiCOM.LinkedButton oLink = oColumn.ExtendedObject;
            oLink.LinkedObjectType = "UDO_F_BDOSFUCR_D";

            oColumn = oColumns.Add("CrtrName", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Name");
            oColumn.Editable = false;

            oColumn = oColumns.Item("CrtrName");
            oColumn.DataBind.SetBound(true, "@BDOSFUN1", "U_CrtrName");

            oColumn = oColumns.Add("CrtrValue", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Value");
            oColumn.Editable = false;

            oColumn = oColumns.Item("CrtrValue");
            oColumn.DataBind.SetBound(true, "@BDOSFUN1", "U_CrtrValue");

            oColumn = oColumns.Add("CrtrPr", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Percentage");
            oColumn.Editable = false;

            oColumn = oColumns.Item("CrtrPr");
            oColumn.DataBind.SetBound(true, "@BDOSFUN1", "U_CrtrPr");

            //oMatrix.Clear();
            //oDBDataSource.Query();
            //oMatrix.LoadFromDataSource();

            GC.Collect();
        }

        public static void chooseFromList(SAPbouiCOM.Form oForm, SAPbouiCOM.ItemEvent pVal, SAPbouiCOM.IChooseFromListEvent oCFLEvento)
        {
            try
            {
                oForm.Freeze(true);

                if (pVal.BeforeAction)
                {
                    if (oCFLEvento.ChooseFromListUID == "ProjectCodeCFL")
                    {
                        SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID);
                        SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();

                        SAPbouiCOM.Condition oCon = oCons.Add();
                        oCon.Alias = "Active";
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = "Y";
                        oCFL.SetConditions(oCons);
                    }
                }
                else
                {
                    SAPbouiCOM.DataTable oDataTable = oCFLEvento.SelectedObjects;

                    if (oDataTable != null)
                    {
                        if (oCFLEvento.ChooseFromListUID == "ProjectCodeCFL")
                        {
                            string prjCode = oDataTable.GetValue("PrjCode", 0);
                            LanguageUtils.IgnoreErrors<string>(() => oForm.Items.Item("PrjCodeE").Specific.Value = prjCode);
                        }
                        else if (oCFLEvento.ChooseFromListUID == "FuelCriteriaCodeCFL")
                        {
                            string crtrCode = oDataTable.GetValue("Code", 0);
                            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            string query = @"SELECT 
                                           ""@BDOSFUCR"".""Name"", 
                                           ""@BDOSFUCR"".""U_Value"", 
                                           ""@BDOSFUCR"".""U_Percentage""
                                    FROM   ""@BDOSFUCR""
                                    WHERE  ""@BDOSFUCR"".""Code"" = '" + crtrCode + @"'";

                            oRecordSet.DoQuery(query);

                            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("CrtrMTR").Specific));
                            LanguageUtils.IgnoreErrors<string>(() => oMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value = crtrCode);
                            if (!oRecordSet.EoF)
                            {
                                LanguageUtils.IgnoreErrors<string>(() => oMatrix.Columns.Item("CrtrName").Cells.Item(pVal.Row).Specific.Value = oRecordSet.Fields.Item("Name").Value);
                                LanguageUtils.IgnoreErrors<string>(() => oMatrix.Columns.Item("CrtrValue").Cells.Item(pVal.Row).Specific.Value = oRecordSet.Fields.Item("U_Value").Value);
                                LanguageUtils.IgnoreErrors<string>(() => oMatrix.Columns.Item("CrtrPr").Cells.Item(pVal.Row).Specific.Value = oRecordSet.Fields.Item("U_Percentage").Value.ToString());
                            }
                        }
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
                oForm.ClientHeight = Program.uiApp.Desktop.Width / 6;
                oForm.Height = Program.uiApp.Desktop.Width / 4;
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
                oForm.Items.Item("1_U_S").Left = oForm.Items.Item("0_U_S").Left;
                oForm.Items.Item("1_U_S").Top = oForm.Items.Item("0_U_S").Top + oForm.Items.Item("0_U_S").Height + 1;
                oForm.Items.Item("1_U_E").Left = oForm.Items.Item("0_U_E").Left;
                oForm.Items.Item("1_U_E").Top = oForm.Items.Item("0_U_E").Top + oForm.Items.Item("0_U_E").Height + 1;
                oForm.Items.Item("1").Top = oForm.ClientHeight - 25;
                oForm.Items.Item("2").Top = oForm.ClientHeight - 25;

                SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("CrtrMTR").Specific));
                int mtrWidth = oForm.ClientWidth / 3 * 2;
                oForm.Items.Item("CrtrMTR").Width = mtrWidth;
                oMatrix.Columns.Item("LineID").Width = 19;
                mtrWidth -= 19;
                oMatrix.Columns.Item("CrtrCode").Width = mtrWidth / 4;
                oMatrix.Columns.Item("CrtrName").Width = mtrWidth / 4;
                oMatrix.Columns.Item("CrtrValue").Width = mtrWidth / 4;
                oMatrix.Columns.Item("CrtrPr").Width = mtrWidth / 4;
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
                bool isFixed = oForm.DataSources.DBDataSources.Item("@BDOSFUNR").GetValue("U_Fixed", 0) == "Y";
                oForm.Items.Item("PerKmS").Visible = isFixed;
                oForm.Items.Item("PerKmE").Visible = isFixed;
                oForm.Items.Item("PerHrS").Visible = isFixed;
                oForm.Items.Item("PerHrE").Visible = isFixed;
                oForm.Items.Item("addMTRB").Visible = !isFixed;
                oForm.Items.Item("delMTRB").Visible = !isFixed;
                oForm.Items.Item("CrtrMTR").Visible = !isFixed;
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
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("CrtrMTR").Specific;
                oMatrix.FlushToDataSource();

                SAPbouiCOM.DBDataSource oDBDataSourceMTR = oForm.DataSources.DBDataSources.Item("@BDOSFUN1");
                if (!string.IsNullOrEmpty(oDBDataSourceMTR.GetValue("U_CrtrCode", oDBDataSourceMTR.Size - 1)))
                {
                    oDBDataSourceMTR.InsertRecord(oDBDataSourceMTR.Size);
                }
                oDBDataSourceMTR.SetValue("LineId", oDBDataSourceMTR.Size - 1, oDBDataSourceMTR.Size.ToString());

                oMatrix.LoadFromDataSource();

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
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

        public static void deleteMatrixRow(SAPbouiCOM.Form oForm)
        {
            oForm.Freeze(true);
            try
            {
                SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("CrtrMTR").Specific));
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
                        oForm.DataSources.DBDataSources.Item("@BDOSFUN1").RemoveRecord(row - deletedRowCount);
                        firstRow = row;
                    }
                }

                SAPbouiCOM.DBDataSource oDBDataSourceMTR = oForm.DataSources.DBDataSources.Item("@BDOSFUN1");
                int rowCount = oDBDataSourceMTR.Size;

                for (int i = 1; i <= rowCount; i++)
                {
                    string itemCode = oDBDataSourceMTR.GetValue("U_CrtrCode", i - 1);
                    if (!string.IsNullOrEmpty(itemCode))
                    {
                        oDBDataSourceMTR.SetValue("LineId", i - 1, i.ToString());
                    }
                }
                oMatrix.LoadFromDataSource();

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE && deletedRowCount > 0)
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                }
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

        public static bool checkRemoving(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.DBDataSource DocDBSourceTAXP = oForm.DataSources.DBDataSources.Item(0);
            string code = DocDBSourceTAXP.GetValue("Code", 0).Trim();

            Dictionary<string, string> listTables = new Dictionary<string, string>();
            listTables.Add("@BDOSFUCN", "U_FuNrCode");

            return CommonFunctions.codeIsUsed(listTables, code);
        }
    }
}
