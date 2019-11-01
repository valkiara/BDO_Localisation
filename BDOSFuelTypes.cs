using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace BDO_Localisation_AddOn
{
    class BDOSFuelTypes
    {
        public static void createMasterDataUDO(out string errorText)
        {
            string tableName = "BDOSFUTP";
            string description = "Fuel Types";

            int result = UDO.addUserTable(tableName, description, SAPbobsCOM.BoUTBTableType.bott_MasterData, out errorText);

            if (result != 0)
            {
                return;
            }

            Dictionary<string, object> fieldskeysMap;
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "ItemCode");
            fieldskeysMap.Add("TableName", "BDOSFUTP");
            fieldskeysMap.Add("Description", "Item No.");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "ItemName");
            fieldskeysMap.Add("TableName", "BDOSFUTP");
            fieldskeysMap.Add("Description", "Item Description");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 100);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "UomEntry");
            fieldskeysMap.Add("TableName", "BDOSFUTP");
            fieldskeysMap.Add("Description", "UoM Abs. Entry");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Numeric);
            fieldskeysMap.Add("EditSize", 11);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "UomCode");
            fieldskeysMap.Add("TableName", "BDOSFUTP");
            fieldskeysMap.Add("Description", "UoM Code");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 20);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "PerKm");
            fieldskeysMap.Add("TableName", "BDOSFUTP");
            fieldskeysMap.Add("Description", "Per 100 km");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "PerHr");
            fieldskeysMap.Add("TableName", "BDOSFUTP");
            fieldskeysMap.Add("Description", "Per Hour");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            GC.Collect();
        }

        public static void registerUDO()
        {
            SAPbobsCOM.UserObjectsMD oUserObjectMD = null;
            SAPbobsCOM.UserObjectMD_FindColumns oUDOFind = null;
            SAPbobsCOM.UserObjectMD_FormColumns oUDOForm = null;
            GC.Collect();
            oUserObjectMD = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);
            Marshal.ReleaseComObject(oUserObjectMD);
            oUserObjectMD = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD) as SAPbobsCOM.UserObjectsMD;
            oUDOFind = oUserObjectMD.FindColumns;
            oUDOForm = oUserObjectMD.FormColumns;

            var retval = oUserObjectMD.GetByKey("UDO_F_BDOSFUTP_D");

            if (!retval)
            {
                oUserObjectMD.Code = "UDO_F_BDOSFUTP_D";
                oUserObjectMD.Name = "Fuel Types";
                oUserObjectMD.TableName = "BDOSFUTP";
                oUserObjectMD.ObjectType = SAPbobsCOM.BoUDOObjType.boud_MasterData;
                oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tYES;

                //Find
                oUDOFind.ColumnAlias = "Code";
                oUDOFind.ColumnDescription = "Code";
                oUDOFind.Add();
                oUDOFind.ColumnAlias = "Name";
                oUDOFind.ColumnDescription = "Name";
                oUDOFind.Add();
                oUDOFind.ColumnAlias = "U_ItemCode";
                oUDOFind.ColumnDescription = "Item No.";
                oUDOFind.Add();
                oUDOFind.ColumnAlias = "U_UomEntry";
                oUDOFind.ColumnDescription = "UoM Abs. Entry";
                oUDOFind.Add();
                oUDOFind.ColumnAlias = "U_UomCode";
                oUDOFind.ColumnDescription = "UoM Code";
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
                    oCreationPackage.UniqueID = "UDO_F_BDOSFUTP_D";
                    oCreationPackage.String = BDOSResources.getTranslate("FuelTypes");
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
                        oForm.Freeze(true);
                        setSizeForm(oForm);
                        oForm.Title = BDOSResources.getTranslate("FuelTypes");
                        oForm.Freeze(false);
                        Program.FORM_LOAD_FOR_VISIBLE = false;
                    }
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_RESIZE && !pVal.BeforeAction)
                {
                    oForm.Freeze(true);
                    resizeForm(oForm);
                    oForm.Freeze(false);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE && !pVal.BeforeAction)
                {
                    if (Program.FORM_LOAD_FOR_ACTIVATE)
                    {
                        SAPbouiCOM.StaticText staticText = oForm.Items.Item("0_U_S").Specific;
                        staticText.Caption = BDOSResources.getTranslate("Code");
                        Program.FORM_LOAD_FOR_ACTIVATE = false;

                        setVisibleFormItems(oForm);
                    }
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                {
                    SAPbouiCOM.IChooseFromListEvent oCFLEvento = ((SAPbouiCOM.IChooseFromListEvent)(pVal));
                    chooseFromList(oForm, pVal, oCFLEvento);
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

            top = top + height + 1;

            string objectType = "4"; //Items
            string itemCodeCFL = "ItemCodeCFL";
            FormsB1.addChooseFromList(oForm, false, objectType, itemCodeCFL);

            formItems = new Dictionary<string, object>();
            itemName = "ItemCodeS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("FuelCode"));
            formItems.Add("LinkTo", "ItemCodeE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }

            formItems = new Dictionary<string, object>();
            itemName = "ItemCodeE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSFUTP");
            formItems.Add("Alias", "U_ItemCode");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("ChooseFromListUID", itemCodeCFL);
            formItems.Add("ChooseFromListAlias", "ItemCode");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }

            formItems = new Dictionary<string, object>();
            itemName = "ItemLB"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            formItems.Add("Left", left_e - 20);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("LinkTo", "ItemCodeE");
            formItems.Add("LinkedObjectType", "4"); //Items

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "ItemNameS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("FuelDescription"));
            formItems.Add("LinkTo", "ItemNameE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }

            formItems = new Dictionary<string, object>();
            itemName = "ItemNameE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSFUTP");
            formItems.Add("Alias", "U_ItemName");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Enabled", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }
            oForm.Items.Item("ItemNameE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "UomEntryS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("UomEntry"));
            formItems.Add("LinkTo", "UomEntryE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }

            formItems = new Dictionary<string, object>();
            itemName = "UomEntryE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSFUTP");
            formItems.Add("Alias", "U_UomEntry");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e / 3);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Enabled", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }
            oForm.Items.Item("UomEntryE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);

            formItems = new Dictionary<string, object>();
            itemName = "UomLB"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            formItems.Add("Left", left_e - 20);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("LinkTo", "UomEntryE");
            formItems.Add("LinkedObjectType", "10000199"); //UoM Master Data

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }

            formItems = new Dictionary<string, object>();
            itemName = "UomCodeE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSFUTP");
            formItems.Add("Alias", "U_UomCode");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e + width_e / 3);
            formItems.Add("Width", width_e / 3 * 2);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Enabled", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }
            oForm.Items.Item("UomCodeE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);

            top = top + height + 1;

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
            formItems.Add("TableName", "@BDOSFUTP");
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
            formItems.Add("TableName", "@BDOSFUTP");
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

            GC.Collect();
        }

        public static void chooseFromList(SAPbouiCOM.Form oForm, SAPbouiCOM.ItemEvent pVal, SAPbouiCOM.IChooseFromListEvent oCFLEvento)
        {
            try
            {
                oForm.Freeze(true);

                if (pVal.BeforeAction)
                {
                    if (oCFLEvento.ChooseFromListUID == "ItemCodeCFL")
                    {
                        SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID);
                        SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();

                        SAPbouiCOM.Condition oCon = oCons.Add();
                        oCon.Alias = "InvntItem";
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = "Y";
                        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                        oCon = oCons.Add();
                        oCon.Alias = "validFor";
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = "Y";
                        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                        SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        oRecordSet.DoQuery(@"SELECT ""ItmsGrpCod"" FROM ""OITB"" WHERE ""U_BDOSFuel"" = 'Y'");
                        int recordCount = oRecordSet.RecordCount;
                        int i = 0;

                        while (!oRecordSet.EoF)
                        {
                            int itmsGrpCod = Convert.ToInt32(oRecordSet.Fields.Item("ItmsGrpCod").Value);
                            oCon = oCons.Add();
                            oCon.Alias = "ItmsGrpCod";
                            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                            oCon.CondVal = itmsGrpCod.ToString();
                            if (i == 0)
                                oCon.BracketOpenNum = 1;
                            if (i < recordCount - 1)
                            {
                                oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR;
                            }
                            if (i == recordCount - 1)
                            {
                                oCon.BracketCloseNum = 1;
                                oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_NONE;
                            }
                            i++;
                            oRecordSet.MoveNext();
                        }
                        oCFL.SetConditions(oCons);
                    }
                }
                else
                {
                    SAPbouiCOM.DataTable oDataTable = oCFLEvento.SelectedObjects;

                    if (oDataTable != null)
                    {
                        if (oCFLEvento.ChooseFromListUID == "ItemCodeCFL")
                        {
                            string itemCode = oDataTable.GetValue("ItemCode", 0);
                            string itemName = oDataTable.GetValue("ItemName", 0);
                            int uomEntry = oDataTable.GetValue("IUoMEntry", 0);
                            string uomCode = oDataTable.GetValue("InvntryUom", 0);

                            oForm.Items.Item("ItemNameE").Enabled = true;
                            oForm.Items.Item("UomEntryE").Enabled = true;
                            oForm.Items.Item("UomCodeE").Enabled = true;

                            LanguageUtils.IgnoreErrors<string>(() => oForm.Items.Item("ItemCodeE").Specific.Value = itemCode);
                            LanguageUtils.IgnoreErrors<string>(() => oForm.Items.Item("ItemNameE").Specific.Value = itemName);
                            LanguageUtils.IgnoreErrors<string>(() => oForm.Items.Item("UomEntryE").Specific.Value = uomEntry.ToString());
                            LanguageUtils.IgnoreErrors<string>(() => oForm.Items.Item("UomCodeE").Specific.Value = uomCode);

                            oForm.Items.Item("ItemCodeE").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            oForm.Items.Item("ItemNameE").Enabled = false;
                            oForm.Items.Item("UomEntryE").Enabled = false;
                            oForm.Items.Item("UomCodeE").Enabled = false;
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
                GC.Collect();
            }
        }

        public static void resizeForm(SAPbouiCOM.Form oForm)
        {
            try
            {
                SAPbouiCOM.Item oItem = oForm.Items.Item("1");
                oItem.Top = oForm.ClientHeight - 25;

                oItem = oForm.Items.Item("2");
                oItem.Top = oForm.ClientHeight - 25;
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

        public static void setVisibleFormItems(SAPbouiCOM.Form oForm)
        {
            try
            {
                oForm.Items.Item("0_U_E").Enabled = (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE);
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

        public static bool checkRemoving(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.DBDataSource DocDBSourceTAXP = oForm.DataSources.DBDataSources.Item(0);
            string code = DocDBSourceTAXP.GetValue("Code", 0).Trim();

            Dictionary<string, string> listTables = new Dictionary<string, string>();
            listTables.Add("OITM", "U_BDOSFuTp");
            listTables.Add("@BDOSFUC1", "U_FuTpCode");

            return CommonFunctions.codeIsUsed(listTables, code);
        }
    }
}
