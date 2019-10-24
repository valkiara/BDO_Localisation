using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BDO_Localisation_AddOn
{
    class BDOSFuelConsumptionAct
    {
        public static void createDocumentUDO(out string errorText)
        {
            string tableName = "BDOSFUCN";
            string description = "Fuel Consumption Act";

            int result = UDO.addUserTable(tableName, description, SAPbobsCOM.BoUTBTableType.bott_Document, out errorText);

            if (result != 0)
            {
                return;
            }

            Dictionary<string, object> fieldskeysMap;

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "DocDate");
            fieldskeysMap.Add("TableName", "BDOSFUCN");
            fieldskeysMap.Add("Description", "Posting Date");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Date);
            fieldskeysMap.Add("EditSize", 15);
            fieldskeysMap.Add("Mandatory", SAPbobsCOM.BoYesNoEnum.tYES);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "DateFrom");
            fieldskeysMap.Add("TableName", "BDOSFUCN");
            fieldskeysMap.Add("Description", "Date From");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Date);
            fieldskeysMap.Add("EditSize", 15);
            fieldskeysMap.Add("Mandatory", SAPbobsCOM.BoYesNoEnum.tYES);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "DateTo");
            fieldskeysMap.Add("TableName", "BDOSFUCN");
            fieldskeysMap.Add("Description", "Date To");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Date);
            fieldskeysMap.Add("EditSize", 15);
            fieldskeysMap.Add("Mandatory", SAPbobsCOM.BoYesNoEnum.tYES);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "PrjCode");
            fieldskeysMap.Add("TableName", "BDOSFUCN");
            fieldskeysMap.Add("Description", "Project Code");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 20);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "FuNrCode");
            fieldskeysMap.Add("TableName", "BDOSFUCN");
            fieldskeysMap.Add("Description", "Specification of Fuel Norm Code");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            tableName = "BDOSFUC1";
            description = "Fuel Consumption Act Child1";

            result = UDO.addUserTable(tableName, description, SAPbobsCOM.BoUTBTableType.bott_DocumentLines, out errorText);

            if (result != 0)
            {
                return;
            }

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "ItemCode");
            fieldskeysMap.Add("TableName", "BDOSFUC1");
            fieldskeysMap.Add("Description", "Item No.");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "ItemName");
            fieldskeysMap.Add("TableName", "BDOSFUC1");
            fieldskeysMap.Add("Description", "Item Description");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 100);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            //წვის ტიპი
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "FuTpCode");
            fieldskeysMap.Add("TableName", "BDOSFUC1");
            fieldskeysMap.Add("Description", "Fuel Type Code");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            //საწვავი
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "FuelCode");
            fieldskeysMap.Add("TableName", "BDOSFUC1");
            fieldskeysMap.Add("Description", "Fuel No.");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            //საწვავის ერთეული
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "FuUomEntry");
            fieldskeysMap.Add("TableName", "BDOSFUC1");
            fieldskeysMap.Add("Description", "Fuel UoM Abs. Entry");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Numeric);
            fieldskeysMap.Add("EditSize", 11);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            //წვა 100 კმ-ში
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "FuPerKm");
            fieldskeysMap.Add("TableName", "BDOSFUC1");
            fieldskeysMap.Add("Description", "Per 100 km");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            //წვა საათში
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "FuPerHr");
            fieldskeysMap.Add("TableName", "BDOSFUC1");
            fieldskeysMap.Add("Description", "Per Hour");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            //ოდომეტრის საწყისი ჩვენება (კმ)
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "OdmtrStart");
            fieldskeysMap.Add("TableName", "BDOSFUC1");
            fieldskeysMap.Add("Description", "Starting Value of Odometer");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            //ოდომეტრის საბოლოო ჩვენება (კმ)
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "OdmtrEnd");
            fieldskeysMap.Add("TableName", "BDOSFUC1");
            fieldskeysMap.Add("Description", "Ending Value of Odometer");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            //ნამუშევარი საათები
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "HrsWorked");
            fieldskeysMap.Add("TableName", "BDOSFUC1");
            fieldskeysMap.Add("Description", "Hours Worked");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            //ხარჯვა ნორმის მიხედვით 
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "NormCn");
            fieldskeysMap.Add("TableName", "BDOSFUC1");
            fieldskeysMap.Add("Description", "Norm Consumption");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            //ფაქტიური ხარჯვა
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "ActuallyCn");
            fieldskeysMap.Add("TableName", "BDOSFUC1");
            fieldskeysMap.Add("Description", "Actually Consumption");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "Dimension1");
            fieldskeysMap.Add("TableName", "BDOSFUC1");
            fieldskeysMap.Add("Description", "Dimension1");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "Dimension2");
            fieldskeysMap.Add("TableName", "BDOSFUC1");
            fieldskeysMap.Add("Description", "Dimension2");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "Dimension3");
            fieldskeysMap.Add("TableName", "BDOSFUC1");
            fieldskeysMap.Add("Description", "Dimension3");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "Dimension4");
            fieldskeysMap.Add("TableName", "BDOSFUC1");
            fieldskeysMap.Add("Description", "Dimension4");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "Dimension5");
            fieldskeysMap.Add("TableName", "BDOSFUC1");
            fieldskeysMap.Add("Description", "Dimension5");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

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
            oUserObjectMD = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD) as SAPbobsCOM.UserObjectsMD;
            oUDOFind = oUserObjectMD.FindColumns;
            oUDOForm = oUserObjectMD.FormColumns;
            oUDOChildTables = oUserObjectMD.ChildTables;

            var retval = oUserObjectMD.GetByKey("UDO_F_BDOSFUCN_D");

            if (!retval)
            {
                oUserObjectMD.Code = "UDO_F_BDOSFUCN_D";
                oUserObjectMD.Name = "Fuel Consumption Act";
                oUserObjectMD.TableName = "BDOSFUCN";
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
                oUDOFind.ColumnDescription = "Document Internal ID";
                oUDOFind.Add();
                oUDOFind.ColumnAlias = "U_DocDate";
                oUDOFind.ColumnDescription = "Posting Date";
                oUDOFind.Add();
                oUDOFind.ColumnAlias = "U_DateFrom";
                oUDOFind.ColumnDescription = "Date From";
                oUDOFind.Add();
                oUDOFind.ColumnAlias = "U_DateTo";
                oUDOFind.ColumnDescription = "Date To";
                oUDOFind.Add();
                oUDOFind.ColumnAlias = "U_PrjCode";
                oUDOFind.ColumnDescription = "Project Code";
                oUDOFind.Add();
                oUDOFind.ColumnAlias = "U_FuNrCode";
                oUDOFind.ColumnDescription = "Specification of Fuel Norm Code";
                oUDOFind.Add();

                //Form
                oUDOForm.FormColumnAlias = "DocEntry";
                oUDOForm.FormColumnDescription = "Document Internal ID";
                oUDOForm.Editable = SAPbobsCOM.BoYesNoEnum.tYES;
                oUDOForm.Add();               

                oUDOChildTables.Add();
                oUDOChildTables.SetCurrentLine(oUDOChildTables.Count - 1);
                oUDOChildTables.TableName = "BDOSFUC1";
                oUDOChildTables.ObjectName = "BDOSFUC1";

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
                    oCreationPackage.UniqueID = "UDO_F_BDOSFUCN_D";
                    oCreationPackage.String = BDOSResources.getTranslate("FuelConsumptionAct");
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

                    //if (pVal.ItemUID == "AssetMTR" && !pVal.BeforeAction)
                    //    addMatrixRow(oForm);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && !pVal.BeforeAction)
                {
                    if (!pVal.InnerEvent && (pVal.ItemUID == "CriteriaOB" || pVal.ItemUID == "FixedOB"))
                    {
                        if (oForm.DataSources.UserDataSources.Item("CriteriaOB").ValueEx == "1")
                            oForm.DataSources.DBDataSources.Item("@BDOSFUCN").SetValue("U_Fixed", 0, "N");
                        else if (oForm.DataSources.UserDataSources.Item("CriteriaOB").ValueEx == "2")
                            oForm.DataSources.DBDataSources.Item("@BDOSFUCN").SetValue("U_Fixed", 0, "Y");

                        oForm.Items.Item("PrjCodeE").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        setVisibleFormItems(oForm);
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
                setVisibleFormItems(oForm);
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

            top += (height + 1);

            FormsB1.addChooseFromList(oForm, false, "63", "ProjectCodeCFL"); //Project Codes
            FormsB1.addChooseFromList(oForm, false, "UDO_F_BDOSFUCN_D", "SpecificationOfFuelNormCodeCFL"); //Specification of Fuel Norm

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
            formItems.Add("TableName", "@BDOSFUCN");
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
            itemName = "FuNrCodeS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("SpecificationOfFuelNorm"));
            formItems.Add("LinkTo", "FuNrCodeE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }

            formItems = new Dictionary<string, object>();
            itemName = "FuNrCodeE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSFUCN");
            formItems.Add("Alias", "U_FuNrCode");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("ChooseFromListUID", "SpecificationOfFuelNormCodeCFL");
            formItems.Add("ChooseFromListAlias", "Code");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }

            formItems = new Dictionary<string, object>();
            itemName = "FuNrLB"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            formItems.Add("Left", left_e - 20);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("LinkTo", "FuNrCodeE");
            formItems.Add("LinkedObjectType", "UDO_F_BDOSFUCN_D"); //Specification of Fuel Norm

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }

            //top = top + height + 1;

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
                return;
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
                return;
            }

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "AssetMTR"; //10 characters
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
                return;
            }

            SAPbouiCOM.DBDataSource oDBDataSource = oForm.DataSources.DBDataSources.Add("@BDOSFUC1");

            FormsB1.addChooseFromList(oForm, false, "4", "ItemCodeCFL"); //Items
            FormsB1.addChooseFromList(oForm, false, "UDO_F_BDOSFUTP_D", "FuelTypeCodeCFL"); //Fuel Types
            FormsB1.addChooseFromList(oForm, false, "4", "FuelCodeCFL"); //Items

            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("AssetMTR").Specific));
            oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
            SAPbouiCOM.Columns oColumns = oMatrix.Columns;

            SAPbouiCOM.Column oColumn = oColumns.Add("LineID", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "#";
            oColumn.Editable = false;

            oColumn.DataBind.SetBound(true, "@BDOSFUC1", "LineId");

            oColumn = oColumns.Add("ItemCode", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("ItemCode");

            oColumn.DataBind.SetBound(true, "@BDOSFUC1", "U_ItemCode");

            oColumn.ChooseFromListUID = "ItemCodeCFL";
            oColumn.ChooseFromListAlias = "Code";
            SAPbouiCOM.LinkedButton oLink = oColumn.ExtendedObject;
            oLink.LinkedObjectType = "4"; //Items

            oColumn = oColumns.Add("ItemName", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("ItemName");
            oColumn.Editable = false;

            oColumn.DataBind.SetBound(true, "@BDOSFUC1", "U_ItemName");

            oColumn = oColumns.Add("FuTpCode", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("FuelTypeCode");

            oColumn.DataBind.SetBound(true, "@BDOSFUC1", "U_FuTpCode");

            oColumn.ChooseFromListUID = "FuelTypeCodeCFL";
            oColumn.ChooseFromListAlias = "Code";
            oLink = oColumn.ExtendedObject;
            oLink.LinkedObjectType = "UDO_F_BDOSFUTP_D"; //Fuel Types

            oColumn = oColumns.Add("FuelCode", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("FuelCode");

            oColumn.DataBind.SetBound(true, "@BDOSFUC1", "U_FuelCode");

            oColumn.ChooseFromListUID = "FuelCodeCFL";
            oColumn.ChooseFromListAlias = "Code";
            oLink = oColumn.ExtendedObject;
            oLink.LinkedObjectType = "4"; //Items

            oColumn = oColumns.Add("FuUomEntry", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("UomEntry");

            oColumn.DataBind.SetBound(true, "@BDOSFUC1", "U_FuUomEntry");

            oLink = oColumn.ExtendedObject;
            oLink.LinkedObjectType = "10000199"; //UoM Master Data

            oColumn = oColumns.Add("FuPerKm", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("PerKm");

            oColumn.DataBind.SetBound(true, "@BDOSFUC1", "U_FuPerKm");

            oColumn = oColumns.Add("FuPerHr", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("PerHr");

            oColumn.DataBind.SetBound(true, "@BDOSFUC1", "U_FuPerHr");

            oColumn = oColumns.Add("OdmtrStart", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("StartingValueOfOdometer");

            oColumn.DataBind.SetBound(true, "@BDOSFUC1", "U_OdmtrStart");

            oColumn = oColumns.Add("OdmtrEnd", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("EndingValueOfOdometer");

            oColumn.DataBind.SetBound(true, "@BDOSFUC1", "U_OdmtrEnd");

            oColumn = oColumns.Add("HrsWorked", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("HoursWorked");

            oColumn.DataBind.SetBound(true, "@BDOSFUC1", "U_HrsWorked");

            oColumn = oColumns.Add("NormCn", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("NormConsumption");

            oColumn.DataBind.SetBound(true, "@BDOSFUC1", "U_NormCn");

            oColumn = oColumns.Add("ActuallyCn", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("ActuallyConsumption");

            oColumn.DataBind.SetBound(true, "@BDOSFUC1", "U_ActuallyCn");

            oMatrix.Clear();
            oDBDataSource.Query();
            oMatrix.LoadFromDataSource();

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

                            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("AssetMTR").Specific));
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

                SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("AssetMTR").Specific));
                int mtrWidth = oForm.ClientWidth / 3 * 2;
                oForm.Items.Item("AssetMTR").Width = mtrWidth;
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
                //bool isFixed = oForm.DataSources.DBDataSources.Item("@BDOSFUCN").GetValue("U_Fixed", 0) == "Y";
                //oForm.Items.Item("PerKmS").Visible = isFixed;
                //oForm.Items.Item("PerKmE").Visible = isFixed;
                //oForm.Items.Item("PerHrS").Visible = isFixed;
                //oForm.Items.Item("PerHrE").Visible = isFixed;
                //oForm.Items.Item("addMTRB").Visible = !isFixed;
                //oForm.Items.Item("delMTRB").Visible = !isFixed;
                //oForm.Items.Item("AssetMTR").Visible = !isFixed;
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
                SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("AssetMTR").Specific));

                int index = 0;
                if (oMatrix.RowCount == 0)
                {
                    index = 1;
                }
                else
                {
                    index = Convert.ToInt32(oMatrix.Columns.Item("LineID").Cells.Item(oMatrix.RowCount).Specific.Value) + 1;
                }

                oMatrix.AddRow(1, -1);
                int row = oMatrix.RowCount;

                LanguageUtils.IgnoreErrors<string>(() => oMatrix.Columns.Item("LineID").Cells.Item(row).Specific.Value = index.ToString());
                LanguageUtils.IgnoreErrors<string>(() => oMatrix.Columns.Item("CrtrCode").Cells.Item(row).Specific.Value = "");
                LanguageUtils.IgnoreErrors<string>(() => oMatrix.Columns.Item("CrtrName").Cells.Item(row).Specific.Value = "");
                LanguageUtils.IgnoreErrors<string>(() => oMatrix.Columns.Item("CrtrValue").Cells.Item(row).Specific.Value = "");
                LanguageUtils.IgnoreErrors<string>(() => oMatrix.Columns.Item("CrtrPr").Cells.Item(row).Specific.Value = "");
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
                SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("AssetMTR").Specific));
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
                        oForm.DataSources.DBDataSources.Item("@BDOSFUC1").RemoveRecord(row - deletedRowCount);
                        firstRow = row;
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
    }
}
