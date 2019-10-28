using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
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

            //საწვავის ერთეული
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "FuUomCode");
            fieldskeysMap.Add("TableName", "BDOSFUC1");
            fieldskeysMap.Add("Description", "Fuel UoM Code");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 20);

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

            oUserObjectMD = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD);

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
                oUDOFind.ColumnAlias = "Remark";
                oUDOFind.ColumnDescription = "Remark";
                oUDOFind.Add();

                //Form
                oUDOForm.FormColumnAlias = "DocEntry";
                oUDOForm.FormColumnDescription = "Internal Number";
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
                        oForm.Title = BDOSResources.getTranslate("FuelConsumptionAct");
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

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                {
                    SAPbouiCOM.IChooseFromListEvent oCFLEvento = ((SAPbouiCOM.IChooseFromListEvent)(pVal));
                    chooseFromList(oForm, pVal, oCFLEvento);

                    //if (pVal.ItemUID == "AssetMTR" && !pVal.BeforeAction)
                    //    addMatrixRow(oForm);
                }

                if (pVal.ItemUID == "AssetMTR" && pVal.ItemChanged)
                {
                    if (pVal.ColUID == "OdmtrStart" || pVal.ColUID == "OdmtrEnd" || pVal.ColUID == "HrsWorked")
                    {
                        calculateConsumptionValue(oForm, pVal.Row - 1);
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
            int left_e = 160;
            int height = 15;
            int top = 5;
            int width_s = 139;
            int width_e = 140;

            int left_s2 = 300;
            int left_e2 = left_s2 + 121;
            int top2 = 5;

            top += (height + 1);

            FormsB1.addChooseFromList(oForm, false, "63", "ProjectCodeCFL"); //Project Codes
            FormsB1.addChooseFromList(oForm, false, "UDO_F_BDOSFUNR_D", "SpecificationOfFuelNormCodeCFL"); //Specification of Fuel Norm

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
            formItems.Add("LinkedObjectType", "UDO_F_BDOSFUNR_D"); //Specification of Fuel Norm

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "DateFromS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("DateFrom"));
            formItems.Add("LinkTo", "DateFromE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }

            formItems = new Dictionary<string, object>();
            itemName = "DateFromE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSFUCN");
            formItems.Add("Alias", "U_DateFrom");
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
            itemName = "DateToS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("DateTo"));
            formItems.Add("LinkTo", "DateToE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }

            formItems = new Dictionary<string, object>();
            itemName = "DateToE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSFUCN");
            formItems.Add("Alias", "U_DateTo");
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

            formItems = new Dictionary<string, object>();
            itemName = "No.S"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s2);
            formItems.Add("Width", width_s / 3);
            formItems.Add("Top", top2);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Number"));
            formItems.Add("LinkTo", "SeriesC");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }

            formItems = new Dictionary<string, object>();
            itemName = "SeriesC"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSFUCN");
            formItems.Add("Alias", "Series");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            formItems.Add("Left", left_s2 + width_s / 3);
            formItems.Add("Width", width_s * 2 / 4);
            formItems.Add("Top", top2);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
            formItems.Add("Description", BDOSResources.getTranslate("Series"));

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
            }

            formItems = new Dictionary<string, object>();
            itemName = "DocNumE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSFUCN");
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
            oForm.Items.Item("DocNumE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);

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
            formItems.Add("TableName", "@BDOSFUCN");
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
            oForm.Items.Item("StatusC").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);

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
            formItems.Add("TableName", "@BDOSFUCN");
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
            oForm.Items.Item("CanceledC").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);

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
            formItems.Add("TableName", "@BDOSFUCN");
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
            oForm.Items.Item("CreateDatE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);

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
            formItems.Add("TableName", "@BDOSFUCN");
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
            oForm.Items.Item("UpdateDatE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);

            top2 += height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "DocDatS"; //10 characters
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
            formItems.Add("TableName", "@BDOSFUCN");
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
            //oForm.Items.Item("DocDateE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);

            top += (4 * height + 1);

            formItems = new Dictionary<string, object>();
            itemName = "addMTRB"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            formItems.Add("Left", left_s);
            formItems.Add("Width", 70);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Add"));

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

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
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

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                throw new Exception(errorText);
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
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("AssetCode");
            oColumn.DataBind.SetBound(true, "@BDOSFUC1", "U_ItemCode");
            oColumn.ChooseFromListUID = "ItemCodeCFL";
            oColumn.ChooseFromListAlias = "ItemCode";
            SAPbouiCOM.LinkedButton oLink = oColumn.ExtendedObject;
            oLink.LinkedObjectType = "4"; //Items

            oColumn = oColumns.Add("ItemName", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("AssetName");
            oColumn.Editable = false;
            oColumn.DataBind.SetBound(true, "@BDOSFUC1", "U_ItemName");

            //-------------------------------------------Fuel Types--------------------------------------            
            oColumn = oColumns.Add("FuTpCode", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("FuelType");
            oColumn.Editable = false;
            oColumn.DataBind.SetBound(true, "@BDOSFUC1", "U_FuTpCode");
            oColumn.ChooseFromListUID = "FuelTypeCodeCFL";
            oColumn.ChooseFromListAlias = "Code";
            oLink = oColumn.ExtendedObject;
            oLink.LinkedObjectType = "UDO_F_BDOSFUTP_D"; //Fuel Types

            oColumn = oColumns.Add("FuelCode", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Fuel");
            oColumn.Editable = false;
            oColumn.DataBind.SetBound(true, "@BDOSFUC1", "U_FuelCode");
            oColumn.ChooseFromListUID = "FuelCodeCFL";
            oColumn.ChooseFromListAlias = "ItemCode";
            oLink = oColumn.ExtendedObject;
            oLink.LinkedObjectType = "4"; //Items

            oColumn = oColumns.Add("FuUomEntry", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("UomEntry");
            oColumn.Editable = false;
            oColumn.Visible = false;
            oColumn.DataBind.SetBound(true, "@BDOSFUC1", "U_FuUomEntry");
            oLink = oColumn.ExtendedObject;
            oLink.LinkedObjectType = "10000199"; //UoM Master Data

            oColumn = oColumns.Add("FuUomCode", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("UomCode");
            oColumn.Editable = false;
            oColumn.DataBind.SetBound(true, "@BDOSFUC1", "U_FuUomCode");

            oColumn = oColumns.Add("FuPerKm", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("PerKm");
            oColumn.Editable = false;
            oColumn.DataBind.SetBound(true, "@BDOSFUC1", "U_FuPerKm");

            oColumn = oColumns.Add("FuPerHr", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("PerHr");
            oColumn.Editable = false;
            oColumn.DataBind.SetBound(true, "@BDOSFUC1", "U_FuPerHr");
            //-------------------------------------------Fuel Types--------------------------------------

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
            oColumn.Editable = false;
            oColumn.DataBind.SetBound(true, "@BDOSFUC1", "U_NormCn");

            oColumn = oColumns.Add("ActuallyCn", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("ActuallyConsumption");
            oColumn.DataBind.SetBound(true, "@BDOSFUC1", "U_ActuallyCn");

            for (int i = 1; i <= 5; i++)
            {
                FormsB1.addChooseFromList(oForm, false, "62", "Dimension" + i + "CFL");

                oColumn = oColumns.Add("Dimension" + i, SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                oColumn.TitleObject.Caption = BDOSResources.getTranslate("Dimension" + i);
                oColumn.DataBind.SetBound(true, "@BDOSFUC1", "U_Dimension" + i);
                oColumn.ChooseFromListUID = "Dimension" + i + "CFL";
                oColumn.ChooseFromListAlias = "OcrCode";
                oLink = oColumn.ExtendedObject;
                oLink.LinkedObjectType = "62"; //Cost Rate
            }

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
                    else if (oCFLEvento.ChooseFromListUID == "SpecificationOfFuelNormCodeCFL")
                    {
                        SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID);
                        SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();

                        SAPbouiCOM.DBDataSource oDBDataSource = oForm.DataSources.DBDataSources.Item("@BDOSFUCN");
                        string prjCode = oDBDataSource.GetValue("U_PrjCode", 0);
                        if (!string.IsNullOrEmpty(prjCode))
                        {
                            SAPbouiCOM.Condition oCon = oCons.Add();
                            oCon.Alias = "U_PrjCode";
                            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                            oCon.CondVal = prjCode;
                            oCFL.SetConditions(oCons);
                        }
                    }
                    else if (oCFLEvento.ChooseFromListUID == "ItemCodeCFL")
                    {
                        SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID);
                        SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();

                        SAPbouiCOM.Condition oCon = oCons.Add();
                        oCon.Alias = "ItemType";
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = "F";
                        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                        oCon = oCons.Add();
                        oCon.Alias = "validFor";
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = "Y";
                        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                        SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        oRecordSet.DoQuery(@"SELECT ""Code"" FROM ""OACS"" WHERE ""U_BDOSVhcle"" = 'Y'");
                        int recordCount = oRecordSet.RecordCount;
                        int i = 0;

                        while (!oRecordSet.EoF)
                        {
                            string assetClassCode = oRecordSet.Fields.Item("Code").Value;
                            oCon = oCons.Add();
                            oCon.Alias = "AssetClass";
                            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                            oCon.CondVal = assetClassCode;
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
                        SAPbouiCOM.DBDataSource oDBDataSource = oForm.DataSources.DBDataSources.Item("@BDOSFUCN");

                        if (oCFLEvento.ChooseFromListUID == "ProjectCodeCFL")
                        {
                            string prjCode = oDataTable.GetValue("PrjCode", 0);
                            LanguageUtils.IgnoreErrors<string>(() => oForm.Items.Item("PrjCodeE").Specific.Value = prjCode);
                        }
                        else if (oCFLEvento.ChooseFromListUID == "SpecificationOfFuelNormCodeCFL")
                        {
                            string specificationOfFuelNormCode = oDataTable.GetValue("Code", 0);
                            string prjCode = oDataTable.GetValue("U_PrjCode", 0);

                            LanguageUtils.IgnoreErrors<string>(() => oForm.Items.Item("FuNrCodeE").Specific.Value = specificationOfFuelNormCode);
                            LanguageUtils.IgnoreErrors<string>(() => oForm.Items.Item("PrjCodeE").Specific.Value = prjCode);

                            updatePerKmHrValue(oForm);
                            calculateConsumptionValue(oForm);
                        }
                        else if (oCFLEvento.ChooseFromListUID == "ItemCodeCFL")
                        {
                            string itemCode = oDataTable.GetValue("ItemCode", 0);
                            string docEntryStr = oDBDataSource.GetValue("DocEntry", 0);
                            int docEntry = 0;

                            if (!string.IsNullOrEmpty(docEntryStr))
                                docEntry = Convert.ToInt32(oDBDataSource.GetValue("DocEntry", 0));

                            SAPbobsCOM.Recordset oRecordSet = getFuelType(itemCode);

                            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("AssetMTR").Specific;

                            LanguageUtils.IgnoreErrors<string>(() => oMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value = itemCode);
                            if (oRecordSet != null)
                            {
                                LanguageUtils.IgnoreErrors<string>(() => oMatrix.Columns.Item("ItemName").Cells.Item(pVal.Row).Specific.Value = oRecordSet.Fields.Item("ItemName").Value);
                                LanguageUtils.IgnoreErrors<string>(() => oMatrix.Columns.Item("FuTpCode").Cells.Item(pVal.Row).Specific.Value = oRecordSet.Fields.Item("Code").Value);
                                LanguageUtils.IgnoreErrors<string>(() => oMatrix.Columns.Item("FuelCode").Cells.Item(pVal.Row).Specific.Value = oRecordSet.Fields.Item("U_ItemCode").Value);
                                LanguageUtils.IgnoreErrors<string>(() => oMatrix.Columns.Item("FuUomEntry").Cells.Item(pVal.Row).Specific.Value = oRecordSet.Fields.Item("U_UomEntry").Value.ToString());
                                LanguageUtils.IgnoreErrors<string>(() => oMatrix.Columns.Item("FuUomCode").Cells.Item(pVal.Row).Specific.Value = oRecordSet.Fields.Item("U_UomCode").Value);
                                //LanguageUtils.IgnoreErrors<string>(() => oMatrix.Columns.Item("FuPerKm").Cells.Item(pVal.Row).Specific.Value = oRecordSet.Fields.Item("U_PerKm").Value.ToString());
                                //LanguageUtils.IgnoreErrors<string>(() => oMatrix.Columns.Item("FuPerHr").Cells.Item(pVal.Row).Specific.Value = oRecordSet.Fields.Item("U_PerHr").Value.ToString());
                                LanguageUtils.IgnoreErrors<string>(() => oMatrix.Columns.Item("FuUomCode").Cells.Item(pVal.Row).Specific.Value = oRecordSet.Fields.Item("U_UomCode").Value);
                                LanguageUtils.IgnoreErrors<string>(() => oMatrix.Columns.Item("OdmtrStart").Cells.Item(pVal.Row).Specific.Value = Convert.ToString(getOdmtrStart(itemCode, docEntry)));
                            }
                            updatePerKmHrValue(oForm, pVal.Row - 1);
                            calculateConsumptionValue(oForm, pVal.Row - 1);
                        }

                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
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
                oForm.ClientHeight = Program.uiApp.Desktop.Height / 4;
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
                int left_e = 160;
                oForm.Items.Item("0_U_E").Left = left_e;
                oForm.Items.Item("0_U_E").Width = 140;
                oForm.Items.Item("1").Top = oForm.ClientHeight - 25;
                oForm.Items.Item("2").Top = oForm.ClientHeight - 25;

                SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("AssetMTR").Specific));
                int mtrWidth = oForm.ClientWidth;
                oForm.Items.Item("AssetMTR").Width = mtrWidth;
                oForm.Items.Item("AssetMTR").Height = oForm.ClientHeight / 2;
                oMatrix.Columns.Item("LineID").Width = 19;
                mtrWidth -= 19;
                mtrWidth /= 17;
                oMatrix.Columns.Item("ItemCode").Width = mtrWidth;
                oMatrix.Columns.Item("ItemName").Width = mtrWidth;
                oMatrix.Columns.Item("FuTpCode").Width = mtrWidth;
                oMatrix.Columns.Item("FuelCode").Width = mtrWidth;
                //oMatrix.Columns.Item("FuUomEntry").Width = mtrWidth;
                oMatrix.Columns.Item("FuUomCode").Width = mtrWidth;
                oMatrix.Columns.Item("FuPerKm").Width = mtrWidth;
                oMatrix.Columns.Item("FuPerHr").Width = mtrWidth;
                oMatrix.Columns.Item("OdmtrStart").Width = mtrWidth;
                oMatrix.Columns.Item("OdmtrEnd").Width = mtrWidth;
                oMatrix.Columns.Item("HrsWorked").Width = mtrWidth;
                oMatrix.Columns.Item("NormCn").Width = mtrWidth;
                oMatrix.Columns.Item("ActuallyCn").Width = mtrWidth;
                oMatrix.Columns.Item("Dimension1").Width = mtrWidth;
                oMatrix.Columns.Item("Dimension2").Width = mtrWidth;
                oMatrix.Columns.Item("Dimension3").Width = mtrWidth;
                oMatrix.Columns.Item("Dimension4").Width = mtrWidth;
                oMatrix.Columns.Item("Dimension5").Width = mtrWidth;
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
                //LanguageUtils.IgnoreErrors<string>(() => oMatrix.Columns.Item("CrtrCode").Cells.Item(row).Specific.Value = "");
                //LanguageUtils.IgnoreErrors<string>(() => oMatrix.Columns.Item("CrtrName").Cells.Item(row).Specific.Value = "");
                //LanguageUtils.IgnoreErrors<string>(() => oMatrix.Columns.Item("CrtrValue").Cells.Item(row).Specific.Value = "");
                //LanguageUtils.IgnoreErrors<string>(() => oMatrix.Columns.Item("CrtrPr").Cells.Item(row).Specific.Value = "");
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

        private static void calculateConsumptionValue(SAPbouiCOM.Form oForm, int rowIndex = -1)
        {
            try
            {
                oForm.Freeze(true);

                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("AssetMTR").Specific;
                oMatrix.FlushToDataSource();

                SAPbouiCOM.DBDataSource oDBDataSourceMTR = oForm.DataSources.DBDataSources.Item("@BDOSFUC1");

                int rowCount = rowIndex == -1 ? oDBDataSourceMTR.Size - 1 : rowIndex;
                int i = rowIndex == -1 ? 0 : rowIndex;

                for (; i <= rowCount; i++)
                {
                    string itemCode = oDBDataSourceMTR.GetValue("U_ItemCode", i);
                    if (!string.IsNullOrEmpty(itemCode))
                    {
                        var fuPerKm = Convert.ToDecimal(oDBDataSourceMTR.GetValue("U_FuPerKm", i));
                        var odmtrStart = Convert.ToDecimal(oDBDataSourceMTR.GetValue("U_OdmtrStart", i));
                        var odmtrEnd = Convert.ToDecimal(oDBDataSourceMTR.GetValue("U_OdmtrEnd", i));
                        var fuPerHr = Convert.ToDecimal(oDBDataSourceMTR.GetValue("U_FuPerHr", i));
                        var hrsWorked = Convert.ToDecimal(oDBDataSourceMTR.GetValue("U_HrsWorked", i));
                        decimal normConsumptionKm = fuPerKm * (odmtrEnd - odmtrStart);
                        decimal normConsumptionHr = fuPerHr * hrsWorked;
                        decimal normConsumption = normConsumptionKm + normConsumptionHr;

                        oDBDataSourceMTR.SetValue("U_NormCn", i, Convert.ToString(normConsumption));
                        oDBDataSourceMTR.SetValue("U_ActuallyCn", i, Convert.ToString(normConsumption));
                    }
                }
                oMatrix.LoadFromDataSource();
            }
            catch (Exception ex)
            {
                Program.uiApp.StatusBar.SetSystemMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        private static void updatePerKmHrValue(SAPbouiCOM.Form oForm, int rowIndex = -1)
        {
            try
            {
                oForm.Freeze(true);

                SAPbouiCOM.DBDataSource oDBDataSource = oForm.DataSources.DBDataSources.Item("@BDOSFUCN");
                string fuNrCode = oDBDataSource.GetValue("U_FuNrCode", 0);

                SAPbobsCOM.Recordset oRecordset = getFuelNormSpecification(fuNrCode);
                decimal perKm = 0;
                decimal perHr = 0;
                decimal crtrPr = 0;
                bool? isFixed = null;

                if (oRecordset != null)
                {
                    isFixed = oRecordset.Fields.Item("U_Fixed").Value == "Y";
                    perKm = Convert.ToDecimal(oRecordset.Fields.Item("U_PerKm").Value);
                    perHr = Convert.ToDecimal(oRecordset.Fields.Item("U_PerHr").Value);
                    crtrPr = Convert.ToDecimal(oRecordset.Fields.Item("U_CrtrPr").Value);
                }

                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("AssetMTR").Specific;
                oMatrix.FlushToDataSource();

                SAPbouiCOM.DBDataSource oDBDataSourceMTR = oForm.DataSources.DBDataSources.Item("@BDOSFUC1");

                int rowCount = rowIndex == -1 ? oDBDataSourceMTR.Size - 1 : rowIndex;
                int i = rowIndex == -1 ? 0 : rowIndex;

                for (; i <= rowCount; i++)
                {
                    string itemCode = oDBDataSourceMTR.GetValue("U_ItemCode", i);
                    if (!string.IsNullOrEmpty(itemCode))
                    {
                        if (!isFixed.HasValue || !isFixed.Value)
                        {
                            SAPbobsCOM.Recordset oRecordsetFuelType = getFuelType(itemCode);
                            if (oRecordsetFuelType != null)
                            {
                                perKm = Convert.ToDecimal(oRecordsetFuelType.Fields.Item("U_PerKm").Value);
                                perHr = Convert.ToDecimal(oRecordsetFuelType.Fields.Item("U_PerHr").Value);
                            }
                            if (isFixed.HasValue)
                            {
                                perKm += perKm * crtrPr;
                                perHr += perHr * crtrPr;
                            }
                        }
                        oDBDataSourceMTR.SetValue("U_FuPerKm", i, Convert.ToString(perKm));
                        oDBDataSourceMTR.SetValue("U_FuPerHr", i, Convert.ToString(perHr));
                    }
                }
                oMatrix.LoadFromDataSource();
            }
            catch (Exception ex)
            {
                Program.uiApp.StatusBar.SetSystemMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        private static SAPbobsCOM.Recordset getFuelNormSpecification(string fuNrCode)
        {
            try
            {
                SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                string query = @"SELECT 
                   ""@BDOSFUNR"".""Code"", 
                   ""@BDOSFUNR"".""Name"", 
                   ""@BDOSFUNR"".""U_PrjCode"", 
                   ""@BDOSFUNR"".""U_Fixed"", 
                   ""@BDOSFUNR"".""U_PerKm"", 
                   ""@BDOSFUNR"".""U_PerHr"",
                   SUM(""@BDOSFUN1"".""U_CrtrPr"")/100 AS ""U_CrtrPr""
                FROM ""@BDOSFUNR""
                LEFT JOIN ""@BDOSFUN1""
                ON ""@BDOSFUNR"".""Code"" = ""@BDOSFUN1"".""Code""
                WHERE ""@BDOSFUNR"".""Code"" = '" + fuNrCode + @"' 
                GROUP BY ""@BDOSFUNR"".""Code"", 
                   ""@BDOSFUNR"".""Name"", 
                   ""@BDOSFUNR"".""U_PrjCode"", 
                   ""@BDOSFUNR"".""U_Fixed"", 
                   ""@BDOSFUNR"".""U_PerKm"", 
                   ""@BDOSFUNR"".""U_PerHr""";

                oRecordSet.DoQuery(query);

                if (!oRecordSet.EoF)
                {
                    return oRecordSet;
                }
                return null;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private static SAPbobsCOM.Recordset getFuelType(string itemCode)
        {
            try
            {
                SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                string query = @"SELECT
	               ""OITM"".""ItemCode"",
                   ""OITM"".""ItemName"",
	               ""@BDOSFUTP"".""Code"",
                   ""@BDOSFUTP"".""U_ItemCode"",
	               ""@BDOSFUTP"".""U_UomEntry"",
                   ""@BDOSFUTP"".""U_UomCode"",
	               ""@BDOSFUTP"".""U_PerKm"",
	               ""@BDOSFUTP"".""U_PerHr""
                FROM ""OITM""
                LEFT JOIN ""@BDOSFUTP"" ON ""@BDOSFUTP"".""Code"" = ""OITM"".""U_BDOSFuTp""
                WHERE ""OITM"".""ItemCode"" = '" + itemCode + "'";

                oRecordSet.DoQuery(query);

                if (!oRecordSet.EoF)
                {
                    return oRecordSet;
                }
                return null;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private static double getOdmtrStart(string itemCode, int docEntry)
        {
            try
            {
                SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                StringBuilder query = new StringBuilder();
                query.Append("SELECT TOP 1 \"@BDOSFUC1\".\"U_OdmtrEnd\" \n");
                query.Append("FROM \"@BDOSFUCN\" \n");
                query.Append("INNER JOIN \"@BDOSFUC1\" \n");
                query.Append("ON \"@BDOSFUCN\".\"DocEntry\" = \"@BDOSFUC1\".\"DocEntry\" \n");
                query.Append("WHERE \"@BDOSFUCN\".\"Canceled\" = 'N' \n");
                query.Append("AND \"@BDOSFUC1\".\"U_ItemCode\" = '" + itemCode + "' \n");
                query.Append("AND \"@BDOSFUCN\".\"DocEntry\" <> " + docEntry + " \n");
                query.Append("ORDER BY \n");
                query.Append("\"@BDOSFUCN\".\"U_DocDate\" DESC, \n");
                query.Append(" \"@BDOSFUCN\".\"DocEntry\" DESC");

                oRecordSet.DoQuery(query.ToString());

                if (!oRecordSet.EoF)
                {
                    return oRecordSet.Fields.Item("U_OdmtrEnd").Value;
                }
                return 0;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
    }
}
