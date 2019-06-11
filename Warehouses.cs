using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BDO_Localisation_AddOn
{
    class Warehouses
    {
        public static void createUserFields(out string errorText)
        {
            errorText = null;
            Dictionary<string, object> fieldskeysMap;
            Dictionary<string, string> listValidValuesDict;

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSPrjCod");
            fieldskeysMap.Add("TableName", "OWHS");
            fieldskeysMap.Add("Description", "Project");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            listValidValuesDict = new Dictionary<string, string>();
            listValidValuesDict.Add("Project", "Project");
            listValidValuesDict.Add("WriteOff", "Write Off");
            listValidValuesDict.Add("Main", "Main");
            listValidValuesDict.Add("Fuel", "Fuel");
            listValidValuesDict.Add("Other", "Other");

            fieldskeysMap.Add("Name", "BDOSWhType");
            fieldskeysMap.Add("TableName", "OWHS");
            fieldskeysMap.Add("Description", "Warehouse Type");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);
            fieldskeysMap.Add("ValidValues", listValidValuesDict);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSRespID");
            fieldskeysMap.Add("TableName", "OWHS");
            fieldskeysMap.Add("Description", "Responsible ID");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 20);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSRespN");
            fieldskeysMap.Add("TableName", "OWHS");
            fieldskeysMap.Add("Description", "Responsible Name");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 100);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            GC.Collect();
        }

        public static void createFormItems(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            Dictionary<string, object> formItems;
            Dictionary<string, string> listValidValuesDict = null;
            string itemName = "";

            SAPbouiCOM.Item oItem = oForm.Items.Item("42");
            int pane = 1;
            int height = oItem.Height;
            double top = oItem.Top;

            int left_s = oItem.Left;
            int width_s = oItem.Width;

            oItem = oForm.Items.Item("41");
            int left_e = oItem.Left;
            int width_e = oItem.Width;

            top = top + height + 5;

            formItems = new Dictionary<string, object>();
            itemName = "BDOSPrjS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Project"));
            formItems.Add("LinkTo", "BDOSPrjE");
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            string objectType = "63";
            string uniqueID_lf_Project = "Project_CFL";
            FormsB1.addChooseFromList(oForm, false, objectType, uniqueID_lf_Project);

            formItems = new Dictionary<string, object>();
            itemName = "BDOSPrjE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OWHS");
            formItems.Add("Alias", "U_BDOSPrjCod");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("ChooseFromListUID", uniqueID_lf_Project);
            formItems.Add("ChooseFromListAlias", "PrjCode");
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "BDOSPrjLB"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            formItems.Add("Left", left_e - 20);
            formItems.Add("Top", top);
            formItems.Add("Height", 14);
            formItems.Add("UID", itemName);
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);
            formItems.Add("LinkTo", "BDOSPrjE");
            formItems.Add("LinkedObjectType", objectType);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 5;

            SAPbobsCOM.Warehouses oWarehouses = null;
            SAPbobsCOM.ValidValues oValidValues = null;
            SAPbobsCOM.Fields oFields = null;

            oWarehouses = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oWarehouses);
            oFields = oWarehouses.UserFields.Fields;
            oValidValues = oFields.Item("U_BDOSWhType").ValidValues;

            listValidValuesDict = new Dictionary<string, string>();
            for (int i = 0; i < oValidValues.Count; i++)
            {
                string value = oValidValues.Item(i).Value;
                listValidValuesDict.Add(value, BDOSResources.getTranslate(value));
            }

            formItems = new Dictionary<string, object>();
            itemName = "BDOSWhTpS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("WarehouseType"));
            formItems.Add("LinkTo", "BDOSWhTpE");
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "BDOSWhTpE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OWHS");
            formItems.Add("Alias", "U_BDOSWhType");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
            formItems.Add("DisplayDesc", true);
            formItems.Add("ValidValues", listValidValuesDict);
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 5;

            formItems = new Dictionary<string, object>();
            itemName = "BDOSRespS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Responsible"));
            formItems.Add("LinkTo", "BDOSRespE");
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            objectType = "171";
            string uniqueID_lf_Empl = "Empl_CFL";
            FormsB1.addChooseFromList(oForm, false, objectType, uniqueID_lf_Empl);

            formItems = new Dictionary<string, object>();
            itemName = "BDOSRespE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OWHS");
            formItems.Add("Alias", "U_BDOSRespID");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", 30);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("ChooseFromListUID", uniqueID_lf_Empl);
            formItems.Add("ChooseFromListAlias", "empID");
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "BDOSRespNE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OWHS");
            formItems.Add("Alias", "U_BDOSRespN");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e + 35);
            formItems.Add("Width", width_e - 35);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);
            formItems.Add("Enabled", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "BDOSRespLB"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            formItems.Add("Left", left_e - 20);
            formItems.Add("Top", top);
            formItems.Add("Height", 14);
            formItems.Add("UID", itemName);
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);
            formItems.Add("LinkTo", "BDOSRespE");
            formItems.Add("LinkedObjectType", objectType);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            GC.Collect();
        }

        public static void chooseFromList(SAPbouiCOM.Form oForm, SAPbouiCOM.IChooseFromListEvent oCFLEvento, SAPbouiCOM.ItemEvent pVal, out string errorText)
        {
            errorText = null;

            bool beforeAction = pVal.BeforeAction;
            string sCFL_ID = oCFLEvento.ChooseFromListUID;
            try
            {
                if (beforeAction == false)
                {
                    SAPbouiCOM.DataTable oDataTable = null;
                    oDataTable = oCFLEvento.SelectedObjects;

                    if (oDataTable != null)
                    {
                        if (sCFL_ID == "Project_CFL")
                        {
                            string prjCode = Convert.ToString(oDataTable.GetValue("PrjCode", 0));

                            try
                            {
                                SAPbouiCOM.EditText oEdit = oForm.Items.Item("BDOSPrjE").Specific;
                                oEdit.Value = prjCode;
                            }
                            catch { }
                            finally
                            {
                                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                                {
                                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                                }
                            }
                        }
                        else if (sCFL_ID == "Empl_CFL")
                        {
                            string empID = Convert.ToString(oDataTable.GetValue("empID", 0));
                            string empName = oDataTable.GetValue("firstName", 0) + " " + oDataTable.GetValue("lastName", 0);

                            try
                            {
                                SAPbouiCOM.EditText oEdit = oForm.Items.Item("BDOSRespE").Specific;
                                oEdit.Value = empID;
                            }
                            catch { }

                            try
                            {
                                SAPbouiCOM.EditText oEdit = oForm.Items.Item("BDOSRespNE").Specific;
                                oEdit.Value = empName;
                            }
                            catch { }

                            setVisibleFormItems(oForm, out errorText);
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            {
                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                            }

                        }
                    }
                }
            }
            catch (Exception ex)
            {
                int errCode;
                string errMsg;

                Program.oCompany.GetLastError(out errCode, out errMsg);
                errorText = BDOSResources.getTranslate("ErrorDescription") + " : " + errMsg + "! " + BDOSResources.getTranslate("Code") + " : " + errCode + "! " + BDOSResources.getTranslate("OtherInfo") + " : " + ex.Message;
            }
            finally
            {
                GC.Collect();
            }
        }

        public static void uiApp_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD & pVal.BeforeAction == true)
                {
                    createFormItems(oForm, out errorText);
                    Program.FORM_LOAD_FOR_VISIBLE = true;
                    Program.FORM_LOAD_FOR_ACTIVATE = true;
                    formDataLoad(oForm, out errorText);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE & pVal.BeforeAction == false)
                {
                    if (Program.FORM_LOAD_FOR_ACTIVATE == true)
                    {
                        setVisibleFormItems(oForm, out errorText); ;
                        Program.FORM_LOAD_FOR_ACTIVATE = false;
                    }
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                {
                    SAPbouiCOM.ChooseFromListEvent oCFLEvento = null;
                    oCFLEvento = ((SAPbouiCOM.ChooseFromListEvent)(pVal));
                    chooseFromList(oForm, oCFLEvento, pVal, out errorText);
                }

                if ((pVal.ItemUID == "BDOSRespE") && pVal.ItemChanged == true && pVal.ActionSuccess == true)
                {
                    oForm.Freeze(true);
                    if (oForm.DataSources.DBDataSources.Item("OWHS").GetValue("U_BDOSRespID", 0) == "")
                    {
                        SAPbouiCOM.EditText oEdit = oForm.Items.Item("BDOSRespNE").Specific;
                        oEdit.Value = "";
                    }
                    oForm.Freeze(false);
                }

                //if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_RESIZE && pVal.BeforeAction == false)
                //{
                //    oForm.Freeze(true);
                //    resizeForm(oForm, out errorText);
                //    oForm.Freeze(false);
                //}
            }
        }

        public static void uiApp_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(BusinessObjectInfo.FormTypeEx, Program.currentFormCount);

            if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD & BusinessObjectInfo.BeforeAction == false)
            {
                formDataLoad(oForm, out errorText);
            }
        }

        public static void formDataLoad(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            try
            {
                setVisibleFormItems(oForm, out errorText);
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

        public static void setVisibleFormItems(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            oForm.Freeze(true);

            try
            {
                oForm.Items.Item("6").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                oForm.Items.Item("BDOSRespNE").Enabled = false;
            }
            catch { }

            oForm.Freeze(false);
        }
    }
}
