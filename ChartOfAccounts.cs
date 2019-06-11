using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BDO_Localisation_AddOn
{
    class ChartOfAccounts
    {
        public static void createUserFields( out string errorText)
        {
            errorText = null;
            Dictionary<string, object> fieldskeysMap;

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSItmAct");
            fieldskeysMap.Add("TableName", "OACT");
            fieldskeysMap.Add("Description", "Item Account");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 1);
            fieldskeysMap.Add("DefaultValue", "N");

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSEmpAct");
            fieldskeysMap.Add("TableName", "OACT");
            fieldskeysMap.Add("Description", "Employee Account");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 1);
            fieldskeysMap.Add("DefaultValue", "N");

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSWthCur");
            fieldskeysMap.Add("TableName", "OACT");
            fieldskeysMap.Add("Description", "Without Currency Detailing");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 1);
            fieldskeysMap.Add("DefaultValue", "N");

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            GC.Collect();
        }

        public static void uiApp_ItemEvent(  string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD & pVal.BeforeAction == true)
                {
                    createFormItems( oForm, out errorText);
                }

                if (pVal.ItemUID == "3" || pVal.ItemUID == "4" || pVal.ItemUID == "5" || pVal.ItemUID == "26" || pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE & pVal.BeforeAction == false)
                {
                    //formDataLoad( oForm, out errorText);
                    setVisibility(oForm, out errorText);
                }

            }
        }

        public static void uiApp_FormDataEvent(  ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            
            SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(BusinessObjectInfo.FormTypeEx, Program.currentFormCount);

            if (oForm.TypeEx == "804" & BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD & BusinessObjectInfo.BeforeAction == false)
            {
                //formDataLoad( oForm, out errorText);
            }
        }

        public static void createFormItems(  SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            SAPbouiCOM.Item oItem = oForm.Items.Item("49");

            int height = oItem.Height;
            int top = oItem.Top;

            oItem = oForm.Items.Item("34");
            int left = oItem.Left;
            int width = oItem.Width;

            Dictionary<string, object> formItems;
            string itemName;

            formItems = new Dictionary<string, object>();
            itemName = "BDOSItmAct";
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OACT");
            formItems.Add("Alias", "U_BDOSItmAct");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
            formItems.Add("Length", 1);
            formItems.Add("Left", left);
            formItems.Add("Width", width);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("ItemAccount"));
            formItems.Add("ValOff", "N");
            formItems.Add("ValOn", "Y");
            formItems.Add("DisplayDesc", true);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 1;

            try
            {
                SAPbouiCOM.Item oItemEmpAct = oForm.Items.Item("BDOSEmpAct");
            }
            catch
            {
            formItems = new Dictionary<string, object>();
            itemName = "BDOSEmpAct";
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OACT");
            formItems.Add("Alias", "U_BDOSEmpAct");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
            formItems.Add("Length", 1);
            formItems.Add("Left", left);
            formItems.Add("Width", width);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("EmployeeAccount"));
            formItems.Add("ValOff", "N");
            formItems.Add("ValOn", "Y");
            formItems.Add("DisplayDesc", true);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }
            }

            oItem = oForm.Items.Item("46");
            left = oItem.Left;
            width = oItem.Width;
            height = oItem.Height;
            top = oItem.Top;

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "BDOSWthCur";
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OACT");
            formItems.Add("Alias", "U_BDOSWthCur");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
            formItems.Add("Length", 1);
            formItems.Add("Left", left);
            formItems.Add("Width", width);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("WithoutCurrencyDetailing"));
            formItems.Add("ValOff", "N");
            formItems.Add("ValOn", "Y");
            formItems.Add("DisplayDesc", true);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

        }

        public static void setVisibility(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            oForm.Freeze(true);
            bool visibleProperties = oForm.DataSources.DBDataSources.Item("OACT").GetValue("Postable", 0) == "Y";

            try
            {
                oForm.Items.Item("BDOSItmAct").Visible = visibleProperties;
                oForm.Items.Item("BDOSEmpAct").Visible = visibleProperties;

                bool notLocCurr = (oForm.DataSources.DBDataSources.Item("OACT").GetValue("ActCurr", 0) != CommonFunctions.getLocalCurrency());
                oForm.Items.Item("BDOSWthCur").Visible = (visibleProperties && notLocCurr);
            }
            catch
            { }

            oForm.Freeze(false);

        }
    }
}
