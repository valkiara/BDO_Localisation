using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace BDO_Localisation_AddOn
{
    class DocumentSettings
    {
        public static void createUserFields(out string errorText)
        {
            errorText = null;
            Dictionary<string, object> fieldskeysMap;
            Dictionary<string, string> listValidValuesDict;

            fieldskeysMap = new Dictionary<string, object>();
            listValidValuesDict = new Dictionary<string, string>();
            listValidValuesDict.Add("DoNotBlock", "Do Not Block");
            listValidValuesDict.Add("ByCompany", "By Company");
            listValidValuesDict.Add("ByWarehouse", "By Warehouse");

            fieldskeysMap.Add("Name", "BDOSBlcPDt"); //უარყოფითი ნაშთების კონტროლი დოკ.თარიღის მიხედვით
            fieldskeysMap.Add("TableName", "OADM");
            fieldskeysMap.Add("Description", "Block Negative Stock On Posting Date");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);
            fieldskeysMap.Add("ValidValues", listValidValuesDict);
            fieldskeysMap.Add("DefaultValue", "DoNotBlock");

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            GC.Collect();
        }

        public static void createFormItems(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            Dictionary<string, object> formItems;
            Dictionary<string, string> listValidValuesDict = null;
            string itemName = "";

            SAPbouiCOM.Item oItem = oForm.Items.Item("25400292");
            int left_s = oItem.Left;            
            int height = oItem.Height;
            int top = oForm.Items.Item("51").Top;
            int width_s = oItem.Width;
            int pane = oItem.FromPane;

            oItem = oForm.Items.Item("242000001");
            int left_e = oItem.Left;
            int width_e = oItem.Width;

            top = top + height + 5;

            formItems = new Dictionary<string, object>();
            itemName = "BDOSBlcPDS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("BlockNegativeStockOnPostingDate"));
            formItems.Add("LinkTo", "BDOSBlcPDt");
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            listValidValuesDict = new Dictionary<string, string>();
            listValidValuesDict.Add("DoNotBlock", BDOSResources.getTranslate("DoNotBlock"));
            listValidValuesDict.Add("ByCompany", BDOSResources.getTranslate("ByCompany"));
            listValidValuesDict.Add("ByWarehouse", BDOSResources.getTranslate("ByWarehouse"));

            formItems = new Dictionary<string, object>();
            itemName = "BDOSBlcPDt"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OADM");
            formItems.Add("Alias", "U_BDOSBlcPDt");
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

            GC.Collect();
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
                }
            }
        }
    }
}
