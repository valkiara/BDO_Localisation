using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace BDO_Localisation_AddOn
{
    static partial class APCorrectionInvoice
    {
        public static void createFormItems(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            BDO_WBReceivedDocs.createFormItems(oForm, "ORPC", out errorText);

            Dictionary<string, object> formItems = null;

            string itemName = "";

            double height = oForm.Items.Item("10001019").Height;
            double top = oForm.Items.Item("10001019").Top+height+1;
            double left_s = oForm.Items.Item("10001018").Left;
            double left_e = oForm.Items.Item("10001019").Left;
            double width_e = oForm.Items.Item("10001018").Width;
            double width_s = oForm.Items.Item("10001019").Width;



            formItems = new Dictionary<string, object>();
            itemName = "BDO_CNTp_s"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("OperationType"));
            formItems.Add("LinkTo", "BDO_CNTp");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            List<string> listValidValues = new List<string>();
            listValidValues.Add(BDOSResources.getTranslate("Correction")); //0 //კორექტირება
            listValidValues.Add(BDOSResources.getTranslate("Return")); //1 //დაბრუნება

            formItems = new Dictionary<string, object>();
            itemName = "BDO_CNTp";
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "ORPC");
            formItems.Add("Alias", "U_BDO_CNTp");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
            formItems.Add("DisplayDesc", true);
            formItems.Add("ValidValues", listValidValues);

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
            string errorText;

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD && pVal.BeforeAction)
                {
                    createFormItems(oForm, out errorText);
                }
               
            }
        }
        public static void uiApp_FormDataEvent(  ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;

            SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(BusinessObjectInfo.FormTypeEx, Program.currentFormCount);

            if (oForm.TypeEx == "70011")
            {
                if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD & BusinessObjectInfo.BeforeAction == true)
                {
                    oForm.Freeze(true);
                    int panelLevel = oForm.PaneLevel;
                    string sdocDate = oForm.Items.Item("10").Specific.Value;
                    oForm.PaneLevel = 7;
                    oForm.Items.Item("1000").Specific.Value = sdocDate;
                    oForm.PaneLevel = panelLevel;
                    oForm.Freeze(false);

                }
            }
        }
    }
}
