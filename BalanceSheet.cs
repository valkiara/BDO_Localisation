using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BDO_Localisation_AddOn
{
    class BalanceSheet
    {
        public static void createFormItems(  SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            Dictionary<string, object> formItems = null;

            string itemName = "";

            double top = oForm.Items.Item("1").Top;
            double width = oForm.Items.Item("1").Width;
            double left = oForm.Items.Item("1").Left + width;
            double height = oForm.Items.Item("1").Height;

            formItems = new Dictionary<string, object>();
            itemName = "ExportExc";
            formItems.Add("Caption", BDOSResources.getTranslate("Export"));
            formItems.Add("Size", 20);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            formItems.Add("Left", left + 5);
            formItems.Add("Width", width);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            GC.Collect();
        }

        public static void exportExcel(  SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            FinancialReports.ExportToExcel(oForm, "12", "1", "4", "5", out errorText); 
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

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK & pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "ExportExc")
                    {
                        oForm.Freeze(true);
                        exportExcel(  oForm, out errorText);
                        oForm.Update();
                        oForm.Freeze(false);
                    }
                }
            }

        }
    }
}
