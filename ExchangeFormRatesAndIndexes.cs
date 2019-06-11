using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace BDO_Localisation_AddOn
{
    static partial class ExchangeFormRatesAndIndexes
    {
        public static void createFormItems(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            SAPbouiCOM.Item oItemOK = oForm.Items.Item("1");
            Dictionary<string, object> formItems = new Dictionary<string, object>();
            
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            formItems.Add("Left", oItemOK.Left);
            formItems.Add("Width", oItemOK.Width * 2 + 5);
            formItems.Add("Top", oForm.Items.Item("17").Top);
            formItems.Add("Height", oForm.Items.Item("17").Height);
            formItems.Add("Caption", BDOSResources.getTranslate("ImportRate"));
            formItems.Add("UID", "BDO_ImRate");
           
            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            GC.Collect();
        }

        public static void uiApp_ItemEvent(  string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
                {
                    if (oForm.Modal == false)
                    {
                    createFormItems(oForm, out errorText);
                }
                }

                if (pVal.ItemUID == "BDO_ImRate" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK)
                {
                    Program.oExchangeFormRatesAndIndexes = oForm;
                    BDO_ImportRateForm.createForm( Program.oExchangeFormRatesAndIndexes, out errorText);
                }

                //კურსების ღილაკის დამალვა/გამოჩენა
                if (pVal.ItemUID == "6" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
                {
                    oForm.Items.Item("BDO_ImRate").Visible = false;
                }
                else if (pVal.ItemUID == "5" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
                {
                    oForm.Items.Item("BDO_ImRate").Visible = true;
                }
            }
        }
    }
}
