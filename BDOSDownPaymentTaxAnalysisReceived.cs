using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BDO_Localisation_AddOn
{
    class BDOSDownPaymentTaxAnalysisReceived
    {
        public static void uiApp_ItemEvent(  string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            
            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (oForm.Title.Contains("Down Payment Tax Invoice Received Analysis") != true)
                {
                    return;
                }

                if (pVal.ItemUID == "1000003" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED & pVal.BeforeAction == false)
                {
                    SAPbouiCOM.EditText oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("1000003").Specific;
                    DateTime Date = DateTime.ParseExact(oEditText.Value, "yyyyMMdd", null);
                    Date = new DateTime(Date.Year, Date.Month, 1);
                    oEditText.Value = Date.ToString("yyyyMMdd");
                }
            }
        }


    }
}
