using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BDO_Localisation_AddOn
{
   static partial class PaymentMeans
    {
        public static void uiApp_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            //if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_CLOSE && pVal.BeforeAction == true && pVal.InnerEvent == false)
            //{
            //    int answer = Program.uiApp.MessageBox(BDOSResources.getTranslate("DoYouWantToClose") + "?", 1, BDOSResources.getTranslate("Yes"), BDOSResources.getTranslate("No"), "");

            //    if (answer != 1)
            //    {
            //        BubbleEvent = false;
            //    }
            //}

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DRAW && pVal.BeforeAction == false && Program.openPaymentMeansByCurrRateChange)
                {
                    try
                    {
                        oForm.Items.Item("34").Specific.Value = FormsB1.ConvertDecimalToStringForEditboxStrings(Program.transferSumFC);
                    }
                    catch (Exception ex)
                    {
                        errorText = ex.Message;
                    }
                    finally
                    {
                        Program.transferSumFC = 0;
                    }
                }

                if (pVal.ItemChanged == true && pVal.ItemUID == "34" && pVal.BeforeAction == false && Program.openPaymentMeansByCurrRateChange)
                {
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
            }

            if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_CLOSE && pVal.BeforeAction == false)
            {
                Program.openPaymentMeans = true;
            }
        }
    }
}
