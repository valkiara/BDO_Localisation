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

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if(oForm.TypeEx == "196")
                {
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DRAW && pVal.BeforeAction == false && Program.openPaymentMeansByPostDateChange)
                    {
                        try
                        {
                            oForm.Items.Item("44").Specific.Value = Program.newPostDateStr;
                            oForm.Items.Item("12").Specific.Value = FormsB1.ConvertDecimalToStringForEditboxStrings(Program.overallAmount);
                            oForm.Items.Item("34").Specific.Value = FormsB1.ConvertDecimalToStringForEditboxStrings(Program.transferSumFC);
                        }
                        catch (Exception ex)
                        {
                            errorText = ex.Message;
                        }
                        finally
                        {
                            Program.transferSumFC = 0;
                            Program.overallAmount = 0;
                            Program.newPostDateStr = null;
                        }
                    }

                    if (pVal.ItemChanged == true && pVal.ItemUID == "34" && pVal.BeforeAction == false && Program.openPaymentMeansByPostDateChange)
                    {
                        //Program.openPaymentMeansByCurrRateChange = false;
                        oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    }

                    if (pVal.ItemChanged == true && pVal.ItemUID == "8" && pVal.BeforeAction == false)
                    {
                        CommonFunctions.fillDocRate(oForm, "OVPM", "OVPM");
                    }
                }
                else
                {
                    if (pVal.ItemChanged == true && pVal.ItemUID == "8" && pVal.BeforeAction == false)
                    {
                        CommonFunctions.fillDocRate(oForm, "ORCT", "ORCT");
                    }
                }
            }

            if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_CLOSE && !pVal.BeforeAction)
            {
                Program.openPaymentMeans = true;
            }
        }
    }
}
