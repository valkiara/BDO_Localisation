using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;
using System.Data;
using System.Runtime.InteropServices;

namespace BDO_Localisation_AddOn
{
    class BDOSReportSettlementReconciliationAct
    {
        public static void chooseFromList( SAPbouiCOM.Form oForm, SAPbouiCOM.IChooseFromListEvent oCFLEvento, string itemUID, bool beforeAction, out string errorText)
        {
            errorText = null;

            try
            {
                if (beforeAction == false)
                {
                    SAPbouiCOM.DataTable oDataTable = null;
                    oDataTable = oCFLEvento.SelectedObjects;

                    if (oDataTable != null)
                    {
                        oForm.Items.Item("1000027").Enabled = true;
                        oForm.Items.Item("1000033").Enabled = true;

                        string vatNumber = Convert.ToString(oDataTable.GetValue("LicTradNum", 0)).Trim();
                        oForm.Items.Item("1000027").Specific.Value = vatNumber == "" ? "-" : vatNumber;
                        oForm.Items.Item("1000033").Specific.Value = Convert.ToString(oDataTable.GetValue("CardCode", 0));
                        oForm.Items.Item("1000003").Click();

                        oForm.Items.Item("1000027").Enabled = false;
                        oForm.Items.Item("1000033").Enabled = false;
                    }
                }

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

        public static void clearItems(SAPbouiCOM.Form oForm)
        {
            oForm.Freeze(true);

            oForm.Items.Item("1000033").Enabled = false;
            oForm.Items.Item("1000027").Enabled = false;
            oForm.Freeze(false);
        }

        public static void uiApp_ItemEvent(  string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = "";

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (oForm.Title.Contains("Settlement Reconciliation Act") != true)
                {
                    return;
                }

                if (pVal.ItemUID == "1000015" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST) 
                {
                    SAPbouiCOM.IChooseFromListEvent oCFLEvento = null;
                    oCFLEvento = ((SAPbouiCOM.IChooseFromListEvent)(pVal));
                    chooseFromList( oForm, oCFLEvento, pVal.ItemUID, pVal.BeforeAction, out errorText);
                }

                if (pVal.ItemUID == "1000015" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_KEY_DOWN && pVal.BeforeAction == false)
                {
                    clearItems(oForm);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD & pVal.BeforeAction == true)
                {
                    clearItems(oForm);
                }
            }
        }

    }
}
