using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Data;

namespace BDO_Localisation_AddOn
{
    static partial class ExchangeRateDiffs
    {

        public static void uiApp_ItemEvent(  string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            
            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD & pVal.BeforeAction == true)
                {
                    Program.Exchange_Rate_Save_Click = false;
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
                {


                    if (pVal.ItemUID == "1" && pVal.BeforeAction == true)
                    {
                        if (Program.Exchange_Rate_Save_Click == false)
                        {
                            Program.uiApp.SetStatusBarMessage(BDOSResources.getTranslate("YouMustRunFromPreviusReport"), SAPbouiCOM.BoMessageTime.bmt_Short, true);
                            BubbleEvent = false;
                        }

                    }

                    if (pVal.ItemUID == "17" && pVal.BeforeAction == false)
                    {
                        try
                        {
                            bool result = UpdateRTM1Table();
                            
                        }
                        catch (Exception ex)
                        {
                            Program.uiApp.SetStatusBarMessage(BDOSResources.getTranslate("CantUpdateRTM1") + ": " + ex, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                        }
                    }
                }


            }
        }

        public static void uiApp_ItemEvent1(  string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            
            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
                {
                    if (pVal.ItemUID == "49" && pVal.BeforeAction == false)
                    {
                        try
                        {
                            bool result = UpdateRTM1Table();
                            Program.Exchange_Rate_Save_Click = true;
                        }
                        catch
                        {
                            Program.Exchange_Rate_Save_Click = true;
                        }
                    }
                }


            }
        }

        private static bool UpdateRTM1Table()
        {
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string query = @"select
                            * 
                            from""RTM1"" 
                            inner join ""OACT"" on ""RTM1"".""JdtAcctCod"" = ""OACT"".""AcctCode"" 
                            and ""OACT"".""ExchRate"" = 'N'";

            oRecordSet.DoQuery(query);

            while (!oRecordSet.EoF)
            {
                string TransId = oRecordSet.Fields.Item("TransId").Value.ToString();

                SAPbobsCOM.Recordset oRecordSetUpdate = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string queryUpdate = @"update ""RTM1""
                                    set ""Valid"" = 'N'
                                    Where ""TransId"" = " + TransId + "";
                oRecordSetUpdate.DoQuery(queryUpdate);
                

                oRecordSet.MoveNext();
            }

            return true;
        }
    }
}
