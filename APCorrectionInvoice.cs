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
