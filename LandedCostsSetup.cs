using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BDO_Localisation_AddOn
{
    class LandedCostsSetup
    {
        public static void uiApp_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD & BusinessObjectInfo.BeforeAction == true)
            {
                return;
            }

            SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(BusinessObjectInfo.FormTypeEx, Program.currentFormCount);

            if (oForm.TypeEx == "898")
            {
                if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD || BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE)
                {
                    if (BusinessObjectInfo.BeforeAction == true)
                    {
                        checkAccounts(oForm, out errorText);
                        if (errorText != null)
                        {
                            BubbleEvent = false;
                        }

                    }
                }
            }
        }

        public static void checkAccounts(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            string olist = "";
            string acctCode;
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("3").Specific;

            for (int i = 1; i <= oMatrix.RowCount; i++)
            {
                acctCode = oMatrix.Columns.Item("LaCAllcAcc").Cells.Item(i).Specific.Value.Trim();
                if (acctCode != "" && olist.Contains(", " + acctCode))
                {
                    errorText = BDOSResources.getTranslate("DuplicateAccountInTheLine") + " " + (i).ToString();
                    Program.uiApp.StatusBar.SetSystemMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    olist = olist + ", " + acctCode;
                }
            }
        }
    }
}
