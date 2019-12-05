using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BDO_Localisation_AddOn
{
    static partial class BPBankAccounts
    {
        public static void createUserFields(out string errorText)
        {
            Dictionary<string, object> fieldskeysMap;
            Dictionary<string, string> listValidValuesDict;

            listValidValuesDict = new Dictionary<string, string>();
            listValidValuesDict.Add("Y", "Yes");
            listValidValuesDict.Add("N", "No");

            fieldskeysMap = new Dictionary<string, object>(); //პროგრამა
            fieldskeysMap.Add("Name", "treasury");
            fieldskeysMap.Add("TableName", "OCRB");
            fieldskeysMap.Add("Description", "Treasury");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 1);
            fieldskeysMap.Add("ValidValues", listValidValuesDict);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            GC.Collect();
        }

        public static void changeFormItems(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("3").Specific;

            SAPbouiCOM.Columns oColumns = oMatrix.Columns;
            SAPbouiCOM.Column oColumn = oColumns.Item("U_treasury");      
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Treasury");
            oColumn.Width = 70;
        }

        public static void uiApp_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD && pVal.BeforeAction)
                {
                    changeFormItems(oForm);
                }
            }
        }
    }
}
