using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace BDO_Localisation_AddOn
{
    static partial class HouseBankAccounts
    {
        public static void createUserFields( out string errorText)
        {
            Dictionary<string, object> fieldskeysMap;
            Dictionary<string, string> listValidValuesDict;

            listValidValuesDict = new Dictionary<string, string>();
            listValidValuesDict.Add("empty", "");
            listValidValuesDict.Add("TBC", "TBC (Web-Service)");         
            listValidValuesDict.Add("BOG", "BOG (Web-Service)");

            fieldskeysMap = new Dictionary<string, object>(); //პროგრამა
            fieldskeysMap.Add("Name", "program");
            fieldskeysMap.Add("TableName", "DSC1");
            fieldskeysMap.Add("Description", "Program");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);
            fieldskeysMap.Add("ValidValues", listValidValuesDict);
            errorText = null;

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            bool result = UDO.addNewValidValuesUserFieldsMD( "DSC1", "program", "BOG", "BOG (Web-Service)", out errorText);

            GC.Collect();
        }
    
        public static void changeFormItems(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("3").Specific));
            
            SAPbouiCOM.Columns oColumns = oMatrix.Columns;         
            SAPbouiCOM.Column oColumn;

            oColumn = oColumns.Item("U_program");
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Program"); 
            oColumn.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
        }

        public static void checkGLAccounts(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            List<string> oList = new List<string>();
            SAPbouiCOM.DBDataSource oDBDataSource = oForm.DataSources.DBDataSources.Item("DSC1");

            for (int i = 0; i < oDBDataSource.Size; i++)
            {
                string GLAccount = oDBDataSource.GetValue("GLAccount", i).Trim();
                if (string.IsNullOrEmpty(GLAccount) == false)
                {
                    oList.Add(GLAccount);
                }
            }

            List<string> duplicates = new List<string>();
            if (oList.Count > 0)
            {
                duplicates = oList.GroupBy(s => s).SelectMany(grp => grp.Skip(1)).ToList();
                if (duplicates.Count > 0)
                {
                    errorText = BDOSResources.getTranslate("GLAccountMustBeUniqueDuplicateGLAccounts") + " : " + string.Join(",", duplicates);
                }
            }
        }

        public static void uiApp_FormDataEvent(  ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(BusinessObjectInfo.FormTypeEx, Program.currentFormCount);

            if (oForm.TypeEx == "60701")
            {
                //შემოწმება
                if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD)
                {
                    if (BusinessObjectInfo.BeforeAction == true)
                    {
                        checkGLAccounts(oForm, out errorText);
                    
                        if (errorText != null)
                        {
                            Program.uiApp.MessageBox(errorText);
                            BubbleEvent = false;
                        }
                    }                  
                }
                
                if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE)
                {
                    if (BusinessObjectInfo.BeforeAction == true)
                    {
                        checkGLAccounts(oForm, out errorText);

                        if (errorText != null)
                        {
                            Program.uiApp.MessageBox(errorText);
                            BubbleEvent = false;
                        }
                    }
                }
            }
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
                    changeFormItems(oForm, out errorText);
                }
            }
        }
    }
}
