using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BDO_Localisation_AddOn
{
    class BDO_StockRevaluation
    {
        public static void createFormItems(SAPbouiCOM.Form oForm)
        {
            try{
                Dictionary<string, object> formItems = new Dictionary<string, object>();

                SAPbouiCOM.Item oItem = oForm.Items.Item("62");
                int height = oItem.Height;
                int top = oForm.Items.Item("1002").Top + (oForm.Items.Item("1002").Top - oForm.Items.Item("8").Top);
                int left_s = oItem.Left;
                int width_s = oItem.Width;


                int left_e = oForm.Items.Item("61").Left;
                int width_e = oForm.Items.Item("61").Width;

                string itemName = "BDOSLanCos"; //10 characters
                formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                formItems.Add("Left", left_s);
                formItems.Add("Width", width_s);
                formItems.Add("Top", top + width_s);
                formItems.Add("Height", height);
                formItems.Add("UID", itemName);
                formItems.Add("Caption", BDOSResources.getTranslate("LandedCost"));
                formItems.Add("Enabled", false);

                string errorText = "";
                FormsB1.createFormItem(oForm, formItems, out errorText);
            } 
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {

            }
        }

        public static void uiApp_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;
            //string bstrUDOObjectType = "162";
            //int docEntry = 0;
            //Program.uiApp.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_UserDefinedObject, bstrUDOObjectType, docEntry.ToString());

            //if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            //{
                //SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm("70001", pVal.FormTypeCount);

            //}
        }

        public static void createDocument()
        {

        }
        //70001
    }
}
