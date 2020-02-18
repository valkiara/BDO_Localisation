using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BDO_Localisation_AddOn
{
    class BDO_StockRevaluation
    {
        public static void createFormItems(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            int left = 6;
            int Top = 5;
            string itemName = "";
            Dictionary<string, object> formItems = null;

            formItems = new Dictionary<string, object>();
            itemName = "LndCostS";
            formItems.Add("Size", 20);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            formItems.Add("Left", left);
            formItems.Add("Width", 120);
            formItems.Add("Top", Top + 10);
            formItems.Add("Caption", BDOSResources.getTranslate("LandedCost"));
            formItems.Add("UID", itemName);
            formItems.Add("LinkTo", "LndCostE");
            //13 ar unda vutxra razec minda ro gavides is unda vutxra eg 13 ari objectType
            formItems.Add("LinkedObjectType", "69");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            left = left + 128 + 10;

            formItems = new Dictionary<string, object>();
            itemName = "LndCostE";
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "UserDataSources");
            formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
            formItems.Add("Length", 30);
            formItems.Add("Size", 20);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("TableName", "");
            formItems.Add("Alias", itemName);
            formItems.Add("Bound", true);
            formItems.Add("Left", left);
            formItems.Add("Width", 100);
            formItems.Add("Top", Top + 10);
            formItems.Add("Height", 19);
            formItems.Add("UID", itemName);
            formItems.Add("Enabled", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

        }

        public static void fillLandedCostNumber(SAPbouiCOM.Form oForm, string docNum)
        {
            oForm.DataSources.UserDataSources.Item("LndCostE").ValueEx = docNum;
        }

        public static void fillStockRevaluation(string docNum)
        {

        }
        public static void uiApp_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;
            SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);
            int a = 7;
        }

            /*
            public static void registerUDO(out string errorText)
            {
                errorText = null;
                string code = "UDO_F_BDO_PTBT_D"; //20 characters (must include at least one alphabetical character).
                Dictionary<string, object> formProperties;


                UDO.registerUDO(code, formProperties, out errorText);

            }
            */








            /*
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

        */
        }
    }
