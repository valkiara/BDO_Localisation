using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BDO_Localisation_AddOn
{
    static partial class ARDownPayment
    {
        public static void createFormItems(  SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            
            double height = oForm.Items.Item("86").Height;
            double top = oForm.Items.Item("86").Top + height * 1.5 + 1;
            double left_s = oForm.Items.Item("86").Left;
            double left_e = oForm.Items.Item("46").Left;
            double width_e = oForm.Items.Item("46").Width;

            ////-------------------------------------------ანგარიშ-ფაქტურა----------------------------------->

            //formItems = new Dictionary<string, object>();
            //itemName = "BDO_TaxTxt"; //10 characters
            //formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            //formItems.Add("Left", left_s);
            //formItems.Add("Width", width_e * 1.5);
            //formItems.Add("Top", top);
            //formItems.Add("Height", height);
            //formItems.Add("UID", itemName);
            //formItems.Add("Caption", BDOSResources.getTranslate("CreateTaxInvoice"));
            //formItems.Add("TextStyle", 4);
            //formItems.Add("FontSize", 10);
            //formItems.Add("Enabled", true);

            //FormsB1.createFormItem(oForm, formItems, out errorText);
            //if (errorText != null)
            //{
            //    return;
            //}

            //bool multiSelection = false;
            //string objectType = "UDO_F_BDO_TAXS_D"; //Tax invoice sent document
            //string uniqueID_TaxInvoiceSentCFL = "TaxInvoiceSent_CFL";
            //FormsB1.addChooseFromList( oForm, multiSelection, objectType, uniqueID_TaxInvoiceSentCFL);

            //formItems = new Dictionary<string, object>();
            //itemName = "BDO_TaxDoc"; //10 characters
            //formItems.Add("isDataSource", true);
            //formItems.Add("DataSource", "UserDataSources");
            //formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
            //formItems.Add("Length", 11);
            //formItems.Add("TableName", "");
            //formItems.Add("Alias", itemName);
            //formItems.Add("Bound", true);
            //formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            //formItems.Add("Left", left_e + width_e - 40);
            //formItems.Add("Width", 40);
            //formItems.Add("Top", top);
            //formItems.Add("Height", height);
            //formItems.Add("UID", itemName);
            //formItems.Add("AffectsFormMode", false);
            //formItems.Add("DisplayDesc", true);
            //formItems.Add("Enabled", false);
            //formItems.Add("ChooseFromListUID", uniqueID_TaxInvoiceSentCFL);
            //formItems.Add("ChooseFromListAlias", "DocEntry");

            //FormsB1.createFormItem(oForm, formItems, out errorText);
            //if (errorText != null)
            //{
            //    return;
            //}

            //formItems = new Dictionary<string, object>();
            //itemName = "BDO_TaxLB"; //10 characters
            //formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            //formItems.Add("Left", left_e + width_e - 40 - 20);
            //formItems.Add("Top", top);
            //formItems.Add("Height", height);
            //formItems.Add("UID", itemName);
            //formItems.Add("LinkTo", "BDO_TaxDoc");
            //formItems.Add("LinkedObjectType", objectType);

            //FormsB1.createFormItem(oForm, formItems, out errorText);
            //if (errorText != null)
            //{
            //    return;
            //}

            //top = top + height + 1;

            //oForm.DataSources.UserDataSources.Add("BDO_TaxSer", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
            //oForm.DataSources.UserDataSources.Add("BDO_TaxNum", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
            //oForm.DataSources.UserDataSources.Add("BDO_TaxDat", SAPbouiCOM.BoDataType.dt_DATE, 20);
            ////<-------------------------------------------ანგარიშ-ფაქტურა-----------------------------------

            GC.Collect();
        }
    }
}
