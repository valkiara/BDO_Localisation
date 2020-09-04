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
    class BDOSOutgoingPaymentsWizard
    {
        public static string blnktAgrOld;

        public static void createForm(out string errorText)
        {
            int formHeight = Program.uiApp.Desktop.Height;
            int formWidth = Program.uiApp.Desktop.Width;
            Dictionary<string, object> formItems;
            string itemName;
            SAPbouiCOM.Columns oColumns;
            SAPbouiCOM.Column oColumn;

            SAPbouiCOM.DataTable oDataTable;

            int left_s = 5;
            int left_e = 180;
            int left_s2 = formWidth - 550;
            int left_e2 = left_s2 + 220;
            int width_s = 155;
            int width_e = 200;
            int top = 10;
            int height = 15;

            //ფორმის აუცილებელი თვისებები
            Dictionary<string, object> formProperties = new Dictionary<string, object>();
            formProperties.Add("UniqueID", "BDOSSOPWizzForm");
            formProperties.Add("BorderStyle", SAPbouiCOM.BoFormBorderStyle.fbs_Sizable);
            formProperties.Add("Title", BDOSResources.getTranslate("OutgoingPaymentWizard"));
            formProperties.Add("ClientWidth", formWidth);
            formProperties.Add("ClientHeight", formHeight);

            SAPbouiCOM.Form oForm;
            bool newForm;
            bool formExist = FormsB1.createForm(formProperties, out oForm, out newForm, out errorText);

            if (formExist == true)
            {
                if (newForm)
                {
                    FormsB1.addChooseFromList(oForm, false, "2", "BusinessPartner_CFL");

                    //Conditions for Business Partner -->
                    SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item("BusinessPartner_CFL");
                    SAPbouiCOM.Conditions oCons = oCFL.GetConditions();

                    SAPbouiCOM.Condition oCon = oCons.Add();
                    oCon.Alias = "CardType";
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCon.CondVal = "S"; //მომწოდებელი

                    oCFL.SetConditions(oCons);
                    //Conditions for Business Partner <--

                    formItems = new Dictionary<string, object>();
                    itemName = "BPCodeS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("CardCode"));

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "BPCode"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("TableName", "");
                    formItems.Add("Length", 20);
                    formItems.Add("Size", 20);
                    formItems.Add("Alias", "BPCode");
                    formItems.Add("Bound", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("Left", left_e);
                    formItems.Add("Width", width_e);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("ChooseFromListUID", "BusinessPartner_CFL");
                    formItems.Add("ChooseFromListAlias", "CardCode");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "BPCodeLB"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    formItems.Add("Left", left_e - 20);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("LinkTo", "BPCode");
                    formItems.Add("LinkedObjectType", "2");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "DocPsDtS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s2);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("DocumentPostingDate"));

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "DocPstDt";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_DATE);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("Length", 1);
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Left", left_e2);
                    formItems.Add("Width", width_e);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("ValueEx", DateTime.Now.ToString("yyyyMMdd"));

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    top = top + height + 1;

                    FormsB1.addChooseFromList(oForm, false, "187", "BPBankAcc_CFL");

                    formItems = new Dictionary<string, object>();
                    itemName = "BPBankAccS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("BPBankAccount"));

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "BPBankAcc"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("TableName", "");
                    formItems.Add("Length", 50);
                    formItems.Add("Size", 50);
                    formItems.Add("Alias", "BPBankAcc");
                    formItems.Add("Bound", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("Left", left_e);
                    formItems.Add("Width", width_e);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("ChooseFromListUID", "BPBankAcc_CFL");
                    formItems.Add("ChooseFromListAlias", "Account");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    FormsB1.addChooseFromList(oForm, false, "1", "CTAcc_CFL");

                    //Conditions for Control Account -->
                    oCFL = oForm.ChooseFromLists.Item("CTAcc_CFL");
                    oCons = oCFL.GetConditions();

                    oCon = oCons.Add();
                    oCon.Alias = "LocManTran"; //Lock Manual Transaction (Control Account)
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCon.CondVal = "Y";

                    oCFL.SetConditions(oCons);
                    //Conditions for Control Account -->

                    formItems = new Dictionary<string, object>();
                    itemName = "CTAccS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s2);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("ControlAccount"));

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "CTAcc"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("TableName", "");
                    formItems.Add("Length", 50);
                    formItems.Add("Size", 50);
                    formItems.Add("Alias", "CTAcc");
                    formItems.Add("Bound", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("Left", left_e2);
                    formItems.Add("Width", width_e);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("ChooseFromListUID", "CTAcc_CFL");
                    formItems.Add("ChooseFromListAlias", "AcctCode");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "CTAccLB"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    formItems.Add("Left", left_e2 - 20);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("LinkTo", "CTAcc");
                    formItems.Add("LinkedObjectType", "1");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    top = top + height + 1;

                    FormsB1.addChooseFromList(oForm, false, "1", "GLAcc_CFL");

                    //Conditions for G/L Account -->
                    oCFL = oForm.ChooseFromLists.Item("GLAcc_CFL");
                    oCons = oCFL.GetConditions();

                    oCon = oCons.Add();
                    oCon.Alias = "Postable"; //Active Account, (Title Account)
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCon.CondVal = "Y";

                    oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                    oCon = oCons.Add();
                    oCon.Alias = "LocManTran"; //Lock Manual Transaction (Control Account)
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCon.CondVal = "N";

                    oCFL.SetConditions(oCons);
                    //Conditions for G/L Account <--

                    formItems = new Dictionary<string, object>();
                    itemName = "GLAccS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("GLAccount"));

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "GLAcc"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("TableName", "");
                    formItems.Add("Length", 50);
                    formItems.Add("Size", 50);
                    formItems.Add("Alias", "GLAcc");
                    formItems.Add("Bound", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("Left", left_e);
                    formItems.Add("Width", width_e);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("ChooseFromListUID", "GLAcc_CFL");
                    formItems.Add("ChooseFromListAlias", "AcctCode");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "GLAccLB"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    formItems.Add("Left", left_e - 20);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("LinkTo", "GLAcc");
                    formItems.Add("LinkedObjectType", "1");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    //formItems = new Dictionary<string, object>();
                    //itemName = "WHtaxS"; //10 characters
                    //formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    //formItems.Add("Left", left_s2);
                    //formItems.Add("Width", width_s);
                    //formItems.Add("Top", top);
                    //formItems.Add("Height", height);
                    //formItems.Add("UID", itemName);
                    //formItems.Add("Caption", BDOSResources.getTranslate("WithholdingTax"));

                    //FormsB1.createFormItem(oForm, formItems, out errorText);
                    //if (errorText != null)
                    //{
                    //    return;
                    //}

                    //formItems = new Dictionary<string, object>();
                    //itemName = "WHTax"; //10 characters
                    //formItems.Add("isDataSource", true);
                    //formItems.Add("DataSource", "UserDataSources");
                    //formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    //formItems.Add("TableName", "");
                    //formItems.Add("Length", 4);
                    //formItems.Add("Size", 20);
                    //formItems.Add("Alias", itemName);
                    //formItems.Add("Bound", true);
                    //formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                    //formItems.Add("Left", left_e2);
                    //formItems.Add("Width", width_e);
                    //formItems.Add("Top", top);
                    //formItems.Add("Height", height);
                    //formItems.Add("UID", itemName);
                    //formItems.Add("DisplayDesc", true);
                    //formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);

                    //FormsB1.createFormItem(oForm, formItems, out errorText);
                    //if (errorText != null)
                    //{
                    //    return;
                    //}

                    formItems = new Dictionary<string, object>();
                    itemName = "CashFlowIS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s2);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("PrimaryFormItem"));

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    Dictionary<string, string> CFWList = CommonFunctions.getCashFlowLineItemsList();

                    formItems = new Dictionary<string, object>();
                    itemName = "CashFlowI"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("Length", 11);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
                    formItems.Add("Left", left_e2);
                    formItems.Add("Width", width_e);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("ValidValues", CFWList);
                    formItems.Add("ValueEx", CommonFunctions.getOADM("CfwOutDflt").ToString());

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    top = top + height + 1;

                    formItems = new Dictionary<string, object>();
                    itemName = "DispTypeS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("DispatchType"));

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    Dictionary<string, string> listValidValuesDict = new Dictionary<string, string>();
                    listValidValuesDict.Add("", "");
                    listValidValuesDict.Add("BULK", "BULK"); //BULK - სტანდარტული გადარიცხვა
                    listValidValuesDict.Add("MT103", "MT103"); //MT103 ინდივიდუალური გადარიცხვა

                    formItems = new Dictionary<string, object>();
                    itemName = "DispType"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("TableName", "");
                    formItems.Add("Length", 20);
                    formItems.Add("Size", 20);
                    formItems.Add("Alias", "DispType");
                    formItems.Add("Bound", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                    formItems.Add("Left", left_e);
                    formItems.Add("Width", width_e);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("ValidValues", listValidValuesDict);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "DescrptS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s2);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("Description"));

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "Descrpt"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("Length", 254);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("Left", left_e2);
                    formItems.Add("Width", width_e);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    top = top + height + 1;

                    formItems = new Dictionary<string, object>();
                    itemName = "ChrgDtlsS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("ChrgDtls"));

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    listValidValuesDict = new Dictionary<string, string>();
                    listValidValuesDict.Add("", "");
                    listValidValuesDict.Add("SHA", "SHA");
                    listValidValuesDict.Add("OUR", "OUR");

                    formItems = new Dictionary<string, object>();
                    itemName = "ChrgDtls"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("TableName", "");
                    formItems.Add("Length", 20);
                    formItems.Add("Size", 20);
                    formItems.Add("Alias", "ChrgDtls");
                    formItems.Add("Bound", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                    formItems.Add("Left", left_e);
                    formItems.Add("Width", width_e);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("ValidValues", listValidValuesDict);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    //Budget Cash Flow - Chartulia Alami da Gashvebulia Construction
                    if (CommonFunctions.IsDevelopment())
                    {
                        top = top + height + 1;

                        formItems = new Dictionary<string, object>();
                        itemName = "BDOSDefCfS"; //10 characters
                        formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                        formItems.Add("Left", left_s2);
                        formItems.Add("Width", width_s);
                        formItems.Add("Top", top);
                        formItems.Add("Height", height);
                        formItems.Add("UID", itemName);
                        formItems.Add("Caption", BDOSResources.getTranslate("BudgetCashFlow"));
                        formItems.Add("LinkTo", "BDOSBdgCfE");

                        FormsB1.createFormItem(oForm, formItems, out errorText);
                        if (errorText != null)
                        {
                            return;
                        }

                        FormsB1.addChooseFromList(oForm, false, "UDO_F_BDOSBUCFW_D", "Budg_CFLHD");

                        formItems = new Dictionary<string, object>();
                        itemName = "BDOSDefCfE"; //10 characters
                        formItems.Add("isDataSource", true);
                        formItems.Add("DataSource", "UserDataSources");
                        formItems.Add("TableName", "");
                        formItems.Add("Length", 200);
                        formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                        formItems.Add("Alias", "BDOSDefCfE");
                        formItems.Add("Bound", true);
                        formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                        formItems.Add("Left", left_e2);
                        formItems.Add("Width", 30);
                        formItems.Add("Top", top);
                        formItems.Add("Height", height);
                        formItems.Add("UID", itemName);
                        formItems.Add("ChooseFromListUID", "Budg_CFLHD");
                        formItems.Add("ChooseFromListAlias", "Code");

                        FormsB1.createFormItem(oForm, formItems, out errorText);
                        if (errorText != null)
                        {
                            return;
                        }

                        formItems = new Dictionary<string, object>();
                        itemName = "BDOSDefCfN"; //10 characters
                        formItems.Add("isDataSource", true);
                        formItems.Add("DataSource", "UserDataSources");
                        formItems.Add("TableName", "");
                        formItems.Add("Length", 200);
                        formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                        formItems.Add("Alias", "BDOSDefCfN");
                        formItems.Add("Bound", true);
                        formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                        formItems.Add("Left", left_e2 + 30 + 5);
                        formItems.Add("Width", 80);
                        formItems.Add("Top", top);
                        formItems.Add("Height", height);
                        formItems.Add("UID", itemName);

                        FormsB1.createFormItem(oForm, formItems, out errorText);
                        if (errorText != null)
                        {
                            return;
                        }

                        formItems = new Dictionary<string, object>();
                        itemName = "fillBdgFl";
                        formItems.Add("Caption", BDOSResources.getTranslate("Fill"));
                        formItems.Add("Size", 20);
                        formItems.Add("DisplayDesc", true);
                        formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                        formItems.Add("Left", left_e2 + 40 + 5 + 70);
                        formItems.Add("Width", 45);
                        formItems.Add("Top", top);
                        formItems.Add("Height", height);
                        formItems.Add("UID", itemName);

                        FormsB1.createFormItem(oForm, formItems, out errorText);
                        if (errorText != null)
                        {
                            return;
                        }
                    }

                    top = top + 2 * height + 1;

                    formItems = new Dictionary<string, object>();
                    itemName = "InCheck";
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", 19);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Image", "HANA_CHECKBOX_CH");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "InUncheck";
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left_s + 20 + 1);
                    formItems.Add("Width", 19);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Image", "HANA_CHECKBOX_UH");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    //formItems = new Dictionary<string, object>();
                    //itemName = "AddRow";
                    //formItems.Add("Caption", BDOSResources.getTranslate("AddNewRow"));
                    //formItems.Add("Size", 20);
                    //formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    //formItems.Add("Left", left_s + (20 + 1) * 2);
                    //formItems.Add("Width", 100);
                    //formItems.Add("Top", top);
                    //formItems.Add("Height", height);
                    //formItems.Add("UID", itemName);

                    //FormsB1.createFormItem(oForm, formItems, out errorText);
                    //if (errorText != null)
                    //{
                    //    return;
                    //}

                    //formItems = new Dictionary<string, object>();
                    //itemName = "delMTRB";
                    //formItems.Add("Caption", BDOSResources.getTranslate("DeleteRow"));
                    //formItems.Add("Size", 20);
                    //formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    //formItems.Add("Left", left_s + 105 + (20 + 1) * 2);
                    //formItems.Add("Width", 100);
                    //formItems.Add("Top", top);
                    //formItems.Add("Height", height);
                    //formItems.Add("UID", itemName);

                    //FormsB1.createFormItem(oForm, formItems, out errorText);
                    //if (errorText != null)
                    //{
                    //    return;
                    //}

                    formItems = new Dictionary<string, object>();
                    itemName = "CreatDocmt";
                    formItems.Add("Caption", BDOSResources.getTranslate("CreateDocuments"));
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left_s + (20 + 1) * 2); //left_s + 105 * 2 + (20 + 1) * 2);
                    formItems.Add("Width", 150);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    top = top + height + 5;
                    int heightMTR = 550;
                    formItems = new Dictionary<string, object>();
                    itemName = "InvoiceMTR"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", 600);
                    formItems.Add("Top", top);
                    formItems.Add("Height", heightMTR);
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("InvoiceMTR").Specific;
                    oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
                    oColumns = oMatrix.Columns;

                    oDataTable = oForm.DataSources.DataTables.Add("InvoiceMTR");

                    oDataTable.Columns.Add("LineNum", SAPbouiCOM.BoFieldsType.ft_Integer, 50); //ინდექსი 
                    oDataTable.Columns.Add("CheckBox", SAPbouiCOM.BoFieldsType.ft_Text, 1); // 
                    oDataTable.Columns.Add("DocEntry", SAPbouiCOM.BoFieldsType.ft_Integer, 50); //ენთრი
                    oDataTable.Columns.Add("DocNum", SAPbouiCOM.BoFieldsType.ft_Integer, 50); //ნომერი
                    oDataTable.Columns.Add("Project", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 50);
                    oDataTable.Columns.Add("InstallmentID", SAPbouiCOM.BoFieldsType.ft_Integer, 6); //გადარიცხვის ID
                    oDataTable.Columns.Add("LineID", SAPbouiCOM.BoFieldsType.ft_Integer, 11);
                    oDataTable.Columns.Add("DocType", SAPbouiCOM.BoFieldsType.ft_Text, 50); //დოკუმენტის ტიპი
                    oDataTable.Columns.Add("DocDate", SAPbouiCOM.BoFieldsType.ft_Date, 50); //თარიღი
                    oDataTable.Columns.Add("DueDate", SAPbouiCOM.BoFieldsType.ft_Date, 50); //თარიღი
                    oDataTable.Columns.Add("WTCode", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 4); //Withholding Tax Data (WTCode)
                    oDataTable.Columns.Add("Arrears", SAPbouiCOM.BoFieldsType.ft_Text, 1); //* აჩვენებს, რომ Due Date ნაკლებია ან ტოლი გადახდის თარიღზე
                    oDataTable.Columns.Add("OverdueDays", SAPbouiCOM.BoFieldsType.ft_Integer, 50); //გადახდის თარიღსა და Due Date-ს შორის სხვაობა
                    oDataTable.Columns.Add("Total", SAPbouiCOM.BoFieldsType.ft_Sum); //თანხა
                    oDataTable.Columns.Add("WTSum", SAPbouiCOM.BoFieldsType.ft_Sum); //თანხა                    
                    oDataTable.Columns.Add("PensEmplr", SAPbouiCOM.BoFieldsType.ft_Sum); //თანხა  //დამსაქმებელი
                    oDataTable.Columns.Add("PensEmpld", SAPbouiCOM.BoFieldsType.ft_Sum); //თანხა  //დასაქმებული
                    oDataTable.Columns.Add("BalanceDue", SAPbouiCOM.BoFieldsType.ft_Sum); //დოკუმენტის დაურეკონსილირებელი თანხა - ვალის ნაშთი
                    oDataTable.Columns.Add("TotalPaymentLC", SAPbouiCOM.BoFieldsType.ft_Sum); //Default - Balance Due
                    oDataTable.Columns.Add("TotalPaymentFC", SAPbouiCOM.BoFieldsType.ft_Sum); //Default - Balance Due
                    //oDataTable.Columns.Add("TotalPaymentNet", SAPbouiCOM.BoFieldsType.ft_Sum); //Default - Balance Due
                    oDataTable.Columns.Add("Currency", SAPbouiCOM.BoFieldsType.ft_Text, 50); //დოკუმენტის ვალუტა
                    oDataTable.Columns.Add("UseBlaAgRt", SAPbouiCOM.BoFieldsType.ft_Text, 1);
                    oDataTable.Columns.Add("BlnktAgr", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20); //Blanket Agreement
                    oDataTable.Columns.Add("CFWId", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 11);
                    oDataTable.Columns.Add("Description", SAPbouiCOM.BoFieldsType.ft_Text, 100);
                    oDataTable.Columns.Add("Comments", SAPbouiCOM.BoFieldsType.ft_Text, 254); //კომენტარი

                    if (CommonFunctions.IsDevelopment())
                    {
                        oDataTable.Columns.Add("BudgetCashFlowID", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 11);
                        oDataTable.Columns.Add("BudgetCashFlowName", SAPbouiCOM.BoFieldsType.ft_Text, 100);
                    }

                    string uniqueID_lf_Budg_CFL = "Budg_CFL";

                    if (CommonFunctions.IsDevelopment())
                    {
                        FormsB1.addChooseFromList(oForm, false, "UDO_F_BDOSBUCFW_D", uniqueID_lf_Budg_CFL);
                    }

                    FormsB1.addChooseFromList(oForm, false, "63", "Proj_CFL");
                    FormsB1.addChooseFromList(oForm, false, "1250000025", "BlnktAgr_CFL"); //Blanket Agreement
                    SAPbouiCOM.LinkedButton oLink;

                    for (int count = 0; count < oDataTable.Columns.Count; count++)
                    {
                        var column = oDataTable.Columns.Item(count);
                        string columnName = column.Name;

                        if (columnName == "LineNum")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = "#";
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind("InvoiceMTR", columnName);
                        }
                        else if (columnName == "CheckBox")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
                            oColumn.TitleObject.Caption = "";
                            oColumn.ValOff = "N";
                            oColumn.ValOn = "Y";
                            oColumn.DataBind.Bind("InvoiceMTR", columnName);
                        }
                        else if (columnName == "DocEntry")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate(columnName);
                            oColumn.Editable = false;
                            oLink = oColumn.ExtendedObject;
                            oLink.LinkedObjectType = "13"; // - A/R Invoice, "14" - A/R Credit Note, A/R Down Payment Request - "203", Journal Entry - "30"
                            oColumn.DataBind.Bind("InvoiceMTR", columnName);
                        }
                        else if (columnName == "InstallmentID")
                        {
                            oColumn = oColumns.Add("InstlmntID", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate(columnName);
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind("InvoiceMTR", columnName);
                        }
                        else if (columnName == "LineID")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate(columnName);
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind("InvoiceMTR", columnName);
                            oColumn.Visible = false;
                        }
                        else if (columnName == "DocType")
                        {
                            oColumn = oColumns.Add("DocType", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate(columnName);
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind("InvoiceMTR", columnName);
                            oColumn.DisplayDesc = true;
                            oColumn.TitleObject.Sortable = true;
                            oColumn.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;

                            oColumn.ValidValues.Add("204", "DT");
                            oColumn.ValidValues.Add("18", "PU"); //BDOSResources.getTranslate("ARInvoice")
                            oColumn.ValidValues.Add("163", "CU"); //BDOSResources.getTranslate("ARCreditNote")
                        }
                        else if (columnName == "WTCode")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate(columnName);
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind("InvoiceMTR", columnName);
                        }
                        else if (columnName == "Arrears")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = "*";
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind("InvoiceMTR", columnName);
                        }
                        else if (columnName == "TotalPaymentLC")
                        {
                            oColumn = oColumns.Add("TtlPmntLC", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate(columnName);
                            oColumn.DataBind.Bind("InvoiceMTR", columnName);
                        }
                        else if (columnName == "TotalPaymentFC")
                        {
                            oColumn = oColumns.Add("TtlPmntFC", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate(columnName);
                            oColumn.DataBind.Bind("InvoiceMTR", columnName);
                        }
                        //else if (columnName == "TotalPaymentNet")
                        //{
                        //    oColumn = oColumns.Add("TtlPmntNt", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                        //    oColumn.TitleObject.Caption = BDOSResources.getTranslate(columnName);
                        //    oColumn.DataBind.Bind("InvoiceMTR", columnName);
                        //}
                        else if (columnName == "WTSum")
                        {
                            oColumn = oColumns.Add("WTSum", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("WTaxAmount");
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind("InvoiceMTR", columnName);
                        }
                        else if (columnName == "PensEmpld")
                        {
                            oColumn = oColumns.Add("PensEmpld", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("PhysicalEntityPension"); //დასაქმებული
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind("InvoiceMTR", columnName);
                        }
                        else if (columnName == "PensEmplr")
                        {
                            oColumn = oColumns.Add("PensEmplr", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("CompanyPension"); //დამსაქმებელი
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind("InvoiceMTR", columnName);
                        }
                        else if (columnName == "BudgetCashFlowID")
                        {
                            oColumn = oColumns.Add("BCFWId", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("BudgetCashFlowCodeOutgoingWizard");
                            oColumn.Editable = true;
                            oColumn.DataBind.Bind("InvoiceMTR", columnName);
                            oColumn.ChooseFromListUID = uniqueID_lf_Budg_CFL;
                            oColumn.ChooseFromListAlias = "Code";
                        }
                        else if (columnName == "BudgetCashFlowName")
                        {
                            oColumn = oColumns.Add("BCFWName", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Name");
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind("InvoiceMTR", columnName);
                        }
                        else if (columnName == "OverdueDays")
                        {
                            oColumn = oColumns.Add("OverdueDay", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate(columnName);
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind("InvoiceMTR", columnName);
                        }
                        else if (columnName == "Comments")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("DocumentRemarks");
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind("InvoiceMTR", columnName);
                        }
                        else if (columnName == "Project")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Project");
                            oColumn.Editable = true;
                            oColumn.DataBind.Bind("InvoiceMTR", columnName);
                            oColumn.ChooseFromListUID = "Proj_CFL";
                            oColumn.ChooseFromListAlias = "PrjCode";
                        }
                        else if (columnName == "UseBlaAgRt")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("UseBlAgrRt");
                            oColumn.DataBind.Bind("InvoiceMTR", columnName);
                            oColumn.Editable = false;
                            oColumn.ValOff = "N";
                            oColumn.ValOn = "Y";
                        }
                        else if (columnName == "BlnktAgr")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("BlanketAgreement");
                            oColumn.Editable = true;
                            oColumn.DataBind.Bind("InvoiceMTR", columnName);
                            oColumn.ChooseFromListUID = "BlnktAgr_CFL";
                            oColumn.ChooseFromListAlias = "AbsID";
                            oLink = oColumn.ExtendedObject;
                            oLink.LinkedObjectType = "1250000025";
                        }
                        else if (columnName == "CFWId")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("CashFlowLineItemID");
                            oColumn.DataBind.Bind("InvoiceMTR", columnName);
                            oColumn.DisplayDesc = true;
                            oColumn.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;

                            foreach (KeyValuePair<string, string> keyValue in CFWList)
                            {
                                oColumn.ValidValues.Add(keyValue.Key, keyValue.Value);
                            }
                        }
                        else if (columnName == "Description")
                        {
                            oColumn = oColumns.Add("Descrpt", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate(columnName);
                            oColumn.DataBind.Bind("InvoiceMTR", columnName);
                        }
                        else
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate(columnName);
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind("InvoiceMTR", columnName);
                        }
                    }
                    oMatrix.Clear();
                    oMatrix.LoadFromDataSource();
                    oMatrix.AutoResizeColumns();

                    top = top + heightMTR + 20;

                    formItems = new Dictionary<string, object>();
                    itemName = "TtlPmntLCS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s2);
                    formItems.Add("Width", width_s + 40);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("TotalPaymentLC"));

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "TtlPmntLCE"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SUM);
                    formItems.Add("Length", 11);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("Left", left_e2);
                    formItems.Add("Width", width_e - 40);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Enabled", false);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    top = top + height + 1;

                    formItems = new Dictionary<string, object>();
                    itemName = "TtlPmntFCS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s2);
                    formItems.Add("Width", width_s + 40);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("TotalPaymentFC"));

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "TtlPmntFCE"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SUM);
                    formItems.Add("Length", 11);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("Left", left_e2);
                    formItems.Add("Width", width_e - 40);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Enabled", false);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "TtlPmntFCT"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("Length", 11);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("Left", left_e2);
                    formItems.Add("Width", width_e - 40);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Enabled", false);
                    formItems.Add("Visible", false);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }
                }
                resizeItems(oForm);
                oForm.Visible = true;
                oForm.Select();
            }
        }

        private static void resizeItems(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.Item oMatrixItem = oForm.Items.Item("InvoiceMTR");

            oMatrixItem.Height = oForm.Height - 270;
            oMatrixItem.Width = oForm.Width - 20;
            int height = 15;

            int top = oMatrixItem.Top + oMatrixItem.Height + 20;
            oForm.Items.Item("TtlPmntLCS").Top = top;
            oForm.Items.Item("TtlPmntLCE").Top = top;
            top = top + height + 1;
            oForm.Items.Item("TtlPmntFCS").Top = top;
            oForm.Items.Item("TtlPmntFCE").Top = top;
            oForm.Items.Item("TtlPmntFCT").Top = top;
        }

        private static int createPaymentDocument(SAPbouiCOM.Form oForm, DataRow headerLine, DataTable AccountPaymentsLines)
        {
            string errorText;

            string docDateS = oForm.DataSources.UserDataSources.Item("DocPstDt").ValueEx;
            DateTime docDate = Convert.ToDateTime(DateTime.ParseExact(docDateS, "yyyyMMdd", CultureInfo.InvariantCulture));

            SAPbobsCOM.SBObob oSBOBob = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);

            DataTable DTSourceVPM2 = new DataTable();
            DTSourceVPM2.Columns.Add("InvType");
            DTSourceVPM2.Columns.Add("DocEntry");
            DTSourceVPM2.Columns.Add("AppliedFC");
            DTSourceVPM2.Columns.Add("SumApplied");

            DataTable DTSource = new DataTable();
            DTSource.Columns.Add("WtCode");
            DTSource.Columns.Add("WTLiable");
            DTSource.Columns.Add("CardCode");
            DTSource.Columns.Add("PrjCode");
            DTSource.Columns.Add("U_liablePrTx");
            DTSource.Columns.Add("U_prBase");
            DTSource.Columns.Add("U_BDOSWhtAmt");
            DTSource.Columns.Add("NoDocSum");
            DTSource.Columns.Add("U_BDOSPnPhAm");
            DTSource.Columns.Add("U_BDOSPnCoAm");

            string localCurrency = Program.LocalCurrency;
            string bpBankAccount = headerLine["BankAccount"].ToString();
            string transferAccount = headerLine["TransferAccount"].ToString();
            string invCurrency = headerLine["Currency"].ToString();
            string cardCode = headerLine["CardCode"].ToString();
            string project = headerLine["Project"].ToString();
            string wtCode = headerLine["WTCode"].ToString();
            string useBlaAgRt = headerLine["UseBlaAgRt"].ToString();
            string blnktAgr = headerLine["BlnktAgr"].ToString();

            decimal wtAmount = Convert.ToDecimal(headerLine["WtAmount"], NumberFormatInfo.InvariantInfo);
            decimal pensEmployedAmount = Convert.ToDecimal(headerLine["pensEmployedAmount"], NumberFormatInfo.InvariantInfo);
            decimal pensEmployerAmount = Convert.ToDecimal(headerLine["pensEmployerAmount"], NumberFormatInfo.InvariantInfo);
            decimal transferSumFC = Convert.ToDecimal(headerLine["PayblAmtFC"], NumberFormatInfo.InvariantInfo);
            decimal transferSumLC = Convert.ToDecimal(headerLine["PayblAmtLC"], NumberFormatInfo.InvariantInfo);
            decimal transferSum;

            SAPbobsCOM.Payments oPayment = (SAPbobsCOM.Payments)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments);

            oPayment.DocObjectCode = SAPbobsCOM.BoPaymentsObjectType.bopot_OutgoingPayments;
            oPayment.DocTypte = SAPbobsCOM.BoRcptTypes.rSupplier;
            oPayment.DocDate = docDate;
            oPayment.TaxDate = docDate;
            oPayment.ProjectCode = project;
            oPayment.CardCode = cardCode;
            oPayment.Remarks = headerLine["remarks"].ToString();

            if (string.IsNullOrEmpty(bpBankAccount))
                oPayment.IsPayToBank = SAPbobsCOM.BoYesNoEnum.tNO;
            else
            {
                oPayment.IsPayToBank = SAPbobsCOM.BoYesNoEnum.tYES;
                oPayment.PayToBankAccountNo = bpBankAccount;
                SAPbobsCOM.Recordset oRecordSet = CommonFunctions.getBPBankInfo(bpBankAccount, cardCode);
                if (oRecordSet != null)
                {
                    oPayment.PayToBankCountry = oRecordSet.Fields.Item("Country").Value;
                    oPayment.PayToBankCode = oRecordSet.Fields.Item("BankCode").Value;
                    Marshal.ReleaseComObject(oRecordSet);
                }
            }

            if (CommonFunctions.IsDevelopment())
            {
                string budgetCashFlowID = headerLine["BudgetCashFlowID"].ToString();
                string budgetCashFlowName = headerLine["BudgetCashFlowName"].ToString();

                if (!string.IsNullOrEmpty(budgetCashFlowID))
                {
                    oPayment.UserFields.Fields.Item("U_BDOSBdgCf").Value = budgetCashFlowID;
                    oPayment.UserFields.Fields.Item("U_BDOSBdgCfN").Value = budgetCashFlowName;
                }
            }

            decimal docRate;
            decimal docRateByBlnktAgr = 0;
            if (!string.IsNullOrEmpty(blnktAgr))
            {
                oPayment.BlanketAgreement = Convert.ToInt32(blnktAgr);
                oPayment.UserFields.Fields.Item("U_UseBlaAgRt").Value = useBlaAgRt;
                string docCur;
                if (useBlaAgRt == "Y")
                    docRateByBlnktAgr = Convert.ToDecimal(BlanketAgreement.GetBlAgremeentCurrencyRate(Convert.ToInt32(blnktAgr), out docCur, docDate), NumberFormatInfo.InvariantInfo);
            }

            if (invCurrency == localCurrency)
            {
                docRate = 0;
                transferSum = transferSumLC;
            }
            else
            {
                docRate = useBlaAgRt == "Y" ? docRateByBlnktAgr : Convert.ToDecimal(oSBOBob.GetCurrencyRate(invCurrency, docDate).Fields.Item("CurrencyRate").Value, NumberFormatInfo.InvariantInfo);
                transferSum = transferSumFC;
            }

            oPayment.DocCurrency = invCurrency;
            oPayment.LocalCurrency = oPayment.DocCurrency == localCurrency ? SAPbobsCOM.BoYesNoEnum.tYES : SAPbobsCOM.BoYesNoEnum.tNO;
            oPayment.DocRate = Convert.ToDouble(docRate, NumberFormatInfo.InvariantInfo);

            oPayment.TransferAccount = transferAccount;
            oPayment.TransferDate = docDate;
            oPayment.TransferSum = Convert.ToDouble(transferSum, NumberFormatInfo.InvariantInfo);

            //if (!string.IsNullOrEmpty(wtCode))
            //{
            //    oPayment.WTCode = wtCode;
            //    oPayment.WtBaseSum = Convert.ToDouble(transferSum, NumberFormatInfo.InvariantInfo);
            //    oPayment.WTAmount = Convert.ToDouble(wtAmount + pensEmployedAmount, NumberFormatInfo.InvariantInfo);
            //}

            oPayment.UserFields.Fields.Item("U_status").Value = oPayment.IsPayToBank == SAPbobsCOM.BoYesNoEnum.tYES ? "readyToLoad" : "notToUpload";
            oPayment.UserFields.Fields.Item("U_chrgDtls").Value = oPayment.IsPayToBank == SAPbobsCOM.BoYesNoEnum.tYES ? headerLine["ChrgDtls"].ToString() : "";
            oPayment.UserFields.Fields.Item("U_dsptchType").Value = oPayment.IsPayToBank == SAPbobsCOM.BoYesNoEnum.tYES ? headerLine["DispType"].ToString() : "";
            oPayment.UserFields.Fields.Item("U_descrpt").Value = headerLine["Descrpt"].ToString();
            oPayment.UserFields.Fields.Item("U_addDescrpt").Value = headerLine["AddDescrpt"].ToString();
            oPayment.UserFields.Fields.Item("U_BDOSWhtAmt").Value = Convert.ToDouble(wtAmount, NumberFormatInfo.InvariantInfo);
            oPayment.UserFields.Fields.Item("U_BDOSPnPhAm").Value = Convert.ToDouble(pensEmployedAmount, NumberFormatInfo.InvariantInfo);
            oPayment.UserFields.Fields.Item("U_BDOSPnCoAm").Value = Convert.ToDouble(pensEmployerAmount, NumberFormatInfo.InvariantInfo);

            decimal noDocSum = 0;
            //ცხრილური ნაწილი
            DataRow accountPaymentsLine;
            for (int i = 0; i < AccountPaymentsLines.Rows.Count; i++)
            {
                accountPaymentsLine = AccountPaymentsLines.Rows[i];

                if (accountPaymentsLine["DocEntry"].ToString() != "0")
                {
                    SAPbobsCOM.BoRcptInvTypes InvType;
                    int InvTypeInt = Convert.ToInt32(accountPaymentsLine["InvType"]);

                    if (InvTypeInt == 18)
                        InvType = SAPbobsCOM.BoRcptInvTypes.it_PurchaseInvoice;
                    else if (InvTypeInt == 204)
                        InvType = SAPbobsCOM.BoRcptInvTypes.it_PurchaseDownPayment;
                    else
                        InvType = SAPbobsCOM.BoRcptInvTypes.it_APCorrectionInvoice;

                    oPayment.Invoices.DocEntry = Convert.ToInt32(accountPaymentsLine["DocEntry"]);
                    oPayment.Invoices.InvoiceType = InvType;
                    oPayment.Invoices.InstallmentId = Convert.ToInt32(accountPaymentsLine["InstallmentId"]);

                    decimal sumApplied;
                    decimal sumAppliedLC = Convert.ToDecimal(accountPaymentsLine["SumApplied"], NumberFormatInfo.InvariantInfo);
                    decimal sumAppliedFC = oPayment.DocCurrency != localCurrency ? sumAppliedLC / Convert.ToDecimal(oPayment.DocRate, NumberFormatInfo.InvariantInfo) : 0;
                    decimal balanceDue = Convert.ToDecimal(accountPaymentsLine["BalanceDue"], NumberFormatInfo.InvariantInfo);

                    if (oPayment.DocCurrency == localCurrency)
                    {
                        sumApplied = Math.Min(balanceDue, sumAppliedLC);
                        oPayment.Invoices.SumApplied = Convert.ToDouble(sumApplied, NumberFormatInfo.InvariantInfo);
                        noDocSum += (sumAppliedLC - balanceDue);
                    }
                    else
                    {
                        sumApplied = Math.Min(balanceDue, sumAppliedFC);
                        oPayment.Invoices.AppliedFC = Convert.ToDouble(sumApplied, NumberFormatInfo.InvariantInfo);
                        noDocSum += (sumAppliedFC - balanceDue);
                    }

                    DataRow DTSourceRowVPM2 = DTSourceVPM2.Rows.Add();
                    DTSourceRowVPM2["DocEntry"] = Convert.ToInt32(accountPaymentsLine["DocEntry"]);
                    DTSourceRowVPM2["InvType"] = InvTypeInt;
                    DTSourceRowVPM2["SumApplied"] = oPayment.Invoices.SumApplied;
                    DTSourceRowVPM2["AppliedFC"] = oPayment.Invoices.AppliedFC;
                    oPayment.Invoices.Add();
                }
            }

            bool cashFlowRelevant = CommonFunctions.isAccountCashFlowRelevant(transferAccount);
            if (cashFlowRelevant)
            {
                oPayment.PrimaryFormItems.CashFlowLineItemID = Convert.ToInt32(headerLine["CashFlowID"]);
                oPayment.PrimaryFormItems.AmountFC = Convert.ToDouble(transferSumFC, NumberFormatInfo.InvariantInfo);
                if (oPayment.DocCurrency == localCurrency)
                    oPayment.PrimaryFormItems.AmountLC = Convert.ToDouble(transferSumLC, NumberFormatInfo.InvariantInfo);
                oPayment.PrimaryFormItems.PaymentMeans = SAPbobsCOM.PaymentMeansTypeEnum.pmtBankTransfer;
                oPayment.PrimaryFormItems.Add();
            }

            bool physicalEntityTax = (BusinessPartners.isWTLiable(cardCode) && CommonFunctions.getValue("OWHT", "U_BDOSPhisTx", "WTCode", wtCode).ToString() == "Y");

            DataRow DTSourceRow = DTSource.Rows.Add();

            //SAPbobsCOM.BusinessPartners oBP = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);
            //if (oBP.GetByKey(cardCode))
            //{
            //    wtCode = oBP.WTCode;
            //}
            DTSourceRow["WtCode"] = wtCode;
            DTSourceRow["WTLiable"] = BusinessPartners.isWTLiable(cardCode) ? "Y" : "N";
            DTSourceRow["CardCode"] = cardCode;
            DTSourceRow["PrjCode"] = project;
            DTSourceRow["U_liablePrTx"] = "N";
            DTSourceRow["U_prBase"] = "";
            if (noDocSum > 0)
                DTSourceRow["NoDocSum"] = Convert.ToDouble(noDocSum, NumberFormatInfo.InvariantInfo);
            DTSourceRow["U_BDOSWhtAmt"] = Convert.ToDouble(wtAmount, NumberFormatInfo.InvariantInfo);
            DTSourceRow["U_BDOSPnPhAm"] = Convert.ToDouble(pensEmployedAmount, NumberFormatInfo.InvariantInfo);
            DTSourceRow["U_BDOSPnCoAm"] = Convert.ToDouble(pensEmployerAmount, NumberFormatInfo.InvariantInfo);

            if (noDocSum > 0)
                oPayment.ControlAccount = headerLine["ControlAccount"].ToString();

            CommonFunctions.StartTransaction();

            int resultCode = oPayment.Add();

            if (resultCode != 0)
            {
                string errorMessage;
                if (Program.oCompany.InTransaction)
                {
                    CommonFunctions.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                }

                Program.oCompany.GetLastError(out resultCode, out errorMessage);
                Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("DocumentNotCreated") + ". " + BDOSResources.getTranslate("ReasonIs") + ": " + errorMessage, SAPbouiCOM.BoMessageTime.bmt_Short);
                return 0;
            }
            else
            {
                string docEntryS;
                Program.oCompany.GetNewObjectCode(out docEntryS);
                DataTable reLines;

                DataTable JrnLinesDT = OutgoingPayment.createAdditionalEntries(null, null, DTSource, DTSourceVPM2, oPayment.DocCurrency, out reLines, Convert.ToDecimal(oPayment.DocRate));
                OutgoingPayment.JrnEntry(docEntryS, docEntryS, oPayment.DocDate, JrnLinesDT, reLines, out errorText);

                if (!string.IsNullOrEmpty(errorText))
                {
                    if (Program.oCompany.InTransaction)
                    {
                        CommonFunctions.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                    }
                    Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("DocumentNotCreated") + ". " + BDOSResources.getTranslate("ReasonIs") + ": " + errorText, SAPbouiCOM.BoMessageTime.bmt_Short);
                }
                else
                {
                    if (Program.oCompany.InTransaction)
                    {
                        CommonFunctions.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                    }
                    Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("DocumentCreatedSuccesfully") + ": " + docEntryS, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                }

                return Convert.ToInt32(docEntryS);
            }
        }

        private static void fillBdgFlowItems(SAPbouiCOM.Form oForm)
        {
            try
            {
                oForm.Freeze(true);

                SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("InvoiceMTR").Specific;
                oMatrix.FlushToDataSource();

                SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("InvoiceMTR");

                string bCode = oForm.DataSources.UserDataSources.Item("BDOSDefCfE").ValueEx.Trim();
                string bName = UDO.GetUDOFieldValueByParam("UDO_F_BDOSBUCFW_D", "Code", bCode, "Name");

                for (int row = 0; row < oDataTable.Rows.Count; row++)
                {
                    oDataTable.SetValue("BudgetCashFlowID", row, bCode);
                    oDataTable.SetValue("BudgetCashFlowName", row, bName);
                }
                oMatrix.LoadFromDataSource();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                oForm.Update();
                oForm.Freeze(false);
            }
        }

        public static void uiApp_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                try
                {
                    SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_RESIZE && !pVal.BeforeAction)
                    {
                        resizeItems(oForm);
                    }

                    else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                    {
                        SAPbouiCOM.IChooseFromListEvent oCFLEvento = (SAPbouiCOM.IChooseFromListEvent)pVal;
                        chooseFromList(oForm, pVal, oCFLEvento);
                    }

                    else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK)
                    {
                        if (!pVal.BeforeAction)
                        {
                            if (pVal.ItemUID == "InCheck" || pVal.ItemUID == "InUncheck")
                            {
                                checkUncheckMTR(oForm, pVal.ItemUID);
                                showSelectedRowsTotalPayment(oForm);
                            }
                            else if (pVal.ItemUID == "fillBdgFl")
                                fillBdgFlowItems(oForm);
                            else if (pVal.ItemUID == "AddRow")
                                AddRow(oForm);
                            else if (pVal.ItemUID == "CreatDocmt")
                                createPaymentDocument(oForm);
                            else if (pVal.ItemUID == "delMTRB")
                                deleteMatrixRow(oForm);
                            else if (pVal.ItemUID == "InvoiceMTR" && pVal.ColUID == "UseBlaAgRt")
                            {
                                updateTotalPaymentRow(oForm, "TtlPmntFC", pVal.Row);
                                showSelectedRowsTotalPayment(oForm);
                            }
                            else if (pVal.ItemUID == "InvoiceMTR" && pVal.ColUID == "CheckBox")
                                showSelectedRowsTotalPayment(oForm);
                        }
                    }

                    else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_GOT_FOCUS)
                    {
                        if (!pVal.BeforeAction)
                        {
                            if (pVal.ItemUID == "InvoiceMTR" && pVal.ColUID == "BlnktAgr")
                            {
                                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(pVal.ItemUID).Specific;
                                blnktAgrOld = oMatrix.GetCellSpecific(pVal.ColUID, pVal.Row).Value;
                            }
                        }
                    }

                    else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_LOST_FOCUS)
                    {
                        if (!pVal.BeforeAction)
                        {
                            if (pVal.ItemUID == "InvoiceMTR" && pVal.ColUID == "BlnktAgr")
                            {
                                oForm.Freeze(true);
                                try
                                {
                                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(pVal.ItemUID).Specific;
                                    string blnktAgr = oMatrix.GetCellSpecific(pVal.ColUID, pVal.Row).Value;
                                    if (blnktAgr != blnktAgrOld && !string.IsNullOrEmpty(blnktAgrOld) && string.IsNullOrEmpty(blnktAgr))
                                    {
                                        int rowIndex = pVal.Row;

                                        SAPbouiCOM.CheckBox oCheckBox = oMatrix.Columns.Item("UseBlaAgRt").Cells.Item(rowIndex).Specific;
                                        oCheckBox.Checked = false;

                                        setMTRCellEditableSetting(oForm, pVal.ItemUID, rowIndex);
                                        blnktAgrOld = null;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    blnktAgrOld = null;
                                    throw new Exception(ex.Message);
                                }
                                finally
                                {
                                    oForm.Freeze(false);
                                }
                            }
                        }
                    }

                    else if (pVal.ItemChanged)
                    {
                        if (!pVal.BeforeAction)
                        {
                            if (pVal.ItemUID == "DocPstDt")
                                fillMTRInvoice(oForm);
                            else if (pVal.ItemUID == "Descrpt")
                                updateRow(oForm, false, true);
                            else if (pVal.ItemUID == "CashFlowI")
                                updateRow(oForm, true, false);
                            else if (pVal.ItemUID == "InvoiceMTR")
                            {
                                if (pVal.ColUID == "TtlPmntLC" || pVal.ColUID == "TtlPmntFC")
                                {
                                    updateTotalPaymentRow(oForm, pVal.ColUID, pVal.Row);
                                    showSelectedRowsTotalPayment(oForm);
                                    //if (pVal.ColUID == "TtlPmntLC")
                                    //    fillGrossAmount(oForm, pVal.ColUID, pVal.Row);
                                }
                            }
                            //else if (pVal.ItemUID == "WHTax")
                            //{
                            //    string whTaxCode = oForm.DataSources.UserDataSources.Item("WHTax").ValueEx;
                            //    string cardCode = oForm.DataSources.UserDataSources.Item("BPCode").ValueEx;
                            //    bool physicalEntityTax = (BusinessPartners.isWTLiable(cardCode) && CommonFunctions.getValue("OWHT", "U_BDOSPhisTx", "WTCode", whTaxCode).ToString() == "Y");
                            //    SAPbouiCOM.Matrix oMatrix = oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("InvoiceMTR").Specific;
                            //    oMatrix.FlushToDataSource();
                            //    if (oMatrix.RowCount > 0)
                            //    {
                            //        calculatePensionAmt(oForm, whTaxCode, physicalEntityTax);
                            //    }
                            //}
                        }
                    }

                    else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED)
                    {
                        if (pVal.ItemUID == "InvoiceMTR")
                            matrixColumnSetLinkedObjectTypeInvoicesMTR(oForm, pVal);
                    }
                }
                catch (Exception ex)
                {
                    Program.uiApp.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
            }
        }

        private static void checkDueAmount(SAPbouiCOM.Form oForm, int row, out bool BubbleEvent)
        {
            BubbleEvent = true;

            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("InvoiceMTR").Specific;
            SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("InvoiceMTR");

            row = row - 1;
            decimal TotalPymnt = Convert.ToDecimal(oMatrix.GetCellSpecific("TtlPmntLC", row + 1).Value, CultureInfo.InvariantCulture);
            decimal BalanceDue = Convert.ToDecimal(oMatrix.GetCellSpecific("BalanceDue", row + 1).Value, CultureInfo.InvariantCulture);
            if (BalanceDue < TotalPymnt)
            {
                TotalPymnt = Convert.ToDecimal(oDataTable.GetValue("TotalPaymentLC", row), CultureInfo.InvariantCulture);
                oMatrix.GetCellSpecific("TtlPmntLC", row + 1).Value = TotalPymnt;
            }
            else
            {
                oDataTable.SetValue("TotalPaymentLC", row, Convert.ToDouble(TotalPymnt, CultureInfo.InvariantCulture));

                oForm.Freeze(true);
                oMatrix.Clear();
                oMatrix.LoadFromDataSource();
                oMatrix.AutoResizeColumns();
                oForm.Freeze(false);
            }
        }

        //private static void fillGrossAmount(SAPbouiCOM.Form oForm, string Column, int row)
        //{
        //    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("InvoiceMTR").Specific;
        //    oMatrix.FlushToDataSource();
        //    SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("InvoiceMTR");

        //    row = row - 1;
        //    decimal totalPaymentNet = Convert.ToDecimal(oDataTable.GetValue("TotalPaymentNet", row), CultureInfo.InvariantCulture);
        //    decimal totalPaymentLC = Convert.ToDecimal(oDataTable.GetValue("TotalPaymentLC", row), CultureInfo.InvariantCulture);

        //    string WHTaxCode = oForm.DataSources.UserDataSources.Item("WHTax").ValueEx;
        //    DataTable WTaxDefinitons = WithholdingTax.getWtaxCodeDefinitionByDate(DateTime.Now);
        //    string filter;
        //    DataRow[] oWHTaxCode;
        //    decimal pensionRate = 0;

        //    SAPbobsCOM.WithholdingTaxCodes oWhTax = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oWithholdingTaxCodes);
        //    if (oWhTax.GetByKey(WHTaxCode))
        //    {
        //        if (oWhTax.UserFields.Fields.Item("U_BDOSPhisTx").Value == "Y")
        //        {
        //            string pensionCoWTCode = CommonFunctions.getOADM("U_BDOSPnCoP").ToString();
        //            filter = "WTCode = '" + pensionCoWTCode + "'";
        //            oWHTaxCode = WTaxDefinitons.Select(filter);
        //            pensionRate = 0;
        //            if (oWHTaxCode.Count() > 0)
        //            {
        //                pensionRate = Convert.ToDecimal(oWHTaxCode[0]["Rate"]);
        //            }
        //        }
        //    }

        //    decimal WTRate = 0;
        //    filter = "WTCode = '" + WHTaxCode + "'";
        //    oWHTaxCode = WTaxDefinitons.Select(filter);
        //    if (oWHTaxCode.Count() > 0)
        //        WTRate = Convert.ToDecimal(oWHTaxCode[0]["Rate"]);

        //    decimal pensSum;
        //    decimal WTSum;

        //    if (Column == "TtlPmntNt")
        //    {
        //        totalPaymentLC = totalPaymentNet / (1 - WTRate / 100) / (1 - pensionRate / 100);
        //        pensSum = totalPaymentLC * pensionRate / 100;
        //        WTSum = (totalPaymentLC - pensSum) * WTRate / 100;
        //    }
        //    else
        //    {
        //        pensSum = totalPaymentLC * pensionRate / 100;
        //        WTSum = (totalPaymentLC - pensSum) * WTRate / 100;
        //        totalPaymentNet = totalPaymentLC - pensSum - WTSum;
        //    }

        //    oDataTable.SetValue("TotalPaymentNet", row, Convert.ToDouble(totalPaymentNet, CultureInfo.InvariantCulture));
        //    oDataTable.SetValue("PensEmpld", row, Convert.ToDouble(pensSum, CultureInfo.InvariantCulture));
        //    oDataTable.SetValue("WTSum", row, Convert.ToDouble(WTSum, CultureInfo.InvariantCulture));
        //    oDataTable.SetValue("TotalPaymentLC", row, Convert.ToDouble(totalPaymentLC, CultureInfo.InvariantCulture));

        //    oForm.Freeze(true);
        //    oMatrix.Clear();
        //    oMatrix.LoadFromDataSource();
        //    oMatrix.AutoResizeColumns();
        //    oForm.Freeze(false);
        //}

        private static void SetInvDocsMatrixRowBackColor(SAPbouiCOM.Form oForm, int row)
        {
            try
            {
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("InvoiceMTR").Specific;

                if (oMatrix.RowCount > 0)
                {
                    oForm.Freeze(false);
                    for (int i = 1; i <= oMatrix.RowCount; i++)
                    {
                        oMatrix.CommonSetting.SetRowBackColor(i, FormsB1.getLongIntRGB(231, 231, 231));
                    }
                    oMatrix.CommonSetting.SetRowBackColor(row, FormsB1.getLongIntRGB(255, 255, 153));
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                oForm.Freeze(true);
                GC.Collect();
            }
        }

        private static void checkUncheckMTR(SAPbouiCOM.Form oForm, string checkOperation)
        {
            try
            {
                oForm.Freeze(true);

                SAPbouiCOM.CheckBox oCheckBox;
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("InvoiceMTR").Specific;

                for (int j = 1; j <= oMatrix.RowCount; j++)
                {
                    oCheckBox = oMatrix.Columns.Item("CheckBox").Cells.Item(j).Specific;
                    oCheckBox.Checked = (checkOperation == "InCheck");
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                oForm.Freeze(false);
                GC.Collect();
            }
        }

        private static void matrixColumnSetLinkedObjectTypeInvoicesMTR(SAPbouiCOM.Form oForm, SAPbouiCOM.ItemEvent pVal)
        {
            try
            {
                if (pVal.ColUID == "DocEntry")
                {
                    if (pVal.BeforeAction)
                    {
                        SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("InvoiceMTR").Specific;

                        SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("InvoiceMTR");
                        string docType = oDataTable.GetValue("DocType", pVal.Row - 1);

                        SAPbouiCOM.Column oColumn;

                        if (docType == "18")
                        {
                            oColumn = oMatrix.Columns.Item(pVal.ColUID);
                            SAPbouiCOM.LinkedButton oLink = oColumn.ExtendedObject;
                            oLink.LinkedObjectType = docType; //ARInvoice
                        }
                        if (docType == "204")
                        {
                            oColumn = oMatrix.Columns.Item(pVal.ColUID);
                            SAPbouiCOM.LinkedButton oLink = oColumn.ExtendedObject;
                            oLink.LinkedObjectType = docType; //ARInvoice
                        }
                        else if (docType == "163")
                        {
                            oColumn = oMatrix.Columns.Item(pVal.ColUID);
                            SAPbouiCOM.LinkedButton oLink = oColumn.ExtendedObject;
                            oLink.LinkedObjectType = docType; //ARCreditNote
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                GC.Collect();
            }
        }

        private static void AddRow(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("InvoiceMTR");
            int rowIndex = oDataTable.Rows.Count;

            SAPbouiCOM.EditText oEditTextDocDate = (SAPbouiCOM.EditText)oForm.Items.Item("DocPstDt").Specific;
            string DocDateS = oEditTextDocDate.Value;
            DateTime DocDate = Convert.ToDateTime(DateTime.ParseExact(DocDateS, "yyyyMMdd", CultureInfo.InvariantCulture));

            string GLAccount = oForm.Items.Item("GLAcc").Specific.Value;

            SAPbobsCOM.ChartOfAccounts oChartOfAccounts = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oChartOfAccounts);
            oChartOfAccounts.GetByKey(GLAccount);

            string Currency = oChartOfAccounts.AcctCurrency;
            if (Currency == "##")
                Currency = Program.LocalCurrency;

            oDataTable.Rows.Add();
            oDataTable.SetValue("LineNum", rowIndex, rowIndex + 1);
            oDataTable.SetValue("CheckBox", rowIndex, "N");
            oDataTable.SetValue("DocEntry", rowIndex, 0);
            oDataTable.SetValue("DocNum", rowIndex, 0);
            oDataTable.SetValue("InstallmentID", rowIndex, 0);
            oDataTable.SetValue("LineID", rowIndex, 0);
            oDataTable.SetValue("DocType", rowIndex, "");
            oDataTable.SetValue("DocDate", rowIndex, DocDate);
            oDataTable.SetValue("DueDate", rowIndex, DocDate);
            oDataTable.SetValue("Arrears", rowIndex, "");
            oDataTable.SetValue("OverdueDays", rowIndex, 0);
            oDataTable.SetValue("Comments", rowIndex, "");
            oDataTable.SetValue("Total", rowIndex, 0);
            oDataTable.SetValue("WTSum", rowIndex, 0);
            oDataTable.SetValue("BalanceDue", rowIndex, 0);
            oDataTable.SetValue("TotalPaymentLC", rowIndex, 0);
            oDataTable.SetValue("TotalPaymentFC", rowIndex, 0);
            oDataTable.SetValue("Currency", rowIndex, Currency);
            //oDataTable.SetValue("TotalPaymentLocal", rowIndex, 0);
            oDataTable.SetValue("Project", rowIndex, "");

            if (CommonFunctions.IsDevelopment())
            {
                string bCode = oForm.DataSources.UserDataSources.Item("BDOSDefCfE").ValueEx.Trim();
                string bName = UDO.GetUDOFieldValueByParam("UDO_F_BDOSBUCFW_D", "Code", bCode, "Name");
                bName = bName == null ? "" : bName;

                oDataTable.SetValue("BudgetCashFlowID", rowIndex, bCode);
                oDataTable.SetValue("BudgetCashFlowName", rowIndex, bName);
            }

            oDataTable.SetValue("CFWId", rowIndex, oForm.DataSources.UserDataSources.Item("CashFlowI").ValueEx);
            oDataTable.SetValue("Description", rowIndex, oForm.DataSources.UserDataSources.Item("Descrpt").ValueEx);

            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("InvoiceMTR").Specific));
            oForm.Freeze(true);
            oMatrix.Clear();
            oMatrix.LoadFromDataSource();
            oMatrix.AutoResizeColumns();

            setEditableSetting(oForm);

            oForm.Update();
            oForm.Freeze(false);
        }

        public static void deleteMatrixRow(SAPbouiCOM.Form oForm)
        {
            oForm.Freeze(true);
            try
            {
                SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("InvoiceMTR").Specific));
                oMatrix.FlushToDataSource();
                int firstRow = 0;
                int row = 0;
                int deletedRowCount = 0;

                SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("InvoiceMTR");

                while (row != -1)
                {
                    row = oMatrix.GetNextSelectedRow(firstRow, SAPbouiCOM.BoOrderType.ot_RowOrder);
                    if (row > -1)
                    {
                        deletedRowCount++;
                        oDataTable.Rows.Remove(row - deletedRowCount);
                        firstRow = row;
                    }
                }

                int rowCount = oDataTable.Rows.Count;

                for (int i = 1; i <= rowCount; i++)
                {
                    string docEntry = oDataTable.GetValue("DocEntry", i - 1).ToString();
                    if (!string.IsNullOrEmpty(docEntry))
                    {
                        oDataTable.SetValue("LineNum", i - 1, i.ToString());
                    }
                }
                oMatrix.LoadFromDataSource();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                GC.Collect();
                oForm.Freeze(false);
            }
        }

        private static void createPaymentDocument(SAPbouiCOM.Form oForm)
        {
            int answer = Program.uiApp.MessageBox(BDOSResources.getTranslate("CreatePaymentDocuments") + "?", 1, BDOSResources.getTranslate("Yes"), BDOSResources.getTranslate("No"), "");

            if (answer == 2)
            {
                return;
            }

            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("InvoiceMTR").Specific;
            oMatrix.FlushToDataSource();

            string docDateS = oForm.DataSources.UserDataSources.Item("DocPstDt").ValueEx;
            string cardCode = oForm.DataSources.UserDataSources.Item("BPCode").ValueEx;
            string bpBankAccount = oForm.DataSources.UserDataSources.Item("BPBankAcc").ValueEx;
            string glAccount = oForm.DataSources.UserDataSources.Item("GLAcc").ValueEx;
            string controlAccount = oForm.DataSources.UserDataSources.Item("CTAcc").ValueEx;
            string dispType = oForm.DataSources.UserDataSources.Item("DispType").ValueEx;
            string chrgDtls = oForm.DataSources.UserDataSources.Item("ChrgDtls").ValueEx;

            int docEntry;
            string prevCurrency = null;
            string prevWTCode = null;
            string prevProject = null;
            //string prevDocIsEmpty = null;
            string prevBudgetCashFlowID = null;
            string prevUseBlaAgRt = null;
            string prevBlnktAgr = null;
            string prevCashFlowID = null;
            string prevDescription = null;
            decimal payblAmtFCTotal = 0;
            decimal payblAmtLCTotal = 0;
            decimal wtAmtTotal = 0;
            decimal pensEmployedTotal = 0;
            decimal pensEmployerTotal = 0;

            string errorText;

            if (string.IsNullOrEmpty(docDateS) || string.IsNullOrEmpty(cardCode) || string.IsNullOrEmpty(glAccount))
            {
                errorText = BDOSResources.getTranslate("TheFollowingFieldsAreMandatory")
                                + " : \"" + BDOSResources.getTranslate("CardCode")
                                + "\", \"" + BDOSResources.getTranslate("DocumentPostingDate")
                                + "\", \"" + BDOSResources.getTranslate("GLAccount") + "\"";
                Program.uiApp.StatusBar.SetSystemMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short);
                return;
            }
            if (!string.IsNullOrEmpty(bpBankAccount) && (string.IsNullOrEmpty(dispType) || string.IsNullOrEmpty(chrgDtls)))
            {
                errorText = BDOSResources.getTranslate("TheFollowingFieldsAreMandatory")
                                + " : \"" + BDOSResources.getTranslate("ChrgDtls")
                                + "\", \"" + BDOSResources.getTranslate("DispatchType") + "\"";
                Program.uiApp.StatusBar.SetSystemMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short);
                return;
            }

            DateTime docDate = Convert.ToDateTime(DateTime.ParseExact(docDateS, "yyyyMMdd", CultureInfo.InvariantCulture));
            SAPbobsCOM.ChartOfAccounts oChartOfAccounts = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oChartOfAccounts);
            oChartOfAccounts.GetByKey(glAccount);

            string payblCur = oChartOfAccounts.AcctCurrency;
            //string whTaxCode = oForm.DataSources.UserDataSources.Item("WHTax").ValueEx;

            DataTable accountHeader = new DataTable();
            DataRow headerLine = accountHeader.Rows.Add();

            accountHeader.Columns.Add("CardCode");
            accountHeader.Columns.Add("Currency");
            accountHeader.Columns.Add("PayblCur");

            DataColumn colDecimal = new DataColumn("PayblCRt");
            colDecimal.DataType = Type.GetType("System.Decimal");
            accountHeader.Columns.Add(colDecimal);

            colDecimal = new DataColumn("PayblAmtLC");
            colDecimal.DataType = Type.GetType("System.Decimal");
            accountHeader.Columns.Add(colDecimal);

            colDecimal = new DataColumn("PayblAmtFC");
            colDecimal.DataType = Type.GetType("System.Decimal");
            accountHeader.Columns.Add(colDecimal);

            colDecimal = new DataColumn("pensEmployedAmount");
            colDecimal.DataType = Type.GetType("System.Decimal");
            accountHeader.Columns.Add(colDecimal);

            colDecimal = new DataColumn("pensEmployerAmount");
            colDecimal.DataType = Type.GetType("System.Decimal");
            accountHeader.Columns.Add(colDecimal);

            colDecimal = new DataColumn("WtAmount");
            colDecimal.DataType = Type.GetType("System.Decimal");
            accountHeader.Columns.Add(colDecimal);

            accountHeader.Columns.Add("BankAccount");
            accountHeader.Columns.Add("TransferAccount");
            accountHeader.Columns.Add("ControlAccount");
            accountHeader.Columns.Add("accrualDate");
            accountHeader.Columns.Add("CashFlowID");
            accountHeader.Columns.Add("remarks");
            accountHeader.Columns.Add("ChrgDtls");
            accountHeader.Columns.Add("DispType");
            accountHeader.Columns.Add("Descrpt");
            accountHeader.Columns.Add("AddDescrpt");
            accountHeader.Columns.Add("Project");
            accountHeader.Columns.Add("WTCode");
            accountHeader.Columns.Add("BudgetCashFlowID");
            accountHeader.Columns.Add("BudgetCashFlowName");
            accountHeader.Columns.Add("UseBlaAgRt");
            accountHeader.Columns.Add("BlnktAgr");

            DataTable accountPaymentsLines = new DataTable();

            accountPaymentsLines.Columns.Add("InvType");
            accountPaymentsLines.Columns.Add("DocEntry");
            accountPaymentsLines.Columns.Add("DocNum");
            accountPaymentsLines.Columns.Add("InstallmentId");

            colDecimal = new DataColumn("SumApplied");
            colDecimal.DataType = Type.GetType("System.Decimal");
            accountPaymentsLines.Columns.Add(colDecimal);

            colDecimal = new DataColumn("BalanceDue");
            colDecimal.DataType = Type.GetType("System.Decimal");
            accountPaymentsLines.Columns.Add(colDecimal);

            SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("InvoiceMTR");

            for (int i = 0; i < oDataTable.Rows.Count; i++)
            {
                if (oDataTable.GetValue("CheckBox", i) == "Y")
                {
                    string description = oDataTable.GetValue("Description", i);
                    if (string.IsNullOrEmpty(description))
                    {
                        Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("DescriptionIsMandatory") + "! " + BDOSResources.getTranslate("TableRow") + ": " + (i + 1) + "! ", SAPbouiCOM.BoMessageTime.bmt_Short);
                        return;
                    }
                    string budgetCashFlowID = null;
                    if (CommonFunctions.IsDevelopment())
                        budgetCashFlowID = oDataTable.GetValue("BudgetCashFlowID", i);
                    string currency = oDataTable.GetValue("Currency", i);
                    decimal totalPaymentLC = Convert.ToDecimal(oDataTable.GetValue("TotalPaymentLC", i), NumberFormatInfo.InvariantInfo);
                    decimal balanceDue = Convert.ToDecimal(oDataTable.GetValue("BalanceDue", i), NumberFormatInfo.InvariantInfo);
                    decimal totalPaymentFC = Convert.ToDecimal(oDataTable.GetValue("TotalPaymentFC", i), NumberFormatInfo.InvariantInfo);
                    decimal wtAmt = Convert.ToDecimal(oDataTable.GetValue("WTSum", i), NumberFormatInfo.InvariantInfo);
                    decimal pensEmployed = Convert.ToDecimal(oDataTable.GetValue("PensEmpld", i), NumberFormatInfo.InvariantInfo);
                    decimal pensEmployer = Convert.ToDecimal(oDataTable.GetValue("PensEmplr", i), NumberFormatInfo.InvariantInfo);
                    string project = oDataTable.GetValue("Project", i);
                    string wtCode = oDataTable.GetValue("WTCode", i);
                    string invType = oDataTable.GetValue("DocType", i);
                    string invDocEntry = oDataTable.GetValue("DocEntry", i).ToString();
                    //string docIsEmpty = oDataTable.GetValue("DocEntry", i).ToString();
                    string invDocNum = oDataTable.GetValue("DocNum", i).ToString();
                    string installmentId = oDataTable.GetValue("InstallmentID", i).ToString();
                    string cashFlowIDStr = oDataTable.GetValue("CFWId", i).ToString();
                    int cashFlowID = string.IsNullOrEmpty(cashFlowIDStr) ? 0 : Convert.ToInt32(cashFlowIDStr);
                    string useBlaAgRt = oDataTable.GetValue("UseBlaAgRt", i);
                    string blnktAgr = oDataTable.GetValue("BlnktAgr", i);

                    if (totalPaymentLC == 0)
                        continue;

                    if (prevProject != project
                        || prevWTCode != wtCode
                        || prevCurrency != currency
                        //|| prevDocIsEmpty != docIsEmpty
                        || prevUseBlaAgRt != useBlaAgRt
                        || prevBlnktAgr != blnktAgr
                        || prevCashFlowID != cashFlowIDStr
                        || prevDescription != description
                        || (CommonFunctions.IsDevelopment() && prevBudgetCashFlowID != budgetCashFlowID))
                    {
                        if (prevProject != null)
                        {
                            headerLine["PayblAmtLC"] = payblAmtLCTotal;
                            headerLine["PayblAmtFC"] = payblAmtFCTotal;
                            headerLine["WtAmount"] = wtAmtTotal;
                            headerLine["pensEmployedAmount"] = pensEmployedTotal;
                            headerLine["pensEmployerAmount"] = pensEmployerTotal;
                            //გაკეთდება დოკუმენტი
                            try
                            {
                                docEntry = createPaymentDocument(oForm, headerLine, accountPaymentsLines);
                            }
                            catch (Exception ex)
                            {
                                Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("DocumentNotCreated") + ". " + BDOSResources.getTranslate("ReasonIs") + ": " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short);
                            }
                        }
                        accountHeader.Rows.Clear();
                        headerLine = accountHeader.Rows.Add();
                        //string whTaxCode = "";
                        headerLine["CardCode"] = oForm.DataSources.UserDataSources.Item("BPCode").ValueEx;
                        headerLine["Currency"] = currency;
                        headerLine["PayblCur"] = payblCur;
                        headerLine["PayblCRt"] = 1;
                        headerLine["BankAccount"] = bpBankAccount;
                        headerLine["TransferAccount"] = glAccount;
                        headerLine["ControlAccount"] = controlAccount;
                        headerLine["WTCode"] = wtCode;
                        headerLine["accrualDate"] = docDate;
                        headerLine["CashFlowID"] = cashFlowID;
                        headerLine["DispType"] = dispType;
                        headerLine["ChrgDtls"] = chrgDtls;
                        headerLine["Project"] = project;

                        if (CommonFunctions.IsDevelopment())
                        {
                            headerLine["BudgetCashFlowID"] = budgetCashFlowID;
                            headerLine["BudgetCashFlowName"] = UDO.GetUDOFieldValueByParam("UDO_F_BDOSBUCFW_D", "Code", budgetCashFlowID, "Name");
                        }

                        headerLine["Descrpt"] = description;
                        headerLine["UseBlaAgRt"] = useBlaAgRt;
                        headerLine["BlnktAgr"] = blnktAgr;

                        payblAmtLCTotal = 0;
                        payblAmtFCTotal = 0;
                        wtAmtTotal = 0;
                        pensEmployedTotal = 0;
                        pensEmployerTotal = 0;

                        accountPaymentsLines.Rows.Clear();
                    }

                    DataRow AccountPaymentsRow = accountPaymentsLines.Rows.Add();

                    AccountPaymentsRow["InvType"] = invType;
                    AccountPaymentsRow["DocEntry"] = invDocEntry;
                    AccountPaymentsRow["DocNum"] = invDocNum;
                    AccountPaymentsRow["InstallmentId"] = installmentId;
                    AccountPaymentsRow["SumApplied"] = totalPaymentLC;
                    AccountPaymentsRow["BalanceDue"] = balanceDue;

                    payblAmtLCTotal += totalPaymentLC;
                    payblAmtFCTotal += totalPaymentFC;
                    wtAmtTotal += wtAmt;
                    pensEmployedTotal += pensEmployed;
                    pensEmployerTotal += pensEmployer;

                    prevCurrency = currency;
                    prevProject = project;
                    prevWTCode = wtCode;
                    //prevDocIsEmpty = docIsEmpty;
                    prevUseBlaAgRt = useBlaAgRt;
                    prevBlnktAgr = blnktAgr;
                    prevCashFlowID = cashFlowIDStr;
                    prevDescription = description;

                    if (CommonFunctions.IsDevelopment())
                    {
                        prevBudgetCashFlowID = budgetCashFlowID;
                    }
                }
            }

            if (payblAmtLCTotal > 0)
            {
                headerLine["PayblAmtLC"] = payblAmtLCTotal;
                headerLine["PayblAmtFC"] = payblAmtFCTotal;
                headerLine["WtAmount"] = wtAmtTotal;
                headerLine["pensEmployedAmount"] = pensEmployedTotal;
                headerLine["pensEmployerAmount"] = pensEmployerTotal;
                try
                {
                    docEntry = createPaymentDocument(oForm, headerLine, accountPaymentsLines);
                }
                catch (Exception ex)
                {
                    Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("DocumentNotCreated") + ". " + BDOSResources.getTranslate("ReasonIs") + ": " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short);
                    CommonFunctions.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }
            }
            fillMTRInvoice(oForm);
        }

        private static void chooseFromList(SAPbouiCOM.Form oForm, SAPbouiCOM.ItemEvent pVal, SAPbouiCOM.IChooseFromListEvent oCFLEvento)
        {
            try
            {
                oForm.Freeze(true);

                if (pVal.BeforeAction)
                {
                    if (oCFLEvento.ChooseFromListUID == "BlnktAgr_CFL")
                    {
                        if (pVal.ItemUID == "InvoiceMTR" && pVal.ColUID == "BlnktAgr")
                        {
                            SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID);
                            SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();

                            SAPbouiCOM.Condition oCon = oCons.Add();
                            oCon.Alias = "BPCode";
                            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                            oCon.CondVal = oForm.DataSources.UserDataSources.Item("BPCode").ValueEx;

                            oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                            SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("InvoiceMTR");

                            oCon = oCons.Add();
                            oCon.Alias = "BPCurr";
                            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                            oCon.CondVal = oDataTable.GetValue("Currency", pVal.Row - 1);

                            oCFL.SetConditions(oCons);
                        }
                    }
                    else if (oCFLEvento.ChooseFromListUID == "BPBankAcc_CFL")
                    {
                        SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID);
                        SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();

                        SAPbouiCOM.Condition oCon = oCons.Add();
                        oCon.Alias = "CardCode";
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = oForm.DataSources.UserDataSources.Item("BPCode").ValueEx;

                        oCFL.SetConditions(oCons);
                    }
                }
                else
                {
                    SAPbouiCOM.DataTable oDataTable = oCFLEvento.SelectedObjects;

                    if (oDataTable != null)
                    {
                        if (oCFLEvento.ChooseFromListUID == "BusinessPartner_CFL")
                        {
                            string CardCode = Convert.ToString(oDataTable.GetValue("CardCode", 0));
                            LanguageUtils.IgnoreErrors<string>(() => oForm.Items.Item("BPCode").Specific.Value = CardCode);

                            //setWhtCodes(oForm);
                            fillMTRInvoice(oForm);
                        }
                        else if (oCFLEvento.ChooseFromListUID == "Budg_CFLHD")
                        {
                            string BCFWId = Convert.ToString(oDataTable.GetValue("Code", 0));
                            string BCFWName = Convert.ToString(oDataTable.GetValue("Name", 0));

                            oForm.DataSources.UserDataSources.Item("BDOSDefCfE").ValueEx = BCFWId;
                            oForm.DataSources.UserDataSources.Item("BDOSDefCfN").ValueEx = BCFWName;
                        }
                        else if (oCFLEvento.ChooseFromListUID == "Proj_CFL")
                        {
                            string PrjCode = Convert.ToString(oDataTable.GetValue("PrjCode", 0));
                            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("InvoiceMTR").Specific;
                            LanguageUtils.IgnoreErrors<string>(() => oMatrix.Columns.Item("Project").Cells.Item(pVal.Row).Specific.Value = PrjCode);
                        }
                        else if (oCFLEvento.ChooseFromListUID == "Budg_CFL")
                        {
                            string BCFWId = Convert.ToString(oDataTable.GetValue("Code", 0));
                            string BCFWName = Convert.ToString(oDataTable.GetValue("Name", 0));
                            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("InvoiceMTR").Specific;
                            LanguageUtils.IgnoreErrors<string>(() => oMatrix.Columns.Item("BCFWId").Cells.Item(pVal.Row).Specific.Value = BCFWId);
                            LanguageUtils.IgnoreErrors<string>(() => oMatrix.Columns.Item("BCFWName").Cells.Item(pVal.Row).Specific.Value = BCFWName);
                        }
                        else if (oCFLEvento.ChooseFromListUID == "BPBankAcc_CFL")
                        {
                            string Account = Convert.ToString(oDataTable.GetValue("Account", 0));
                            LanguageUtils.IgnoreErrors<string>(() => oForm.Items.Item("BPBankAcc").Specific.Value = Account);
                        }
                        else if (oCFLEvento.ChooseFromListUID == "GLAcc_CFL")
                        {
                            string GLAccount = Convert.ToString(oDataTable.GetValue("AcctCode", 0));
                            LanguageUtils.IgnoreErrors<string>(() => oForm.Items.Item("GLAcc").Specific.Value = GLAccount);
                        }
                        else if (oCFLEvento.ChooseFromListUID == "CTAcc_CFL")
                        {
                            string CTAccount = Convert.ToString(oDataTable.GetValue("AcctCode", 0));
                            LanguageUtils.IgnoreErrors<string>(() => oForm.Items.Item("CTAcc").Specific.Value = CTAccount);
                        }
                        else if (oCFLEvento.ChooseFromListUID == "BlnktAgr_CFL")
                        {
                            if (pVal.ItemUID == "InvoiceMTR" && pVal.ColUID == "BlnktAgr")
                            {
                                string absID = Convert.ToString(oDataTable.GetValue("AbsID", 0));

                                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(pVal.ItemUID).Specific;
                                LanguageUtils.IgnoreErrors<string>(() => oMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value = absID);
                                if (!string.IsNullOrEmpty(absID) && !BlanketAgreement.UsesCurrencyExchangeRates(Convert.ToInt32(absID)))
                                {
                                    SAPbouiCOM.CheckBox oCheckBox = oMatrix.Columns.Item("UseBlaAgRt").Cells.Item(pVal.Row).Specific;
                                    oCheckBox.Checked = false;
                                }
                                setMTRCellEditableSetting(oForm, pVal.ItemUID, pVal.Row);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                oForm.Update();
                oForm.Freeze(false);
            }
        }

        private static void setMTRCellEditableSetting(SAPbouiCOM.Form oForm, string mtrName, int rowIndex = 0)
        {
            try
            {
                oForm.Freeze(true);

                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(mtrName).Specific;
                int rowCount = rowIndex == 0 ? oMatrix.RowCount : rowIndex;
                int i = rowIndex == 0 ? 1 : rowIndex;

                for (; i <= rowCount; i++)
                {
                    string absID = oMatrix.GetCellSpecific("BlnktAgr", i).Value;
                    if (!string.IsNullOrEmpty(absID) && BlanketAgreement.UsesCurrencyExchangeRates(Convert.ToInt32(absID)))
                    {
                        oMatrix.CommonSetting.SetCellEditable(i, 21, true);
                    }
                    else
                    {
                        oMatrix.CommonSetting.SetCellEditable(i, 21, false);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                oForm.Update();
                oForm.Freeze(false);
            }
        }

        private static void setWhtCodes(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.ComboBox oItem = oForm.Items.Item("WHTax").Specific;

            try
            {
                while (oItem.ValidValues.Count > 0)
                {
                    oItem.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
                }

                string cardCode = oForm.DataSources.UserDataSources.Item("BPCode").ValueEx;
                SAPbobsCOM.BusinessPartners oBP = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);
                if (oBP.GetByKey(cardCode))
                {
                    int ln = 0;
                    while (ln < oBP.BPWithholdingTax.Count)
                    {
                        oBP.BPWithholdingTax.SetCurrentLine(ln);
                        oItem.ValidValues.Add(oBP.BPWithholdingTax.WTCode, oBP.BPWithholdingTax.WTCode);
                        ln++;
                    }
                    oItem.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private static void fillMTRInvoice(SAPbouiCOM.Form oForm)
        {
            string docDateS = oForm.DataSources.UserDataSources.Item("DocPstDt").ValueEx;
            DateTime date = Convert.ToDateTime(DateTime.ParseExact(docDateS, "yyyyMMdd", CultureInfo.InvariantCulture));
            SAPbobsCOM.SBObob oSBOBob = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);

            string dateE = date.ToString("yyyyMMdd");
            string cardCode = oForm.DataSources.UserDataSources.Item("BPCode").ValueEx;
            SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("InvoiceMTR");
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string betweenDays;

            if (Program.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
            {
                betweenDays = @"DAYS_BETWEEN (T0.""DueDate"", '" + date.ToString("yyyy-MM-dd") + @"') ";
            }
            else
            {
                betweenDays = @"DATEDIFF(DAY, T0.""DueDate"", '" + date.ToString("yyyy-MM-dd") + @"') ";
            }

            string query = @"SELECT
            	 T0.""DocEntry"" AS ""DocEntry"",
                 T0.""Project"",
	             T0.""DocNum"" AS ""DocNum"",
                 T0.""DocCur"" AS ""DocCur"",
            	 T0.""CardCode"" AS ""CardCode"",
            	 T0.""CardName"" AS ""CardName"",
            	 T0.""DocDate"" AS ""DocDate"",
            	 T0.""DueDate"" AS ""DueDate"",
            	 T0.""OpenAmount"" AS ""OpenAmount"",
            	 T0.""InsTotal"" AS ""InsTotal"",
            	 T0.""OpenAmountFC"" AS ""OpenAmountFC"",
                 T0.""WTSum"" AS ""WTSum"",
                 T0.""WTSumFC"" AS ""WTSumFC"",
            	 T0.""InsTotalFC"" AS ""InsTotalFC"",
            	 T0.""ObjType"" AS ""ObjType"",
            	 T0.""Comments"" AS ""Comments"",
                 T0.""InstlmntID"" AS ""InstallmentID"",
                 T0.""WTCode"",
                 T0.""LineID"" AS ""LineID"","
                +
                   betweenDays
                + @"AS ""OverdueDays"" 
            FROM ( SELECT
            	 TT0.""DocEntry"",
                 TT0.""Project"",
            	 TT0.""DocNum"" AS ""DocNum"",
                 TT0.""DocCur"" AS ""DocCur"",
            	 T3.""CardCode"" AS ""CardCode"",
            	 T3.""CardName"" AS ""CardName"",
            	 TT0.""DocDate"" AS ""DocDate"",
	             TT1.""DueDate"" AS ""DueDate"",
            	 TT0.""ObjType"" AS ""ObjType"",
            	 TT0.""Comments"" AS ""Comments"",
                 TT1.""InstlmntID"" AS ""InstlmntID"",
                 TT2.""WTCode"",
                 '0' AS ""LineID"",           	 
            	 SUM(TT1.""InsTotal"" - TT1.""PaidToDate""-TT1.""WTSum""+TT1.""WTApplied"") AS ""OpenAmount"",
                SUM(TT1.""WTSum"" - TT1.""WTApplied"") AS ""WTSum"",            	 
                SUM(TT1.""WTSumFC"" -TT1.""WTAppliedF"") AS ""WTSumFC"",            	 
                SUM(TT1.""InsTotal"") AS ""InsTotal"",
                 SUM(TT1.""InsTotalFC"" - TT1.""PaidFC""-TT1.""WTSumFC""+TT1.""WTAppliedF"") AS ""OpenAmountFC"",
            	 SUM(TT1.""InsTotalFC"") AS ""InsTotalFC"" 
            	FROM OPCH TT0 
            	INNER JOIN PCH6 TT1 ON TT0.""DocEntry"" = TT1.""DocEntry""
                LEFT JOIN PCH5 TT2 ON (TT0.""DocEntry"" = TT2.""AbsEntry"" AND TT0.""ObjType"" = TT2.""ObjType"")
            	INNER JOIN OCRD T3 ON TT0.""CardCode"" = T3.""CardCode"" 
            	WHERE TT0.""DocDate"" <= '" + dateE + @"' 
	            AND TT0.""CardCode"" = N'" + cardCode + @"' 
            	AND (TT0.""DocStatus"" = 'O' 
            		OR (TT1.""Status"" = 'O' 
            			AND TT0.""CANCELED"" = 'N')) 
            	GROUP BY TT0.""DocEntry"",
            	 TT0.""Project"",
            	 TT0.""DocNum"",
                 TT0.""DocCur"",
            	 T3.""CardCode"",
            	 T3.""CardName"",
            	 TT0.""DocDate"",
            	 TT1.""DueDate"",
	             TT0.""ObjType"",
            	 TT0.""Comments"",
                 TT1.""InstlmntID"", 
                 TT2.""WTCode""
            	UNION ALL SELECT
            	 TT0.""DocEntry"",
            	 TT0.""Project"",
            	 TT0.""DocNum"" AS ""DocNum"",
                 TT0.""DocCur"" AS ""DocCur"",
            	 T3.""CardCode"" AS ""CardCode"",
            	 T3.""CardName"" AS ""CardName"",
            	 TT0.""DocDate"" AS ""DocDate"",
            	 TT1.""DueDate"" AS ""DueDate"",
            	 TT0.""ObjType"" AS ""ObjType"",
            	 TT0.""Comments"" AS ""Comments"",
                 TT1.""InstlmntID"" AS ""InstlmntID"", 
                 TT2.""WTCode"",
                 '0' AS ""LineID"",
            	 -SUM(TT1.""InsTotal"" - TT1.""PaidToDate""-TT1.""WTSum""+TT1.""WTApplied"")*-1 AS ""OpenAmount"",
                SUM(TT1.""WTSum"" - TT1.""WTApplied"") AS ""WTSum"",            	 
                SUM(TT1.""WTSumFC"" -TT1.""WTAppliedF"") AS ""WTSumFC"", 
            	 -SUM(TT1.""InsTotal"")*-1 AS ""InsTotal"",
                 -SUM(TT1.""InsTotalFC"" - TT1.""PaidFC""-TT1.""WTSumFC""+TT1.""WTAppliedF"")*-1 AS ""OpenAmountFC"",
            	 -SUM(TT1.""InsTotalFC"")*-1 AS ""InsTotalFC""
            	FROM OCPI TT0 
            	INNER JOIN CPI6 TT1 ON TT0.""DocEntry"" = TT1.""DocEntry""
                LEFT JOIN CPI5 TT2 ON (TT0.""DocEntry"" = TT2.""AbsEntry"" AND TT0.""ObjType"" = TT2.""ObjType"")
            	INNER JOIN OCRD T3 ON TT0.""CardCode"" = T3.""CardCode"" 
            	WHERE  TT0.""DocDate"" <= '" + dateE + @"' 
            	AND TT0.""CardCode"" = N'" + cardCode + @"'
            	AND (TT0.""DocStatus"" = 'O' 
            		OR (TT1.""Status"" = 'O' 
            			AND TT0.""CANCELED"" = 'N')) 
            	GROUP BY TT0.""DocEntry"",
            	 TT0.""Project"",
            	 TT0.""DocNum"",
                 TT0.""DocCur"",
            	 T3.""CardCode"",
            	 T3.""CardName"",
            	 TT0.""DocDate"",
            	 TT1.""DueDate"",
            	 TT0.""ObjType"",
            	 TT0.""Comments"",
                 TT1.""InstlmntID"",
                 TT2.""WTCode""
                 UNION ALL SELECT
            	 TT0.""DocEntry"",
            	 TT0.""Project"",
            	 TT0.""DocNum"" AS ""DocNum"",
                 TT0.""DocCur"" AS ""DocCur"",
            	 T3.""CardCode"" AS ""CardCode"",
            	 T3.""CardName"" AS ""CardName"",
            	 TT0.""DocDate"" AS ""DocDate"",
            	 TT1.""DueDate"" AS ""DueDate"",
            	 TT0.""ObjType"" AS ""ObjType"",
            	 TT0.""Comments"" AS ""Comments"",
                 TT1.""InstlmntID"" AS ""InstlmntID"",
                 TT2.""WTCode"",
                 '0' AS ""LineID"",
            	 -SUM(TT1.""InsTotal"" - TT1.""PaidToDate""-TT1.""WTSum""+TT1.""WTApplied"")*-1 AS ""OpenAmount"",
                SUM(TT1.""WTSum"" - TT1.""WTApplied"") AS ""WTSum"",            	 
                SUM(TT1.""WTSumFC"" -TT1.""WTAppliedF"") AS ""WTSumFC"", 
            	
            	 -SUM(TT1.""InsTotal"")*-1 AS ""InsTotal"",
                 -SUM(TT1.""InsTotalFC"" - TT1.""PaidFC""-TT1.""WTSumFC""+TT1.""WTAppliedF"")*-1 AS ""OpenAmountFC"",
            	 -SUM(TT1.""InsTotalFC"")*-1 AS ""InsTotalFC""
            	FROM ODPO TT0 
            	INNER JOIN DPO6 TT1 ON TT0.""DocEntry"" = TT1.""DocEntry""
                LEFT JOIN DPO5 TT2 ON (TT0.""DocEntry"" = TT2.""AbsEntry"" AND TT0.""ObjType"" = TT2.""ObjType"")
            	INNER JOIN OCRD T3 ON TT0.""CardCode"" = T3.""CardCode"" 
            	WHERE  TT0.""DocDate"" <= '" + dateE + @"' 
            	AND TT0.""CardCode"" = N'" + cardCode + @"'
            	AND (TT0.""DocStatus"" = 'O' 
            		OR (TT1.""Status"" = 'O' 
            			AND TT0.""CANCELED"" = 'N')) 
            	GROUP BY TT0.""DocEntry"",
            	 TT0.""Project"",
            	 TT0.""DocNum"",
                 TT0.""DocCur"",
            	 T3.""CardCode"",
            	 T3.""CardName"",
            	 TT0.""DocDate"",
            	 TT1.""DueDate"",
            	 TT0.""ObjType"",
            	 TT0.""Comments"",
                 TT1.""InstlmntID"",
                 TT2.""WTCode""
            	) T0 
            WHERE (T0.""OpenAmount"" <> '0' OR T0.""OpenAmountFC"" <> '0')
            ORDER BY 
            T0.""Project"",
            T0.""DueDate"",
            	 T0.""DocNum""";

            oRecordSet.DoQuery(query);

            oDataTable.Rows.Clear();

            try
            {
                int rowIndex = 0;
                int docEntry;
                int docNum;
                int installmentID;
                string docType;
                DateTime dueDate;
                decimal openAmount;
                decimal InsTotal;
                decimal totalPaymentLC;
                decimal totalPaymentFC;
                decimal WTSum;
                decimal rate;
                int overdueDays;
                //string whTaxCode = "";
                //string whTaxCode = oForm.DataSources.UserDataSources.Item("WHTax").ValueEx;
                //bool physicalEntityTax = (BusinessPartners.isWTLiable(cardCode) && CommonFunctions.getValue("OWHT", "U_BDOSPhisTx", "WTCode", whTaxCode).ToString() == "Y");

                while (!oRecordSet.EoF)
                {
                    docEntry = Convert.ToInt32(oRecordSet.Fields.Item("DocEntry").Value);
                    docNum = Convert.ToInt32(oRecordSet.Fields.Item("DocNum").Value);
                    installmentID = Convert.ToInt32(oRecordSet.Fields.Item("InstallmentID").Value);
                    docType = Convert.ToString(oRecordSet.Fields.Item("ObjType").Value);
                    dueDate = oRecordSet.Fields.Item("DueDate").Value;
                    totalPaymentLC = Convert.ToDecimal(oRecordSet.Fields.Item("OpenAmount").Value);
                    totalPaymentFC = Convert.ToDecimal(oRecordSet.Fields.Item("OpenAmountFC").Value);
                    openAmount = Convert.ToDecimal(oRecordSet.Fields.Item("OpenAmountFC").Value);
                    if (openAmount == 0)
                        openAmount = Convert.ToDecimal(oRecordSet.Fields.Item("OpenAmount").Value);
                    InsTotal = Convert.ToDecimal(oRecordSet.Fields.Item("InsTotalFC").Value);
                    if (InsTotal == 0)
                        InsTotal = Convert.ToDecimal(oRecordSet.Fields.Item("InsTotal").Value);
                    WTSum = Convert.ToDecimal(oRecordSet.Fields.Item("WTSumFC").Value);
                    if (WTSum == 0)
                        WTSum = Convert.ToDecimal(oRecordSet.Fields.Item("WTSum").Value);
                    overdueDays = Convert.ToInt32(oRecordSet.Fields.Item("OverdueDays").Value);
                    string docCur = Convert.ToString(oRecordSet.Fields.Item("DocCur").Value);
                    if (string.IsNullOrEmpty(docCur))
                        docCur = Program.MainCurrency;

                    if (docCur != Program.LocalCurrency)
                    {
                        rate = Convert.ToDecimal(oSBOBob.GetCurrencyRate(docCur, date).Fields.Item("CurrencyRate").Value);
                        totalPaymentLC = totalPaymentFC * rate;
                    }

                    oDataTable.Rows.Add();
                    oDataTable.SetValue("LineNum", rowIndex, rowIndex + 1);
                    oDataTable.SetValue("CheckBox", rowIndex, "N");
                    oDataTable.SetValue("DocEntry", rowIndex, docEntry);
                    oDataTable.SetValue("DocNum", rowIndex, docNum);
                    oDataTable.SetValue("InstallmentID", rowIndex, oRecordSet.Fields.Item("InstallmentID").Value);
                    oDataTable.SetValue("LineID", rowIndex, oRecordSet.Fields.Item("LineID").Value);
                    oDataTable.SetValue("DocType", rowIndex, docType);
                    oDataTable.SetValue("DocDate", rowIndex, oRecordSet.Fields.Item("DocDate").Value.ToString("yyyyMMdd") == "18991230" ? "" : oRecordSet.Fields.Item("DocDate").Value.ToString("yyyyMMdd"));
                    oDataTable.SetValue("DueDate", rowIndex, oRecordSet.Fields.Item("DueDate").Value.ToString("yyyyMMdd") == "18991230" ? "" : oRecordSet.Fields.Item("DueDate").Value.ToString("yyyyMMdd"));
                    oDataTable.SetValue("WTCode", rowIndex, oRecordSet.Fields.Item("WTCode").Value);
                    oDataTable.SetValue("Arrears", rowIndex, overdueDays >= 0 ? "*" : "");
                    oDataTable.SetValue("OverdueDays", rowIndex, overdueDays);
                    oDataTable.SetValue("Comments", rowIndex, oRecordSet.Fields.Item("Comments").Value);
                    oDataTable.SetValue("Total", rowIndex, Convert.ToDouble(InsTotal));
                    oDataTable.SetValue("WTSum", rowIndex, Convert.ToDouble(WTSum));
                    oDataTable.SetValue("BalanceDue", rowIndex, Convert.ToDouble(openAmount));
                    oDataTable.SetValue("TotalPaymentLC", rowIndex, Convert.ToDouble(totalPaymentLC));
                    oDataTable.SetValue("TotalPaymentFC", rowIndex, Convert.ToDouble(totalPaymentFC));
                    oDataTable.SetValue("Currency", rowIndex, docCur);
                    oDataTable.SetValue("Project", rowIndex, oRecordSet.Fields.Item("Project").Value);

                    if (CommonFunctions.IsDevelopment())
                    {
                        string bCode = oForm.DataSources.UserDataSources.Item("BDOSDefCfE").ValueEx.Trim();
                        string bName = UDO.GetUDOFieldValueByParam("UDO_F_BDOSBUCFW_D", "Code", bCode, "Name");
                        bName = bName == null ? "" : bName;

                        oDataTable.SetValue("BudgetCashFlowID", rowIndex, bCode);
                        oDataTable.SetValue("BudgetCashFlowName", rowIndex, bName);
                    }

                    oDataTable.SetValue("CFWId", rowIndex, oForm.DataSources.UserDataSources.Item("CashFlowI").ValueEx);
                    oDataTable.SetValue("Description", rowIndex, oForm.DataSources.UserDataSources.Item("Descrpt").ValueEx);

                    calculatePensionAmt(oForm, oDataTable, rowIndex + 1);

                    oRecordSet.MoveNext();
                    rowIndex++;
                }

                SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("InvoiceMTR").Specific));
                oForm.Freeze(true);
                oMatrix.Clear();
                oMatrix.LoadFromDataSource();
                oMatrix.AutoResizeColumns();

                setEditableSetting(oForm);

                oForm.DataSources.UserDataSources.Item("TtlPmntLCE").ValueEx = "0";
                oForm.DataSources.UserDataSources.Item("TtlPmntFCE").ValueEx = "0";
                oForm.DataSources.UserDataSources.Item("TtlPmntFCT").ValueEx = "";
                oForm.Items.Item("TtlPmntFCE").Visible = true;
                oForm.Items.Item("TtlPmntFCT").Visible = false;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                oForm.Update();
                oForm.Freeze(false);
                Marshal.FinalReleaseComObject(oRecordSet);
            }
        }

        private static void setEditableSetting(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("InvoiceMTR").Specific;
            oMatrix.FlushToDataSource();

            if (oMatrix.RowCount > 0)
            {
                SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("InvoiceMTR");

                for (int row = 1; row <= oDataTable.Rows.Count; row++)
                {
                    //int docEntry = oDataTable.GetValue("DocEntry", row - 1);
                    string currency = oDataTable.GetValue("Currency", row - 1);

                    //oMatrix.CommonSetting.SetCellEditable(row, 4, docEntry == 0);
                    //oMatrix.CommonSetting.SetCellEditable(row, 19, docEntry == 0);
                    oMatrix.CommonSetting.SetCellEditable(row, 19, currency != Program.LocalCurrency);
                }
            }
        }

        private static void updateRow(SAPbouiCOM.Form oForm, bool cfwIdChng, bool descrptionChng, int rowIndex = 0)
        {
            try
            {
                oForm.Freeze(true);

                string cfwId = oForm.DataSources.UserDataSources.Item("CashFlowI").ValueEx;
                string descrption = oForm.DataSources.UserDataSources.Item("Descrpt").ValueEx;

                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("InvoiceMTR").Specific;
                oMatrix.FlushToDataSource();

                int rowCount = rowIndex == 0 ? oMatrix.RowCount : rowIndex;
                int i = rowIndex == 0 ? 1 : rowIndex;

                SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("InvoiceMTR");
                for (; i <= rowCount; i++)
                {
                    if (cfwIdChng)
                        oDataTable.SetValue("CFWId", i - 1, cfwId);
                    if (descrptionChng)
                        oDataTable.SetValue("Description", i - 1, descrption);
                }

                oMatrix.LoadFromDataSource();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                oForm.Freeze(false);
                GC.Collect();
            }
        }

        private static void updateTotalPaymentRow(SAPbouiCOM.Form oForm, string baseColumn, int rowIndex = 0)
        {
            try
            {
                oForm.Freeze(true);
                //string whTaxCode = "";
                //string whTaxCode = oForm.DataSources.UserDataSources.Item("WHTax").ValueEx;
                //string cardCode = oForm.DataSources.UserDataSources.Item("BPCode").ValueEx;
                //bool physicalEntityTax = (BusinessPartners.isWTLiable(cardCode) && CommonFunctions.getValue("OWHT", "U_BDOSPhisTx", "WTCode", whTaxCode).ToString() == "Y");

                string docDateS = oForm.DataSources.UserDataSources.Item("DocPstDt").ValueEx;
                DateTime date = Convert.ToDateTime(DateTime.ParseExact(docDateS, "yyyyMMdd", CultureInfo.InvariantCulture));
                SAPbobsCOM.SBObob oSBOBob = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);

                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("InvoiceMTR").Specific;
                oMatrix.FlushToDataSource();

                int rowCount = rowIndex == 0 ? oMatrix.RowCount : rowIndex;
                int i = rowIndex == 0 ? 1 : rowIndex;

                SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("InvoiceMTR");
                for (; i <= rowCount; i++)
                {
                    string currency = oDataTable.GetValue("Currency", i - 1);
                    if (currency != Program.LocalCurrency)
                    {
                        decimal totalPaymentLC = Convert.ToDecimal(oDataTable.GetValue("TotalPaymentLC", i - 1), CultureInfo.InvariantCulture);
                        decimal totalPaymentFC = Convert.ToDecimal(oDataTable.GetValue("TotalPaymentFC", i - 1), CultureInfo.InvariantCulture);
                        string useBlaAgRt = oDataTable.GetValue("UseBlaAgRt", i - 1);
                        string blnktAgr = oDataTable.GetValue("BlnktAgr", i - 1);
                        decimal rateByBlnktAgr = 0;
                        string docCur;
                        if (useBlaAgRt == "Y")
                            rateByBlnktAgr = BlanketAgreement.GetBlAgremeentCurrencyRate(Convert.ToInt32(blnktAgr), out docCur, date);

                        decimal rate = useBlaAgRt == "Y" ? rateByBlnktAgr : Convert.ToDecimal(oSBOBob.GetCurrencyRate(currency, date).Fields.Item("CurrencyRate").Value, CultureInfo.InvariantCulture);

                        if (baseColumn == "TtlPmntFC")
                        {
                            totalPaymentLC = totalPaymentFC * rate;
                            oDataTable.SetValue("TotalPaymentLC", i - 1, Convert.ToDouble(totalPaymentLC, CultureInfo.InvariantCulture));
                        }
                        else if (baseColumn == "TtlPmntLC")
                        {
                            totalPaymentFC = totalPaymentLC / rate;
                            oDataTable.SetValue("TotalPaymentFC", i - 1, Convert.ToDouble(totalPaymentFC, CultureInfo.InvariantCulture));
                        }
                    }
                    calculatePensionAmt(oForm, oDataTable, i);
                }
                oMatrix.LoadFromDataSource();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                oForm.Freeze(false);
                GC.Collect();
            }
        }

        private static void showSelectedRowsTotalPayment(SAPbouiCOM.Form oForm)
        {
            try
            {
                oForm.Freeze(true);

                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("InvoiceMTR").Specific;
                oMatrix.FlushToDataSource();
                SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("InvoiceMTR");
                decimal totalPaymentLC = 0;
                decimal totalPaymentFC = 0;
                string currency = null;
                bool differentFC = false;
                for (int i = 0; i < oDataTable.Rows.Count; i++)
                {
                    if (oDataTable.GetValue("CheckBox", i) == "Y")
                    {
                        totalPaymentLC += Convert.ToDecimal(oDataTable.GetValue("TotalPaymentLC", i), CultureInfo.InvariantCulture);
                        totalPaymentFC += Convert.ToDecimal(oDataTable.GetValue("TotalPaymentFC", i), CultureInfo.InvariantCulture);
                        if (!differentFC && oDataTable.GetValue("Currency", i) != Program.LocalCurrency)
                        {
                            if (string.IsNullOrEmpty(currency))
                                currency = currency = oDataTable.GetValue("Currency", i);
                            else if (currency != oDataTable.GetValue("Currency", i))
                            {
                                totalPaymentFC = 0;
                                differentFC = true;
                            }
                        }
                    }
                }
                oForm.DataSources.UserDataSources.Item("TtlPmntLCE").ValueEx = FormsB1.ConvertDecimalToString(totalPaymentLC);
                oForm.DataSources.UserDataSources.Item("TtlPmntFCE").ValueEx = FormsB1.ConvertDecimalToString(totalPaymentFC);
                oForm.DataSources.UserDataSources.Item("TtlPmntFCT").ValueEx = differentFC ? "*" : "";
                oForm.Items.Item("TtlPmntFCE").Visible = !differentFC;
                oForm.Items.Item("TtlPmntFCT").Visible = differentFC;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                oForm.Freeze(false);
                GC.Collect();
            }
        }

        private static void calculatePensionAmt(SAPbouiCOM.Form oForm, SAPbouiCOM.DataTable oDataTable = null, int rowIndex = 0)
        {
            try
            {
                oForm.Freeze(true);
                bool fromWHTax = false;

                SAPbouiCOM.Matrix oMatrix = oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("InvoiceMTR").Specific;
                if (oDataTable == null)
                {
                    fromWHTax = true;
                    oMatrix.FlushToDataSource();
                    oDataTable = oForm.DataSources.DataTables.Item("InvoiceMTR");
                }

                int rowCount = rowIndex == 0 ? oDataTable.Rows.Count : rowIndex;
                int i = rowIndex == 0 ? 1 : rowIndex;

                string errorText;
                string cardCode = oForm.DataSources.UserDataSources.Item("BPCode").ValueEx;
                bool isWTLiable = BusinessPartners.isWTLiable(cardCode);
                Dictionary<string, decimal> physicalEntityPensionRates;

                for (; i <= rowCount; i++)
                {
                    decimal pensEmployer = 0; //დამსაქმებელი
                    decimal pensEmployed = 0; //დასაქმებული
                    decimal whTaxAmt = 0; //საშემოსავლო გადასახადი
                    decimal grossAmt = 0; //გროსი თანხა

                    string wtCode = oDataTable.GetValue("WTCode", i - 1);

                    if (!string.IsNullOrEmpty(wtCode))
                    {
                        bool physicalEntityTax = CommonFunctions.getValue("OWHT", "U_BDOSPhisTx", "WTCode", wtCode).ToString() == "Y";

                        if (isWTLiable)
                        {
                            DateTime date = oDataTable.GetValue("DocDate", i - 1);
                            physicalEntityPensionRates = WithholdingTax.GetPhysicalEntityPensionRates(date, wtCode, out errorText);

                            if (string.IsNullOrEmpty(errorText))
                            {
                                //string docType = oDataTable.GetValue("DocType", i - 1);
                                decimal totalPaymentLC = Convert.ToDecimal(oDataTable.GetValue("TotalPaymentLC", i - 1), CultureInfo.InvariantCulture);
                                decimal totalPaymentFC = Convert.ToDecimal(oDataTable.GetValue("TotalPaymentFC", i - 1), CultureInfo.InvariantCulture);

                                decimal wtRate = physicalEntityPensionRates["WTRate"] / 100;
                                decimal pensionWTaxRate = physicalEntityTax ? physicalEntityPensionRates["PensionWTaxRate"] / 100 : 0;
                                decimal pensionCoWTaxRate = physicalEntityTax ? physicalEntityPensionRates["PensionCoWTaxRate"] / 100 : 0;

                                grossAmt = totalPaymentLC / (1 - wtRate) / (1 - pensionWTaxRate);

                                var physicalEntityTaxesAmt = CommonFunctions.CalcPhysicalEntityTaxes((grossAmt, wtRate, pensionWTaxRate, pensionCoWTaxRate));
                                whTaxAmt = physicalEntityTaxesAmt.whTaxAmt;
                                pensEmployed = physicalEntityTaxesAmt.pensEmployedAmt;
                                pensEmployer = physicalEntityTaxesAmt.pensEmployerAmt;
                            }
                            else
                            {
                                Program.uiApp.StatusBar.SetSystemMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                            }
                        }
                    }
                    oDataTable.SetValue("WTSum", i - 1, Convert.ToDouble(whTaxAmt, CultureInfo.InvariantCulture));
                    oDataTable.SetValue("PensEmpld", i - 1, Convert.ToDouble(pensEmployed, CultureInfo.InvariantCulture));
                    oDataTable.SetValue("PensEmplr", i - 1, Convert.ToDouble(pensEmployer, CultureInfo.InvariantCulture));
                }
                if (fromWHTax)
                    oMatrix.LoadFromDataSource();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                oForm.Freeze(false);
                GC.Collect();
            }
        }

        public static void addMenus()
        {
            SAPbouiCOM.MenuItem menuItem;
            SAPbouiCOM.MenuItem fatherMenuItem;
            SAPbouiCOM.MenuCreationParams oCreationPackage;

            try
            {
                fatherMenuItem = Program.uiApp.Menus.Item("43538");

                // Add a pop-up menu item
                oCreationPackage = (SAPbouiCOM.MenuCreationParams)Program.uiApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                oCreationPackage.Checked = false;
                oCreationPackage.Enabled = true;
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "BDOSSOPWizzForm";
                oCreationPackage.String = BDOSResources.getTranslate("OutgoingPaymentWizard");
                oCreationPackage.Position = fatherMenuItem.SubMenus.Count - 1;

                menuItem = fatherMenuItem.SubMenus.AddEx(oCreationPackage);
            }
            catch
            {

            }
        }
    }
}
