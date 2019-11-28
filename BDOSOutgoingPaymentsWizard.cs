using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;
using System.Data;

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

            bool multiSelection;

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
                    multiSelection = false;
                    string objectTypeCardCode = "2"; //SAPbouiCOM.BoLinkedObject.lf_BusinessPartner, Business Partner object 
                    string uniqueID_lf_BusinessPartnerCFL = "BusinessPartner_CFL";
                    FormsB1.addChooseFromList(oForm, multiSelection, objectTypeCardCode, uniqueID_lf_BusinessPartnerCFL);

                    //პირობის დადება ბიზნესპარტნიორის არჩევის სიაზე
                    SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(uniqueID_lf_BusinessPartnerCFL);
                    SAPbouiCOM.Conditions oCons = oCFL.GetConditions();
                    SAPbouiCOM.Condition oCon = oCons.Add();
                    oCon.Alias = "CardType";
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCon.CondVal = "S"; //მომწოდებელი
                    oCFL.SetConditions(oCons);

                    string uniqueID_lf_HBAccountCFL = "HouseBankAcc_CFL";
                    string objectTypeHB = "231";
                    FormsB1.addChooseFromList(oForm, multiSelection, objectTypeHB, uniqueID_lf_HBAccountCFL);

                    string uniqueID_lf_GLAccCFL = "GLAcc_CFL";
                    string objectTypeGLAcc = "1";
                    FormsB1.addChooseFromList(oForm, multiSelection, objectTypeGLAcc, uniqueID_lf_GLAccCFL);

                    //პირობის დადება ბიზნესპარტნიორის არჩევის სიაზე
                    oCFL = oForm.ChooseFromLists.Item(uniqueID_lf_GLAccCFL);
                    oCons = oCFL.GetConditions();
                    oCon = oCons.Add();
                    oCon.Alias = "Postable";
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCon.CondVal = "Y"; //მომწოდებელი
                    oCFL.SetConditions(oCons);

                    string uniqueID_lf_CTAccCFL = "CTAcc_CFL";
                    objectTypeGLAcc = "1";
                    FormsB1.addChooseFromList(oForm, multiSelection, objectTypeGLAcc, uniqueID_lf_CTAccCFL);

                    //პირობის დადება ბიზნესპარტნიორის არჩევის სიაზე
                    oCFL = oForm.ChooseFromLists.Item(uniqueID_lf_CTAccCFL);
                    oCons = oCFL.GetConditions();
                    oCon = oCons.Add();
                    oCon.Alias = "LocManTran";
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCon.CondVal = "Y"; //მომწოდებელი
                    oCFL.SetConditions(oCons);

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
                    formItems.Add("ChooseFromListUID", uniqueID_lf_BusinessPartnerCFL);
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
                    formItems.Add("LinkedObjectType", objectTypeCardCode);

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
                    formItems.Add("ChooseFromListUID", uniqueID_lf_GLAccCFL);
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
                    formItems.Add("LinkedObjectType", objectTypeGLAcc);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

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
                    formItems.Add("ChooseFromListUID", uniqueID_lf_CTAccCFL);
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
                    formItems.Add("LinkedObjectType", objectTypeGLAcc);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    top = top + height + 1;

                    formItems = new Dictionary<string, object>();
                    itemName = "HBAccS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("BankAcc"));

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "HBAcc"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("TableName", "");
                    formItems.Add("Length", 50);
                    formItems.Add("Size", 50);
                    formItems.Add("Alias", "HBAcc");
                    formItems.Add("Bound", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("Left", left_e);
                    formItems.Add("Width", width_e);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("ChooseFromListUID", uniqueID_lf_HBAccountCFL);
                    formItems.Add("ChooseFromListAlias", "Account");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "WHtaxS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s2);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("WithholdingTax"));

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "WHTax"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("Length", 1);
                    formItems.Add("Size", 20);
                    formItems.Add("Left", left_e2);
                    formItems.Add("Width", width_e);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);

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

                        multiSelection = false;
                        string objectType = "UDO_F_BDOSBUCFW_D";
                        string uniqueID_lf_Budg_CFL_head = "Budg_CFLHD";
                        FormsB1.addChooseFromList(oForm, multiSelection, objectType, uniqueID_lf_Budg_CFL_head);

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
                        formItems.Add("ChooseFromListUID", uniqueID_lf_Budg_CFL_head);
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
                    formItems.Add("Height", 19);
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
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("Image", "HANA_CHECKBOX_UH");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "AddRow";
                    formItems.Add("Caption", BDOSResources.getTranslate("AddNewRow"));
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left_s + (20 + 1) * 2);
                    formItems.Add("Width", 150);
                    formItems.Add("Top", top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "CreatDocmt";
                    formItems.Add("Caption", BDOSResources.getTranslate("CreateDocuments"));
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left_s + 155 + (20 + 1) * 2);
                    formItems.Add("Width", 150);
                    formItems.Add("Top", top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    top = top + height + 5;

                    formItems = new Dictionary<string, object>();
                    itemName = "InvoiceMTR"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", 600);
                    formItems.Add("Top", top);
                    formItems.Add("Height", 550);
                    formItems.Add("UID", itemName);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("AffectsFormMode", false);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }
                    SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("InvoiceMTR").Specific;

                    oColumns = oMatrix.Columns;

                    SAPbouiCOM.LinkedButton oLink;
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
                    oDataTable.Columns.Add("Arrears", SAPbouiCOM.BoFieldsType.ft_Text, 1); //* აჩვენებს, რომ Due Date ნაკლებია ან ტოლი გადახდის თარიღზე
                    oDataTable.Columns.Add("OverdueDays", SAPbouiCOM.BoFieldsType.ft_Integer, 50); //გადახდის თარიღსა და Due Date-ს შორის სხვაობა
                    oDataTable.Columns.Add("Total", SAPbouiCOM.BoFieldsType.ft_Sum); //თანხა
                    oDataTable.Columns.Add("WTSum", SAPbouiCOM.BoFieldsType.ft_Sum); //თანხა                    
                    oDataTable.Columns.Add("PensSum", SAPbouiCOM.BoFieldsType.ft_Sum); //თანხა                    
                    oDataTable.Columns.Add("BalanceDue", SAPbouiCOM.BoFieldsType.ft_Sum); //დოკუმენტის დაურეკონსილირებელი თანხა - ვალის ნაშთი
                    oDataTable.Columns.Add("TotalPayment", SAPbouiCOM.BoFieldsType.ft_Sum); //Default - Balance Due
                    oDataTable.Columns.Add("TotalPaymentNet", SAPbouiCOM.BoFieldsType.ft_Sum); //Default - Balance Due
                    oDataTable.Columns.Add("Currency", SAPbouiCOM.BoFieldsType.ft_Text, 50); //დოკუმენტის ვალუტა
                    oDataTable.Columns.Add("TotalPaymentLocal", SAPbouiCOM.BoFieldsType.ft_Sum); //Default - Balance Due
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
                    //FormsB1.addChooseFromList(oForm, false, "242", "CashFlowLineItem_CFL");

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
                        else if (columnName == "Arrears")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = "*";
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind("InvoiceMTR", columnName);
                        }
                        else if (columnName == "TotalPayment")
                        {
                            oColumn = oColumns.Add("TotalPymnt", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate(columnName);
                            oColumn.DataBind.Bind("InvoiceMTR", columnName);
                        }
                        else if (columnName == "TotalPaymentNet")
                        {
                            oColumn = oColumns.Add("TtlPymntNt", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate(columnName);
                            oColumn.DataBind.Bind("InvoiceMTR", columnName);
                        }
                        else if (columnName == "WTSum")
                        {
                            oColumn = oColumns.Add("WTSum", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("WTaxAmount");
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind("InvoiceMTR", columnName);
                        }
                        else if (columnName == "PensSum")
                        {
                            oColumn = oColumns.Add("PensSum", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("PensionAmount");
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind("InvoiceMTR", columnName);
                        }
                        else if (columnName == "TotalPaymentLocal")
                        {
                            oColumn = oColumns.Add("TotalPmntL", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate(columnName);
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
                }
                resizeItems(oForm);
                oForm.Visible = true;
                oForm.Select();
            }
        }

        public static void resizeItems(SAPbouiCOM.Form oForm)
        {
            try
            {
                SAPbouiCOM.Item oMatrixItem = oForm.Items.Item("InvoiceMTR");

                oMatrixItem.Height = oForm.Height - 220;
                oMatrixItem.Width = oForm.Width - 20;
            }
            catch
            {
            }
        }

        private static int createPaymentDocument(SAPbouiCOM.Form oForm, DataRow headerLine, DataTable AccountPaymentsLines)
        {
            string errorText = null;

            SAPbouiCOM.EditText oEditTextDocDate = (SAPbouiCOM.EditText)oForm.Items.Item("DocPstDt").Specific;
            string DocDateS = oEditTextDocDate.Value;
            DateTime DocDate = Convert.ToDateTime(DateTime.ParseExact(DocDateS, "yyyyMMdd", CultureInfo.InvariantCulture));

            SAPbobsCOM.SBObob vObj;
            vObj = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);

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


            string LocalCurrency = CurrencyB1.getMainCurrency(out errorText);
            string BankAccount = headerLine["BankAccount"].ToString();
            string TransferAccount = headerLine["TransferAccount"].ToString();
            string ControlAccount = headerLine["ControlAccount"].ToString();

            string DocCurrency = headerLine["Currency"].ToString();
            string PayblCur = headerLine["PayblCur"].ToString();
            string remarks = headerLine["remarks"].ToString();

            string ChrgDtls = headerLine["ChrgDtls"].ToString();
            string DispType = headerLine["DispType"].ToString();

            double TransferSumFC = Convert.ToDouble(headerLine["PayblAmtFC"]);
            double TransferSum = Convert.ToDouble(headerLine["PayblAmt"]);
            string CardCode = headerLine["CardCode"].ToString();
            string Project = headerLine["Project"].ToString();
            string WTCode = headerLine["WTCode"].ToString();

            double WtAmount = Convert.ToDouble(headerLine["WtAmount"]);
            double PensioAmount = Convert.ToDouble(headerLine["PensionAmount"]);

            SAPbobsCOM.Payments OutPay = null;

            OutPay = (SAPbobsCOM.Payments)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments);
            OutPay.DocObjectCode = SAPbobsCOM.BoPaymentsObjectType.bopot_OutgoingPayments;

            OutPay.DocDate = DocDate;
            OutPay.ProjectCode = Project;



            if (CommonFunctions.IsDevelopment())
            {
                string BudgetCashFlowID = headerLine["BudgetCashFlowID"].ToString();
                string BudgetCashFlowName = headerLine["BudgetCashFlowName"].ToString();

                if (string.IsNullOrEmpty(BudgetCashFlowID) == false)
                {
                    OutPay.UserFields.Fields.Item("U_BDOSBdgCf").Value = BudgetCashFlowID;
                    OutPay.UserFields.Fields.Item("U_BDOSBdgCfN").Value = BudgetCashFlowName;
                }
            }

            try
            {
                OutPay.UserFields.Fields.Item("U_status").Value = "readyToLoad";
                OutPay.UserFields.Fields.Item("U_chrgDtls").Value = ChrgDtls;
                OutPay.UserFields.Fields.Item("U_dsptchType").Value = DispType;
                OutPay.UserFields.Fields.Item("U_descrpt").Value = headerLine["Descrpt"].ToString();
                OutPay.UserFields.Fields.Item("U_addDescrpt").Value = headerLine["AddDescrpt"].ToString();
            }
            catch
            { }

            OutPay.CardCode = CardCode;
            OutPay.DocTypte = SAPbobsCOM.BoRcptTypes.rSupplier;

            if (BankAccount == "")
            {
                OutPay.IsPayToBank = SAPbobsCOM.BoYesNoEnum.tNO;
            }
            else
            {
                OutPay.IsPayToBank = SAPbobsCOM.BoYesNoEnum.tYES;
            }

            OutPay.TransferAccount = TransferAccount;
            OutPay.ControlAccount = ControlAccount;

            OutPay.Remarks = remarks;
            double DocRate = 0;
            if (DocCurrency == LocalCurrency)
            {
                OutPay.DocRate = 0;
            }
            else
            {
                DocRate = vObj.GetCurrencyRate(DocCurrency, DocDate).Fields.Item("CurrencyRate").Value;
                OutPay.DocRate = DocRate;
            }

            OutPay.DocCurrency = PayblCur;

            if (DocCurrency == PayblCur)
            {
                OutPay.LocalCurrency = SAPbobsCOM.BoYesNoEnum.tNO;
            }
            else
            {
                OutPay.LocalCurrency = SAPbobsCOM.BoYesNoEnum.tYES;
            }

            OutPay.TransferSum = Convert.ToDouble(TransferSum);
            if (WTCode != "")
            {
                OutPay.WTCode = WTCode;
                OutPay.WtBaseSum = TransferSum;
                OutPay.WTAmount = WtAmount + PensioAmount;
            }

            OutPay.UserFields.Fields.Item("U_BDOSWhtAmt").Value = WtAmount;
            OutPay.UserFields.Fields.Item("U_BDOSPnPhAm").Value = PensioAmount;
            OutPay.UserFields.Fields.Item("U_BDOSPnCoAm").Value = PensioAmount;


            decimal OnAccount = 0;
            //ცხრილური ნაწილი
            DataRow AccountPaymentsLine;
            for (int i = 0; i < AccountPaymentsLines.Rows.Count; i++)
            {
                AccountPaymentsLine = AccountPaymentsLines.Rows[i];


                if (AccountPaymentsLine["DocEntry"].ToString() != "0")
                {

                    SAPbobsCOM.BoRcptInvTypes InvType;
                    int InvTypeInt = Convert.ToInt32(AccountPaymentsLine["InvType"]);

                    if (InvTypeInt == 18)
                    {
                        InvType = SAPbobsCOM.BoRcptInvTypes.it_PurchaseInvoice;
                    }
                    else if (InvTypeInt == 204)
                    {
                        InvType = SAPbobsCOM.BoRcptInvTypes.it_PurchaseDownPayment;
                    }
                    else
                    {
                        InvType = SAPbobsCOM.BoRcptInvTypes.it_APCorrectionInvoice;
                    }


                    OutPay.Invoices.DocEntry = Convert.ToInt32(AccountPaymentsLine["DocEntry"]);
                    OutPay.Invoices.InvoiceType = InvType;
                    OutPay.Invoices.SumApplied = (OutPay.DocRate == 0 ? 1 : OutPay.DocRate) * Convert.ToDouble(AccountPaymentsLine["SumApplied"]);
                    OutPay.Invoices.AppliedFC = Convert.ToDouble(AccountPaymentsLine["SumApplied"]);

                    OutPay.Invoices.InstallmentId = Convert.ToInt32(AccountPaymentsLine["InstallmentId"]);

                    DataRow DTSourceRowVPM2 = DTSourceVPM2.Rows.Add();
                    DTSourceRowVPM2["DocEntry"] = Convert.ToInt32(AccountPaymentsLine["DocEntry"]);
                    DTSourceRowVPM2["InvType"] = InvTypeInt;
                    DTSourceRowVPM2["SumApplied"] = OutPay.Invoices.SumApplied;
                    DTSourceRowVPM2["AppliedFC"] = OutPay.Invoices.AppliedFC;
                    OutPay.Invoices.Add();
                }
                else
                {
                    OnAccount = OnAccount + Convert.ToDecimal(AccountPaymentsLine["SumApplied"], CultureInfo.InvariantCulture);
                }
            }

            if (GetAccountCashFlowRelevant(TransferAccount))
            {
                OutPay.PrimaryFormItems.CashFlowLineItemID = Convert.ToInt32(headerLine["CashFlowID"]);
            }
            OutPay.PrimaryFormItems.AmountFC = (OutPay.DocRate == 0 ? 1 : OutPay.DocRate) * TransferSumFC;

            if (DocCurrency == LocalCurrency)
            {
                OutPay.PrimaryFormItems.AmountLC = TransferSum;
            }

            OutPay.PrimaryFormItems.PaymentMeans = SAPbobsCOM.PaymentMeansTypeEnum.pmtBankTransfer;
            OutPay.PrimaryFormItems.Add();


            DataRow DTSourceRow = DTSource.Rows.Add();

            SAPbobsCOM.BusinessPartners oBP;
            oBP = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);

            if (oBP.GetByKey(CardCode))
            {
                WTCode = oBP.WTCode;
            }
            DTSourceRow["WtCode"] = WTCode;
            DTSourceRow["WTLiable"] = "Y";
            DTSourceRow["CardCode"] = CardCode;
            DTSourceRow["PrjCode"] = Project;
            DTSourceRow["U_liablePrTx"] = "N";
            DTSourceRow["U_prBase"] = "";


            DTSourceRow["NoDocSum"] = OnAccount;
            DTSourceRow["U_BDOSWhtAmt"] = WtAmount;
            DTSourceRow["U_BDOSPnPhAm"] = PensioAmount;
            DTSourceRow["U_BDOSPnCoAm"] = PensioAmount;


            CommonFunctions.StartTransaction();

            int resultCode = OutPay.Add();

            if (resultCode != 0)
            {
                string errorMessage = "";
                Program.oCompany.GetLastError(out resultCode, out errorMessage);
                if (Program.oCompany.InTransaction)
                {
                    CommonFunctions.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                }

                errorMessage = "";
                Program.oCompany.GetLastError(out resultCode, out errorMessage);
                Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("DocumentNotCreated") + ". " + BDOSResources.getTranslate("ReasonIs") + ": " + errorMessage, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return 0;
            }
            else
            {
                string docEntryS = "";
                Program.oCompany.GetNewObjectCode(out docEntryS);
                DataTable reLines = null;

                DataTable JrnLinesDT = OutgoingPayment.createAdditionalEntries(null, null, DTSource, DTSourceVPM2, OutPay.DocCurrency, out reLines, Convert.ToDecimal(OutPay.DocRate));
                OutgoingPayment.JrnEntry(docEntryS, docEntryS, OutPay.DocDate, JrnLinesDT, reLines, out errorText);

                if (errorText != null)
                {
                    if (Program.oCompany.InTransaction)
                    {
                        CommonFunctions.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                    }
                    Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("DocumentNotCreated") + ". " + BDOSResources.getTranslate("ReasonIs") + ": " + errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
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

        private static bool GetAccountCashFlowRelevant(string GLAccount)
        {
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string query = @"SELECT
	                        ""CfwRlvnt""
                            FROM ""OACT"" 
                            where ""AcctCode"" = '" + GLAccount + "'";


            oRecordSet.DoQuery(query);

            while (!oRecordSet.EoF)
            {
                return (oRecordSet.Fields.Item("CfwRlvnt").Value == "Y");
            }

            return false;
        }

        public static void fillBdgFlowItems(SAPbouiCOM.Form oForm)
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
            oForm.Update();
            oForm.Freeze(false);
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
                            }
                            else if (pVal.ItemUID == "fillBdgFl")
                            {
                                fillBdgFlowItems(oForm);
                            }
                            else if (pVal.ItemUID == "AddRow")
                            {
                                AddRow(oForm);
                            }
                            else if (pVal.ItemUID == "CreatDocmt")
                            {
                                createPaymentDocuments(oForm);
                            }
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
                        }
                    }

                    else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_VALIDATE && !pVal.InnerEvent)
                    {
                        if (pVal.ItemChanged)
                        {
                            if (pVal.ItemUID == "InvoiceMTR")
                            {
                                if ((pVal.ColUID == "TtlPymntNt" || pVal.ColUID == "TotalPymnt") && !pVal.BeforeAction)
                                {
                                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("InvoiceMTR").Specific;
                                    string RowDocEntry = oMatrix.GetCellSpecific("DocEntry", pVal.Row).Value;
                                    if (RowDocEntry == "0")
                                    {
                                        fillGrossAmount(oForm, pVal.ColUID, pVal.Row);
                                    }
                                }
                                else if (pVal.ColUID == "TotalPymnt" && pVal.BeforeAction)
                                {
                                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("InvoiceMTR").Specific;
                                    string RowDocEntry = oMatrix.GetCellSpecific("DocEntry", pVal.Row).Value;
                                    if (RowDocEntry != "0")
                                    {
                                        checkDueAmount(oForm, pVal.Row, out BubbleEvent);
                                    }
                                }
                            }
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

        public static void checkDueAmount(SAPbouiCOM.Form oForm, int row, out bool BubbleEvent)
        {
            BubbleEvent = true;

            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("InvoiceMTR").Specific;
            SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("InvoiceMTR");

            row = row - 1;
            decimal TotalPymnt = Convert.ToDecimal(oMatrix.GetCellSpecific("TotalPymnt", row + 1).Value, CultureInfo.InvariantCulture);
            decimal BalanceDue = Convert.ToDecimal(oMatrix.GetCellSpecific("BalanceDue", row + 1).Value, CultureInfo.InvariantCulture);
            if (BalanceDue < TotalPymnt)
            {
                TotalPymnt = Convert.ToDecimal(oDataTable.GetValue("TotalPayment", row), CultureInfo.InvariantCulture);
                oMatrix.GetCellSpecific("TotalPymnt", row + 1).Value = TotalPymnt;
            }
            else
            {
                oDataTable.SetValue("TotalPayment", row, Convert.ToDouble(TotalPymnt, CultureInfo.InvariantCulture));

                oForm.Freeze(true);
                oMatrix.Clear();
                oMatrix.LoadFromDataSource();
                oMatrix.AutoResizeColumns();
                oForm.Freeze(false);
            }
        }

        public static void fillGrossAmount(SAPbouiCOM.Form oForm, string Column, int row)
        {
            string errorText;

            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("InvoiceMTR").Specific;
            SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("InvoiceMTR");

            row = row - 1;
            decimal TtlPymntNt = Convert.ToDecimal(oMatrix.GetCellSpecific("TtlPymntNt", row + 1).Value, CultureInfo.InvariantCulture);
            decimal TotalPymnt = Convert.ToDecimal(oMatrix.GetCellSpecific("TotalPymnt", row + 1).Value, CultureInfo.InvariantCulture);

            string WHTaxCode = oForm.Items.Item("WHTax").Specific.Value;
            DataTable WTaxDefinitons = WithholdingTax.getWtaxCodeDefinitionByDate(DateTime.Now);
            string filter;
            DataRow[] oWHTaxCode;
            decimal pensionRate = 0;

            SAPbobsCOM.WithholdingTaxCodes oWhTax;
            oWhTax = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oWithholdingTaxCodes);
            if (oWhTax.GetByKey(WHTaxCode))
            {
                if (oWhTax.UserFields.Fields.Item("U_BDOSPhisTx").Value == "Y")
                {
                    string pensionCoWTCode = CommonFunctions.getOADM("U_BDOSPnCoP").ToString();
                    filter = "WTCode = '" + pensionCoWTCode + "'";
                    oWHTaxCode = WTaxDefinitons.Select(filter);
                    pensionRate = 0;
                    if (oWHTaxCode.Count() > 0)
                    {
                        pensionRate = Convert.ToDecimal(oWHTaxCode[0]["Rate"]);
                    }
                }
            }

            decimal WTRate = 0;
            filter = "WTCode = '" + WHTaxCode + "'";
            oWHTaxCode = WTaxDefinitons.Select(filter);
            if (oWHTaxCode.Count() > 0)
            {
                WTRate = Convert.ToDecimal(oWHTaxCode[0]["Rate"]);
            }

            decimal PensSum;
            decimal WTSum;

            if (Column == "TtlPymntNt")
            {
                TotalPymnt = TtlPymntNt / (1 - WTRate / 100) / (1 - pensionRate / 100);

                PensSum = TotalPymnt * pensionRate / 100;
                WTSum = (TotalPymnt - PensSum) * WTRate / 100;
            }
            else
            {
                PensSum = TotalPymnt * pensionRate / 100;
                WTSum = (TotalPymnt - PensSum) * WTRate / 100;
                TtlPymntNt = TotalPymnt - PensSum - WTSum;
            }

            oDataTable.SetValue("TotalPaymentNet", row, Convert.ToDouble(TtlPymntNt, CultureInfo.InvariantCulture));
            oDataTable.SetValue("PensSum", row, Convert.ToDouble(PensSum, CultureInfo.InvariantCulture));
            oDataTable.SetValue("WTSum", row, Convert.ToDouble(WTSum, CultureInfo.InvariantCulture));
            oDataTable.SetValue("TotalPayment", row, Convert.ToDouble(TotalPymnt, CultureInfo.InvariantCulture));

            oForm.Freeze(true);
            oMatrix.Clear();
            oMatrix.LoadFromDataSource();
            oMatrix.AutoResizeColumns();
            oForm.Freeze(false);
        }

        public static void SetInvDocsMatrixRowBackColor(SAPbouiCOM.Form oForm, int row)
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

        public static void checkUncheckMTR(SAPbouiCOM.Form oForm, string checkOperation)
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

        public static void matrixColumnSetLinkedObjectTypeInvoicesMTR(SAPbouiCOM.Form oForm, SAPbouiCOM.ItemEvent pVal)
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
            {
                string errorText;
                Currency = CurrencyB1.getMainCurrency(out errorText);

            }

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
            oDataTable.SetValue("TotalPayment", rowIndex, 0);
            oDataTable.SetValue("Currency", rowIndex, Currency);
            oDataTable.SetValue("TotalPaymentLocal", rowIndex, 0);
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

        private static void createPaymentDocuments(SAPbouiCOM.Form oForm)
        {
            int answer = Program.uiApp.MessageBox(BDOSResources.getTranslate("CreatePaymentDocuments") + "?", 1, BDOSResources.getTranslate("Yes"), BDOSResources.getTranslate("No"), "");

            if (answer == 2)
            {
                return;
            }

            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("InvoiceMTR").Specific;

            SAPbouiCOM.EditText oEditTextDocDate = (SAPbouiCOM.EditText)oForm.Items.Item("DocPstDt").Specific;
            string DocDateS = oEditTextDocDate.Value;
            DateTime DocDate = Convert.ToDateTime(DateTime.ParseExact(DocDateS, "yyyyMMdd", CultureInfo.InvariantCulture));

            int DocEntry;
            int totalSuccesfull = 0;
            int totalUnsuccesfull = 0;

            int CashFlowID = Convert.ToInt32(oForm.Items.Item("CashFlowI").Specific.Value);

            string prevCurrency = null;
            string prevProject = null;
            string prevDocIsEmpty = null;

            string prevBudgetCashFlowID = null;

            double PayblAmtFCTotal = 0;
            double PayblAmtTotal = 0;
            double WtAmountTotal = 0;
            double PensionAmountTotal = 0;

            string BankAccount = oForm.Items.Item("HBAcc").Specific.Value;
            string Descrpt = oForm.Items.Item("Descrpt").Specific.Value;
            if (BankAccount != "" && Descrpt == "")
            {
                Program.uiApp.StatusBar.SetSystemMessage("DescriptionIsMandatory", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return;
            }
            string GLAccount = oForm.Items.Item("GLAcc").Specific.Value;

            SAPbobsCOM.ChartOfAccounts oChartOfAccounts = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oChartOfAccounts);
            oChartOfAccounts.GetByKey(GLAccount);

            string PayblCur = oChartOfAccounts.AcctCurrency;
            string errorText;
            if (PayblCur == "##")
            {
                PayblCur = CurrencyB1.getMainCurrency(out errorText);
            }

            //WT
            string WHTaxCode = oForm.Items.Item("WHTax").Specific.Value;
            DataTable WTaxDefinitons = WithholdingTax.getWtaxCodeDefinitionByDate(DateTime.Now);
            string filter;
            DataRow[] oWHTaxCode;
            double pensionRate = 0;

            SAPbobsCOM.WithholdingTaxCodes oWhTax;
            oWhTax = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oWithholdingTaxCodes);
            if (oWhTax.GetByKey(WHTaxCode) == true)
            {
                if (oWhTax.UserFields.Fields.Item("U_BDOSPhisTx").Value == "Y")
                {
                    string pensionCoWTCode = CommonFunctions.getOADM("U_BDOSPnCoP").ToString();
                    filter = "WTCode = '" + pensionCoWTCode + "'";
                    oWHTaxCode = WTaxDefinitons.Select(filter);
                    pensionRate = 0;
                    if (oWHTaxCode.Count() > 0)
                    {
                        pensionRate = Convert.ToDouble(oWHTaxCode[0]["Rate"], CultureInfo.InvariantCulture);
                    }
                }
            }

            double WTRate = 0;
            filter = "WTCode = '" + WHTaxCode + "'";
            oWHTaxCode = WTaxDefinitons.Select(filter);
            if (oWHTaxCode.Count() > 0)
            {
                WTRate = Convert.ToDouble(oWHTaxCode[0]["Rate"], CultureInfo.InvariantCulture);
            }

            string ControlAccount = oForm.Items.Item("CTAcc").Specific.Value;
            string DispType = oForm.Items.Item("DispType").Specific.Value.Trim();
            string ChrgDtls = oForm.Items.Item("ChrgDtls").Specific.Value.Trim();

            DataTable AccountHeader = new DataTable();
            DataRow headerLine = AccountHeader.Rows.Add();

            AccountHeader.Columns.Add("CardCode");
            AccountHeader.Columns.Add("Currency");
            AccountHeader.Columns.Add("PayblCur");

            DataColumn colDecimal = new DataColumn("PayblCRt");
            colDecimal.DataType = System.Type.GetType("System.Decimal");
            AccountHeader.Columns.Add(colDecimal);

            colDecimal = new DataColumn("PayblAmt");
            colDecimal.DataType = System.Type.GetType("System.Decimal");
            AccountHeader.Columns.Add(colDecimal);

            colDecimal = new DataColumn("PayblAmtFC");
            colDecimal.DataType = System.Type.GetType("System.Decimal");
            AccountHeader.Columns.Add(colDecimal);

            colDecimal = new DataColumn("PensionAmount");
            colDecimal.DataType = System.Type.GetType("System.Decimal");
            AccountHeader.Columns.Add(colDecimal);

            colDecimal = new DataColumn("WtAmount");
            colDecimal.DataType = System.Type.GetType("System.Decimal");
            AccountHeader.Columns.Add(colDecimal);

            AccountHeader.Columns.Add("BankAccount");
            AccountHeader.Columns.Add("TransferAccount");
            AccountHeader.Columns.Add("ControlAccount");
            AccountHeader.Columns.Add("accrualDate");
            AccountHeader.Columns.Add("CashFlowID");
            AccountHeader.Columns.Add("remarks");
            AccountHeader.Columns.Add("ChrgDtls");
            AccountHeader.Columns.Add("DispType");
            AccountHeader.Columns.Add("Descrpt");
            AccountHeader.Columns.Add("AddDescrpt");
            AccountHeader.Columns.Add("Project");
            AccountHeader.Columns.Add("WTCode");
            AccountHeader.Columns.Add("BudgetCashFlowID");
            AccountHeader.Columns.Add("BudgetCashFlowName");

            DataTable AccountPaymentsLines = new DataTable();

            AccountPaymentsLines.Columns.Add("InvType");
            AccountPaymentsLines.Columns.Add("DocEntry");
            AccountPaymentsLines.Columns.Add("DocNum");
            AccountPaymentsLines.Columns.Add("InstallmentId");

            colDecimal = new DataColumn("SumApplied");
            colDecimal.DataType = System.Type.GetType("System.Decimal");
            AccountPaymentsLines.Columns.Add(colDecimal);

            string WTCode = oForm.Items.Item("WHTax").Specific.Value;

            for (int row = 1; row <= oMatrix.RowCount; row++)
            {
                SAPbouiCOM.CheckBox Edtfield = oMatrix.Columns.Item("CheckBox").Cells.Item(row).Specific;
                bool checkedLine = (Edtfield.Checked);

                if (checkedLine)
                {
                    string Currency = oMatrix.Columns.Item("Currency").Cells.Item(row).Specific.Value;

                    string BudgetCashFlowID = null;
                    if (CommonFunctions.IsDevelopment())
                    {
                        BudgetCashFlowID = oMatrix.Columns.Item("BCFWId").Cells.Item(row).Specific.Value;
                    }

                    //double DocRate = oMatrix.Columns.Item("PayblCRt").Cells.Item(row).Specific.Value;
                    double PayblAmt = Convert.ToDouble(oMatrix.Columns.Item("TotalPymnt").Cells.Item(row).Specific.Value, NumberFormatInfo.InvariantInfo);
                    double PayblAmtFC = Convert.ToDouble(oMatrix.Columns.Item("TotalPymnt").Cells.Item(row).Specific.Value, NumberFormatInfo.InvariantInfo);
                    double WtAmount = Convert.ToDouble(oMatrix.Columns.Item("WTSum").Cells.Item(row).Specific.Value, NumberFormatInfo.InvariantInfo);
                    double PensionAmount = Convert.ToDouble(oMatrix.Columns.Item("PensSum").Cells.Item(row).Specific.Value, NumberFormatInfo.InvariantInfo);

                    string Project = oMatrix.Columns.Item("Project").Cells.Item(row).Specific.Value;
                    string InvType = oMatrix.Columns.Item("DocType").Cells.Item(row).Specific.Value;
                    string InvDocEntry = oMatrix.Columns.Item("DocEntry").Cells.Item(row).Specific.Value;
                    string DocIsEmpty = (oMatrix.Columns.Item("DocEntry").Cells.Item(row).Specific.Value == "0").ToString();
                    string InvDocNum = oMatrix.Columns.Item("DocNum").Cells.Item(row).Specific.Value;
                    string InstallmentId = oMatrix.Columns.Item("InstlmntID").Cells.Item(row).Specific.Value;

                    if (PayblAmt == 0)
                    {
                        continue;
                    }

                    if (prevProject != Project || prevCurrency != Currency || prevDocIsEmpty != DocIsEmpty || (CommonFunctions.IsDevelopment() && prevBudgetCashFlowID != BudgetCashFlowID))
                    {
                        if (prevProject != null)
                        {
                            headerLine["PayblAmt"] = PayblAmtTotal;
                            headerLine["PayblAmtFC"] = PayblAmtFCTotal;
                            headerLine["WtAmount"] = WtAmountTotal;
                            headerLine["PensionAmount"] = PensionAmountTotal;

                            //გაკეთდება დოკუმენტი
                            try
                            {
                                DocEntry = createPaymentDocument(oForm, headerLine, AccountPaymentsLines);
                                if (DocEntry > 0)
                                {
                                    totalSuccesfull++;
                                }
                                else
                                {
                                    totalUnsuccesfull++;
                                }
                            }
                            catch (Exception ex)
                            {
                                totalUnsuccesfull++;
                                Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("DocumentNotCreated") + ". " + BDOSResources.getTranslate("ReasonIs") + ": " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            }

                        }
                        AccountHeader.Rows.Clear();
                        headerLine = AccountHeader.Rows.Add();

                        headerLine["CardCode"] = oForm.Items.Item("BPCode").Specific.Value;
                        headerLine["Currency"] = Currency;
                        headerLine["PayblCur"] = PayblCur;
                        headerLine["PayblCRt"] = 1;
                        headerLine["BankAccount"] = BankAccount;
                        headerLine["TransferAccount"] = GLAccount;
                        headerLine["ControlAccount"] = ControlAccount;
                        headerLine["WTCode"] = DocIsEmpty == "True" ? WTCode : "";
                        headerLine["accrualDate"] = DocDate;
                        headerLine["CashFlowID"] = CashFlowID;
                        headerLine["DispType"] = DispType;
                        headerLine["ChrgDtls"] = ChrgDtls;
                        headerLine["Project"] = Project;

                        if (CommonFunctions.IsDevelopment())
                        {
                            headerLine["BudgetCashFlowID"] = BudgetCashFlowID;
                            headerLine["BudgetCashFlowName"] = UDO.GetUDOFieldValueByParam("UDO_F_BDOSBUCFW_D", "Code", BudgetCashFlowID, "Name");
                        }

                        headerLine["Descrpt"] = Descrpt;

                        PayblAmtTotal = 0;
                        PayblAmtFCTotal = 0;
                        WtAmountTotal = 0;
                        PensionAmountTotal = 0;

                        AccountPaymentsLines.Rows.Clear();
                    }

                    DataRow AccountPaymentsRow = AccountPaymentsLines.Rows.Add();

                    AccountPaymentsRow["InvType"] = InvType;
                    AccountPaymentsRow["DocEntry"] = InvDocEntry;
                    AccountPaymentsRow["DocNum"] = InvDocNum;
                    AccountPaymentsRow["InstallmentId"] = InstallmentId;
                    AccountPaymentsRow["SumApplied"] = PayblAmt;

                    PayblAmtTotal = PayblAmtTotal + PayblAmt;
                    PayblAmtFCTotal = PayblAmtFCTotal + PayblAmtFC;

                    if (DocIsEmpty == "True")
                    {
                        WtAmountTotal = WtAmountTotal + WtAmount;
                        PensionAmountTotal = PensionAmountTotal + PensionAmount;
                    }
                    else
                    {
                        double PayblAmtGross = PayblAmt;
                        PayblAmtGross = PayblAmtGross / (1 - WTRate / 100) / (1 - pensionRate / 100);
                        PensionAmount = PayblAmtGross * pensionRate / 100;
                        WtAmount = (PayblAmtGross - PensionAmount) * WTRate / 100;
                        WtAmountTotal = WtAmountTotal + WtAmount;
                        PensionAmountTotal = PensionAmountTotal + PensionAmount;
                    }

                    prevCurrency = Currency;
                    prevProject = Project;
                    prevDocIsEmpty = DocIsEmpty;

                    if (CommonFunctions.IsDevelopment())
                    {
                        prevBudgetCashFlowID = BudgetCashFlowID;
                    }
                }
            }

            if (PayblAmtTotal > 0)
            {
                headerLine["PayblAmt"] = PayblAmtTotal;
                headerLine["PayblAmtFC"] = PayblAmtFCTotal;
                headerLine["WtAmount"] = WtAmountTotal;
                headerLine["PensionAmount"] = PensionAmountTotal;
                try
                {
                    DocEntry = createPaymentDocument(oForm, headerLine, AccountPaymentsLines);

                    if (DocEntry > 0)
                    {
                        totalSuccesfull++;
                    }
                    else
                    {
                        totalUnsuccesfull++;
                    }
                }
                catch (Exception ex)
                {
                    totalUnsuccesfull++;
                    Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("DocumentNotCreated") + ". " + BDOSResources.getTranslate("ReasonIs") + ": " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
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
                        SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID);
                        SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();

                        SAPbouiCOM.Condition oCon = oCons.Add();
                        oCon.Alias = "BpCode";
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = oForm.DataSources.UserDataSources.Item("BpCode").ValueEx;

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

                            setWhtCodes(oForm);
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
                        else if (oCFLEvento.ChooseFromListUID == "HouseBankAcc_CFL")
                        {
                            string Account = Convert.ToString(oDataTable.GetValue("Account", 0));
                            LanguageUtils.IgnoreErrors<string>(() => oForm.Items.Item("HBAcc").Specific.Value = Account);
                        }
                        else if (oCFLEvento.ChooseFromListUID == "GLAcc_CFL")
                        {
                            string GLAccount = Convert.ToString(oDataTable.GetValue("AcctCode", 0));
                            LanguageUtils.IgnoreErrors<string>(() => oForm.Items.Item("GLAcc").Specific.Value = GLAccount);

                            string Account = getHBAccount(oForm.Items.Item("GLAcc").Specific.Value);
                            LanguageUtils.IgnoreErrors<string>(() => oForm.Items.Item("HBAcc").Specific.Value = Account);
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
                        //}
                        //catch (Exception ex)
                        //{ ra xdebaaaaaaaaaaa?
                        //setWhtCodes(oForm);
                        //fillMTRInvoice(oForm);
                        //}
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
                        oMatrix.CommonSetting.SetCellEditable(i, 20, true);
                    }
                    else
                    {
                        oMatrix.CommonSetting.SetCellEditable(i, 20, false);
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

        private static string getHBAccount(string GLAccount)
        {
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                string query = @"SELECT ""DSC1"".""Account"" FROM ""DSC1"" 
                   WHERE ""DSC1"".""GLAccount"" = '" + GLAccount + "'";

                oRecordSet.DoQuery(query);
                if (!oRecordSet.EoF)
                {
                    return oRecordSet.Fields.Item("Account").Value.ToString();
                }
                return null;
            }
            catch
            {
                return null;
            }
            finally
            {
                oRecordSet = null;
            }


        }

        public static void setWhtCodes(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.ComboBox oItem = oForm.Items.Item("WHTax").Specific;

            try
            {

                while (oItem.ValidValues.Count > 0)
                {
                    oItem.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
                }

                SAPbobsCOM.BusinessPartners oBP;
                string cardCode = oForm.Items.Item("BPCode").Specific.Value;

                oBP = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);
                if (oBP.GetByKey(cardCode) == true)
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
                string error = ex.Message;
            }

        }

        public static void fillMTRInvoice(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.EditText oEditTextDocDate = (SAPbouiCOM.EditText)oForm.Items.Item("DocPstDt").Specific;
            string DocDateS = oEditTextDocDate.Value;
            DateTime date = Convert.ToDateTime(DateTime.ParseExact(DocDateS, "yyyyMMdd", CultureInfo.InvariantCulture));

            string dateE = date.ToString("yyyyMMdd");
            string cardCodeE = oForm.Items.Item("BPCode").Specific.Value;
            SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("InvoiceMTR");
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string errorText;

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
                 '0' AS ""LineID"",           	 
            	 SUM(TT1.""InsTotal"" - TT1.""PaidToDate""-TT1.""WTSum""+TT1.""WTApplied"") AS ""OpenAmount"",
                SUM(TT1.""WTSum"" - TT1.""WTApplied"") AS ""WTSum"",            	 
                SUM(TT1.""WTSumFC"" -TT1.""WTAppliedF"") AS ""WTSumFC"",            	 
                SUM(TT1.""InsTotal"") AS ""InsTotal"",
                 SUM(TT1.""InsTotalFC"" - TT1.""PaidFC""-TT1.""WTSumFC""+TT1.""WTAppliedF"") AS ""OpenAmountFC"",
            	 SUM(TT1.""InsTotalFC"") AS ""InsTotalFC"" 
            	FROM OPCH TT0 
            	INNER JOIN PCH6 TT1 ON TT0.""DocEntry"" = TT1.""DocEntry"" 
            	INNER JOIN OCRD T3 ON TT0.""CardCode"" = T3.""CardCode"" 
            	WHERE TT0.""DocDate"" <= '" + dateE + @"' 
	            AND TT0.""CardCode"" = N'" + cardCodeE + @"' 
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
                 TT1.""InstlmntID"" 
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
                 '0' AS ""LineID"",
            	 -SUM(TT1.""InsTotal"" - TT1.""PaidToDate""-TT1.""WTSum""+TT1.""WTApplied"")*-1 AS ""OpenAmount"",
                SUM(TT1.""WTSum"" - TT1.""WTApplied"") AS ""WTSum"",            	 
                SUM(TT1.""WTSumFC"" -TT1.""WTAppliedF"") AS ""WTSumFC"", 
            	 -SUM(TT1.""InsTotal"")*-1 AS ""InsTotal"",
                 -SUM(TT1.""InsTotalFC"" - TT1.""PaidFC""-TT1.""WTSumFC""+TT1.""WTAppliedF"")*-1 AS ""OpenAmountFC"",
            	 -SUM(TT1.""InsTotalFC"")*-1 AS ""InsTotalFC""
            	FROM OCPI TT0 
            	INNER JOIN CPI6 TT1 ON TT0.""DocEntry"" = TT1.""DocEntry"" 
            	INNER JOIN OCRD T3 ON TT0.""CardCode"" = T3.""CardCode"" 
            	WHERE  TT0.""DocDate"" <= '" + dateE + @"' 
            	AND TT0.""CardCode"" = N'" + cardCodeE + @"'
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
                 TT1.""InstlmntID""
                 
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
                 '0' AS ""LineID"",
            	 -SUM(TT1.""InsTotal"" - TT1.""PaidToDate""-TT1.""WTSum""+TT1.""WTApplied"")*-1 AS ""OpenAmount"",
                SUM(TT1.""WTSum"" - TT1.""WTApplied"") AS ""WTSum"",            	 
                SUM(TT1.""WTSumFC"" -TT1.""WTAppliedF"") AS ""WTSumFC"", 
            	
            	 -SUM(TT1.""InsTotal"")*-1 AS ""InsTotal"",
                 -SUM(TT1.""InsTotalFC"" - TT1.""PaidFC""-TT1.""WTSumFC""+TT1.""WTAppliedF"")*-1 AS ""OpenAmountFC"",
            	 -SUM(TT1.""InsTotalFC"")*-1 AS ""InsTotalFC""
            	FROM ODPO TT0 
            	INNER JOIN DPO6 TT1 ON TT0.""DocEntry"" = TT1.""DocEntry"" 
            	INNER JOIN OCRD T3 ON TT0.""CardCode"" = T3.""CardCode"" 
            	WHERE  TT0.""DocDate"" <= '" + dateE + @"' 
            	AND TT0.""CardCode"" = N'" + cardCodeE + @"'
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
                 TT1.""InstlmntID""
 
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
                int DocEntry;
                int DocNum;
                int InstallmentID;
                string DocType;
                DateTime DueDate;
                decimal OpenAmount = 0;
                decimal InsTotal = 0;
                decimal TotalPayment = 0;
                decimal TotalPaymentLocal = 0;
                decimal WTSum = 0;
                int OverdueDays = 0;

                while (!oRecordSet.EoF)
                {
                    DocEntry = Convert.ToInt32(oRecordSet.Fields.Item("DocEntry").Value);
                    DocNum = Convert.ToInt32(oRecordSet.Fields.Item("DocNum").Value);
                    InstallmentID = Convert.ToInt32(oRecordSet.Fields.Item("InstallmentID").Value);
                    DocType = Convert.ToString(oRecordSet.Fields.Item("ObjType").Value);
                    DueDate = oRecordSet.Fields.Item("DueDate").Value;
                    OpenAmount = Convert.ToDecimal(oRecordSet.Fields.Item("OpenAmountFC").Value);
                    if (OpenAmount == 0)
                    {
                        OpenAmount = Convert.ToDecimal(oRecordSet.Fields.Item("OpenAmount").Value);
                    }
                    TotalPayment = OpenAmount;

                    InsTotal = Convert.ToDecimal(oRecordSet.Fields.Item("InsTotalFC").Value);
                    if (InsTotal == 0)
                    {
                        InsTotal = Convert.ToDecimal(oRecordSet.Fields.Item("InsTotal").Value);
                    }

                    WTSum = Convert.ToDecimal(oRecordSet.Fields.Item("WTSumFC").Value);
                    if (WTSum == 0)
                    {
                        WTSum = Convert.ToDecimal(oRecordSet.Fields.Item("WTSum").Value);
                    }

                    OverdueDays = Convert.ToInt32(oRecordSet.Fields.Item("OverdueDays").Value);
                    string DocCur = Convert.ToString(oRecordSet.Fields.Item("DocCur").Value);

                    if (string.IsNullOrEmpty(DocCur))
                        DocCur = Program.MainCurrencySapCode;

                    oDataTable.Rows.Add();
                    oDataTable.SetValue("LineNum", rowIndex, rowIndex + 1);
                    oDataTable.SetValue("CheckBox", rowIndex, "N");
                    oDataTable.SetValue("DocEntry", rowIndex, DocEntry);
                    oDataTable.SetValue("DocNum", rowIndex, DocNum);
                    oDataTable.SetValue("InstallmentID", rowIndex, oRecordSet.Fields.Item("InstallmentID").Value);
                    oDataTable.SetValue("LineID", rowIndex, oRecordSet.Fields.Item("LineID").Value);
                    oDataTable.SetValue("DocType", rowIndex, DocType);
                    oDataTable.SetValue("DocDate", rowIndex, oRecordSet.Fields.Item("DocDate").Value.ToString("yyyyMMdd") == "18991230" ? "" : oRecordSet.Fields.Item("DocDate").Value.ToString("yyyyMMdd"));
                    oDataTable.SetValue("DueDate", rowIndex, oRecordSet.Fields.Item("DueDate").Value.ToString("yyyyMMdd") == "18991230" ? "" : oRecordSet.Fields.Item("DueDate").Value.ToString("yyyyMMdd"));
                    oDataTable.SetValue("Arrears", rowIndex, OverdueDays >= 0 ? "*" : "");
                    oDataTable.SetValue("OverdueDays", rowIndex, OverdueDays);
                    oDataTable.SetValue("Comments", rowIndex, oRecordSet.Fields.Item("Comments").Value);
                    oDataTable.SetValue("Total", rowIndex, Convert.ToDouble(InsTotal));
                    oDataTable.SetValue("WTSum", rowIndex, Convert.ToDouble(WTSum));
                    oDataTable.SetValue("BalanceDue", rowIndex, Convert.ToDouble(OpenAmount));
                    oDataTable.SetValue("TotalPayment", rowIndex, Convert.ToDouble(TotalPayment));
                    oDataTable.SetValue("Currency", rowIndex, DocCur);
                    oDataTable.SetValue("TotalPaymentLocal", rowIndex, Convert.ToDouble(TotalPaymentLocal));
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

                    oRecordSet.MoveNext();
                    rowIndex++;
                }

                SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("InvoiceMTR").Specific));
                oForm.Freeze(true);
                oMatrix.Clear();
                oMatrix.LoadFromDataSource();
                oMatrix.AutoResizeColumns();

                setEditableSetting(oForm);
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
            finally
            {
                oForm.Update();
                oForm.Freeze(false);
                oRecordSet = null;
            }
        }

        private static void setEditableSetting(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("InvoiceMTR").Specific));

            int i = 1;
            while (i <= oMatrix.RowCount)
            {
                string RowDocEntry = oMatrix.GetCellSpecific("DocEntry", i).Value;
                oMatrix.CommonSetting.SetCellEditable(i, 4, RowDocEntry == "0");
                oMatrix.CommonSetting.SetCellEditable(i, 17, RowDocEntry == "0");
                i++;
            }
        }

        public static void updateRow(SAPbouiCOM.Form oForm, bool cfwIdChng, bool descrptionChng, int rowIndex = 0)
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
