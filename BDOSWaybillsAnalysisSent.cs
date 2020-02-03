using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Xml;
using System.Threading;

namespace BDO_Localisation_AddOn
{
    static partial class BDOSWaybillsAnalysisSent
    {
        public static void createForm(  out string errorText)
        {
            errorText = null;

            //ფორმის აუცილებელი თვისებები
            Dictionary<string, object> formProperties = new Dictionary<string, object>();
            formProperties.Add("UniqueID", "BDOSWBSAn");
            formProperties.Add("BorderStyle", SAPbouiCOM.BoFormBorderStyle.fbs_Sizable);
            formProperties.Add("Title", BDOSResources.getTranslate("SentWaybillsAnalysis"));
            formProperties.Add("Left", 558);
            formProperties.Add("ClientWidth", 750);
            formProperties.Add("Top", 335);
            formProperties.Add("ClientHeight", 400);

            SAPbouiCOM.Form oForm;
            bool newForm;
            bool formExist = FormsB1.createForm( formProperties, out oForm, out newForm, out errorText);

            if (errorText != null)
            {
                return;
            }

            if (formExist == true)
            {
                if (newForm == true)
                {
                    Dictionary<string, object> formItems = null;

                    oForm.DataSources.DataTables.Add("WbTable");

                    SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("WbTable");
                    oDataTable.Columns.Add("BaseCard", SAPbouiCOM.BoFieldsType.ft_Text, 50);//0                
                    oDataTable.Columns.Add("ComparStat", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 50);//1
                    oDataTable.Columns.Add("WB_number", SAPbouiCOM.BoFieldsType.ft_Text, 50);//2

                    oDataTable.Columns.Add("LineNum", SAPbouiCOM.BoFieldsType.ft_Text, 50); //3  
                    oDataTable.Columns.Add("LicTradNum", SAPbouiCOM.BoFieldsType.ft_Text, 50);//4
                    oDataTable.Columns.Add("CardName", SAPbouiCOM.BoFieldsType.ft_Text, 50);//5                   
                    oDataTable.Columns.Add("WB_begDate", SAPbouiCOM.BoFieldsType.ft_Date, 50);//6
                    oDataTable.Columns.Add("TYPE", SAPbouiCOM.BoFieldsType.ft_Text, 50);//7
                    oDataTable.Columns.Add("WB_status", SAPbouiCOM.BoFieldsType.ft_Text, 50);//8
                    oDataTable.Columns.Add("WB_strAddrs", SAPbouiCOM.BoFieldsType.ft_Text, 50);//9
                    oDataTable.Columns.Add("WB_endAddrs", SAPbouiCOM.BoFieldsType.ft_Text, 50);//10
                    oDataTable.Columns.Add("WB_delvDate", SAPbouiCOM.BoFieldsType.ft_Date, 50);//11

                    oDataTable.Columns.Add("WB_ID", SAPbouiCOM.BoFieldsType.ft_Text, 50);//12
                    oDataTable.Columns.Add("WB_actDate", SAPbouiCOM.BoFieldsType.ft_Date, 50);//13              
                    oDataTable.Columns.Add("WB_vehicNum", SAPbouiCOM.BoFieldsType.ft_Text, 50);//14
                    oDataTable.Columns.Add("WB_drivTin", SAPbouiCOM.BoFieldsType.ft_Text, 50);//15
                    oDataTable.Columns.Add("WB_trnsExpn", SAPbouiCOM.BoFieldsType.ft_Text, 50);//16
                    oDataTable.Columns.Add("WBLD_Doc", SAPbouiCOM.BoFieldsType.ft_Text, 50);//17

                    oDataTable.Columns.Add("DocDate", SAPbouiCOM.BoFieldsType.ft_Date, 50);//18
                    oDataTable.Columns.Add("DocNum", SAPbouiCOM.BoFieldsType.ft_Text, 50);//19
                    oDataTable.Columns.Add("BaseDoc", SAPbouiCOM.BoFieldsType.ft_Text, 50);//20
                    oDataTable.Columns.Add("WhsTo", SAPbouiCOM.BoFieldsType.ft_Text, 50);//21
                    oDataTable.Columns.Add("WhsFrom", SAPbouiCOM.BoFieldsType.ft_Text, 50);//22               

                    oDataTable.Columns.Add("RS_Status", SAPbouiCOM.BoFieldsType.ft_Text, 50);//23 
                    oDataTable.Columns.Add("RSAmount", SAPbouiCOM.BoFieldsType.ft_Sum, 50);//24
                    oDataTable.Columns.Add("Gtotal", SAPbouiCOM.BoFieldsType.ft_Sum, 50);//25

                    //დეტალური როცა არის
                    oDataTable.Columns.Add("ItemCode", SAPbouiCOM.BoFieldsType.ft_Text, 50);//26
                    oDataTable.Columns.Add("InvntryUom", SAPbouiCOM.BoFieldsType.ft_Text, 50);//27
                    oDataTable.Columns.Add("BarCode", SAPbouiCOM.BoFieldsType.ft_Text, 50);//28
                    oDataTable.Columns.Add("Quantity", SAPbouiCOM.BoFieldsType.ft_Quantity, 50);//29
                    oDataTable.Columns.Add("RSQuantity", SAPbouiCOM.BoFieldsType.ft_Quantity, 50);//30
                    oDataTable.Columns.Add("RSName", SAPbouiCOM.BoFieldsType.ft_Text, 50);//31
                    //დეტალური როცა არის

                    oDataTable.Columns.Add("BaseType", SAPbouiCOM.BoFieldsType.ft_Text, 50);//32 - საფუძველი დოკუმენტის ტიპი

                    string itemName = "";
                    int left = 6;
                    int Top = 5;
                    //int leftSC = 400;
                    //List<string> listValidValues;

                    formItems = new Dictionary<string, object>();
                    itemName = "DateFrom";
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 50);
                    formItems.Add("Top", Top);
                    formItems.Add("Caption", BDOSResources.getTranslate("StartDate"));
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left = left + 50 + 10;

                    string startOfMonthStr = DateTime.Today.ToString("yyyyMMdd");
                    formItems = new Dictionary<string, object>();
                    itemName = "StartDate";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_DATE);
                    formItems.Add("Length", 1);
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("ValueEx", startOfMonthStr);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left = left + 100 + 10;

                    itemName = "dateTo";
                    formItems = new Dictionary<string, object>();
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 50);
                    formItems.Add("Top", Top);
                    formItems.Add("Caption", "To");
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left = left + 50 + 10;

                    string endOfMonthStr = DateTime.Today.ToString("yyyyMMdd");
                    formItems = new Dictionary<string, object>();
                    itemName = "EndDate";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_DATE);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("Length", 1);
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("ValueEx", endOfMonthStr);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left = left + 100 + 10;

                    formItems = new Dictionary<string, object>();
                    itemName = "WbFillTb";
                    formItems.Add("Caption", BDOSResources.getTranslate("Fill"));
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    //რიგი 2
                    Top = Top + 20;
                    left = 6;

                    formItems = new Dictionary<string, object>();
                    itemName = "WBDocTpSt";
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 50);
                    formItems.Add("Top", Top);
                    formItems.Add("Caption", BDOSResources.getTranslate("Type"));
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left = left + 50 + 10;

                    Dictionary<string, string> listValidValuesDict = null;

                    listValidValuesDict = new Dictionary<string, string>();
                    listValidValuesDict.Add("0", BDOSResources.getTranslate("WithoutFilter"));
                    listValidValuesDict.Add("1", BDOSResources.getTranslate("InternalShipment"));
                    listValidValuesDict.Add("2", BDOSResources.getTranslate("ARInvoice") + " " + BDOSResources.getTranslate("WithTransport"));
                    listValidValuesDict.Add("3", BDOSResources.getTranslate("ARInvoice") + " " + BDOSResources.getTranslate("WithoutTransport"));
                    listValidValuesDict.Add("4", BDOSResources.getTranslate("Distribution"));
                    listValidValuesDict.Add("5", BDOSResources.getTranslate("Return"));
                    listValidValuesDict.Add("6", BDOSResources.getTranslate("SubWaybill"));

                    formItems = new Dictionary<string, object>();
                    itemName = "WBDocTp";
                    formItems.Add("Size", 20);
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("Length", 30);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                    formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("ValidValues", listValidValuesDict);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }


                    left = left + 100 + 10;

                    formItems = new Dictionary<string, object>();
                    itemName = "WBStatusSt";
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 50);
                    formItems.Add("Top", Top);
                    formItems.Add("Caption", BDOSResources.getTranslate("Status"));
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left = left + 50 + 10;

                    listValidValuesDict = new Dictionary<string, string>();
                    listValidValuesDict.Add("-99", BDOSResources.getTranslate("WithoutFilter"));
                    listValidValuesDict.Add("-2", BDOSResources.getTranslate("Canceled"));
                    listValidValuesDict.Add("-1", BDOSResources.getTranslate("deleted"));
                    listValidValuesDict.Add("0", BDOSResources.getTranslate("Saved"));
                    listValidValuesDict.Add("1", BDOSResources.getTranslate("Active"));
                    listValidValuesDict.Add("2", BDOSResources.getTranslate("finished"));
                    listValidValuesDict.Add("7", BDOSResources.getTranslate("SentToTransporter"));

                    formItems = new Dictionary<string, object>();
                    itemName = "WBStatus";
                    formItems.Add("Size", 20);
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("Length", 30);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                    formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("ValidValues", listValidValuesDict);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    //გაფერადება
                    left = left + 100 + 10;

                    formItems = new Dictionary<string, object>();
                    itemName = "StGrColor";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("Length", 1);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 150);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 14);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("SetGridColor"));
                    formItems.Add("ValOff", "N");
                    formItems.Add("ValOn", "Y");
                    formItems.Add("DisplayDesc", true);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    //რიგი 3
                    Top = Top + 20;
                    left = 6;

                    formItems = new Dictionary<string, object>();
                    itemName = "ClientIDSt";
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 50);
                    formItems.Add("Top", Top);
                    formItems.Add("Caption", BDOSResources.getTranslate("BP"));
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left = left + 50 + 10;

                    bool multiSelection = false;
                    string objectType = "2"; //SAPbouiCOM.BoLinkedObject.lf_BusinessPartner, Business Partner object 
                    string uniqueID_lf_BusinessPartnerCFL = "BusinessPartner_CFL";
                    FormsB1.addChooseFromList( oForm, multiSelection, objectType, uniqueID_lf_BusinessPartnerCFL);

                    //პირობის დადება ბიზნესპარტნიორის არჩევის სიაზე
                    SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(uniqueID_lf_BusinessPartnerCFL);
                    SAPbouiCOM.Conditions oCons = oCFL.GetConditions();
                    SAPbouiCOM.Condition oCon = oCons.Add();
                    oCon.Alias = "CardType";
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCon.CondVal = "C"; //კლიენტი
                    oCFL.SetConditions(oCons);

                    formItems = new Dictionary<string, object>();
                    itemName = "ClientID";
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
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("ChooseFromListUID", uniqueID_lf_BusinessPartnerCFL);
                    formItems.Add("ChooseFromListAlias", "CardCode");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    //compare status
                    left = left + 100 + 10;

                    formItems = new Dictionary<string, object>();
                    itemName = "CompStatSt";
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 50);
                    formItems.Add("Top", Top);
                    //formItems.Add("Height", 38);
                    formItems.Add("Caption", BDOSResources.getTranslate("CompSt"));
                    formItems.Add("Description", BDOSResources.getTranslate("CompareStatus"));
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left = left + 50 + 10;

                    listValidValuesDict = new Dictionary<string, string>();
                    listValidValuesDict.Add("0", BDOSResources.getTranslate("WithoutFilter"));
                    listValidValuesDict.Add("1", BDOSResources.getTranslate("OnlyOnSite"));
                    listValidValuesDict.Add("2", BDOSResources.getTranslate("OnlyOnProgram"));
                    listValidValuesDict.Add("3", BDOSResources.getTranslate("AmountsNotEqual"));
                    listValidValuesDict.Add("4", BDOSResources.getTranslate("EqualAmounts"));
                    listValidValuesDict.Add("5", BDOSResources.getTranslate("Linked") + " " + BDOSResources.getTranslate("Document") + " " + BDOSResources.getTranslate("NotPosted"));
                    listValidValuesDict.Add("6", BDOSResources.getTranslate("SavedStatus"));

                    formItems = new Dictionary<string, object>();
                    itemName = "CompStat";
                    formItems.Add("Size", 20);
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("Length", 30);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                    formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("ValidValues", listValidValuesDict);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    //ვარიანტი
                    left = left + 100 + 10;

                    formItems = new Dictionary<string, object>();
                    itemName = "OptionSt";
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 50);
                    formItems.Add("Top", Top);
                    formItems.Add("Caption", BDOSResources.getTranslate("Option"));
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left = left + 50 + 10;

                    listValidValuesDict = new Dictionary<string, string>();
                    listValidValuesDict.Add("0", BDOSResources.getTranslate("Details"));
                    listValidValuesDict.Add("1", BDOSResources.getTranslate("General"));

                    formItems = new Dictionary<string, object>();
                    itemName = "Option";
                    formItems.Add("Size", 20);
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("Length", 30);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                    formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("ValidValues", listValidValuesDict);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }


                    SAPbouiCOM.ComboBox oComboBox_WBDocTp = (SAPbouiCOM.ComboBox)oForm.Items.Item("WBDocTp").Specific;
                    oComboBox_WBDocTp.Select("0", SAPbouiCOM.BoSearchKey.psk_ByValue);

                    SAPbouiCOM.ComboBox oComboBox_WBStatus = (SAPbouiCOM.ComboBox)oForm.Items.Item("WBStatus").Specific;
                    oComboBox_WBStatus.Select("-99", SAPbouiCOM.BoSearchKey.psk_ByValue);

                    SAPbouiCOM.ComboBox oComboBox_CompStat = (SAPbouiCOM.ComboBox)oForm.Items.Item("CompStat").Specific;
                    oComboBox_CompStat.Select("0", SAPbouiCOM.BoSearchKey.psk_ByValue);

                    SAPbouiCOM.ComboBox oComboBox_Option = (SAPbouiCOM.ComboBox)oForm.Items.Item("Option").Specific;
                    oComboBox_Option.Select("1", SAPbouiCOM.BoSearchKey.psk_ByValue);

                    //Grid
                    Top = Top + 30;
                    left = 6;

                    itemName = "WbTable";
                    formItems = new Dictionary<string, object>();
                    formItems.Add("isDataSource", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_GRID);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 750);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 280);
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    Top = Top + 280 + 6;
                    left = 750 - 120;

                    formItems = new Dictionary<string, object>();
                    itemName = "btnColl";
                    formItems.Add("Size", 20);
                    formItems.Add("Caption", BDOSResources.getTranslate("Collapse"));
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 110);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 20);
                    formItems.Add("UID", itemName);
                    formItems.Add("AffectsFormMode", false);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left = 750 - 120 - 115;

                    formItems = new Dictionary<string, object>();
                    itemName = "btnExp";
                    formItems.Add("Size", 20);
                    formItems.Add("Caption", BDOSResources.getTranslate("Expand"));
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 110);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 20);
                    formItems.Add("UID", itemName);
                    formItems.Add("AffectsFormMode", false);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                }

                oForm.Visible = true;
                oForm.Select();
            }
        }

        public static void updateGrid(  SAPbouiCOM.Form oForm, out string errorText)
        {

            Dictionary<string, string> rsSettings = CompanyDetails.getRSSettings( out errorText);
            if (errorText != null)
            {
                return;
            }


            string su = rsSettings["SU"];
            string sp = rsSettings["SP"];
            WayBill oWayBill = new WayBill(su, sp, rsSettings["ProtocolType"]);

            bool chek_service_user = oWayBill.chek_service_user(su, sp, out errorText);
            if (chek_service_user == false)
            {
                errorText = BDOSResources.getTranslate("ServiceUserPasswordNotCorrect");
                return;
            }

            DateTime startDate;
            string startDateStr = oForm.DataSources.UserDataSources.Item("StartDate").ValueEx;
            DateTime BeginDate = new DateTime(1, 1, 1);

            if (DateTime.TryParseExact(startDateStr, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture, DateTimeStyles.None, out startDate))
            {
                BeginDate = startDate;
            }

            DateTime endDate;
            string endDateStr = oForm.DataSources.UserDataSources.Item("EndDate").ValueEx;
            DateTime EndDate = DateTime.Today;

            if (DateTime.TryParseExact(endDateStr, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture, DateTimeStyles.None, out endDate))
            {
                EndDate = endDate;
            }

            string itypes = oForm.DataSources.UserDataSources.Item("WBDocTp").ValueEx;

            string cardCode = oForm.DataSources.UserDataSources.Item("ClientID").Value;
            cardCode = cardCode.Trim();
            string buyer_tin = "";

            if (cardCode != "")
            {
                SAPbobsCOM.BusinessPartners oBP;
                oBP = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);
                oBP.GetByKey(cardCode);

                buyer_tin = oBP.UserFields.Fields.Item("LicTradNum").Value;
            }

            string statuses = oForm.DataSources.UserDataSources.Item("WBStatus").ValueEx;
            string FOption = oForm.DataSources.UserDataSources.Item("option").ValueEx;

            string query = getQueryText(BeginDate, EndDate, cardCode, itypes, statuses, FOption);


            if (itypes == "0" || itypes == "")
            {
                itypes = "1,2,3,4,5,6";
            }


            if (statuses == "-99" || statuses == "")
            {
                statuses = ",,1,2,-1,-2,7,";
            }
            else if (statuses == "0")
            {
                statuses = ",,";
            }
            else
            {
                statuses = "," + statuses + ",";
            }

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRecordSet.DoQuery(query);

            DateTime EndDateForWS = EndDate.AddDays(1).AddMilliseconds(-1);

            DataTable RSDataTable = new DataTable();
            RSDataTable.Columns.Add("RowLinked", typeof(string));
            RSDataTable.Columns.Add("ID", typeof(string));
            RSDataTable.Columns.Add("WAYBILL_NUMBER", typeof(string));
            RSDataTable.Columns.Add("FULL_AMOUNT", typeof(string));
            RSDataTable.Columns.Add("STATUS", typeof(string));
            RSDataTable.Columns.Add("TIN", typeof(string));
            RSDataTable.Columns.Add("NAME", typeof(string));
            RSDataTable.Columns.Add("BEGIN_DATE", typeof(string));
            RSDataTable.Columns.Add("TYPE", typeof(string));
            RSDataTable.Columns.Add("START_ADDRESS", typeof(string));
            RSDataTable.Columns.Add("END_ADDRESS", typeof(string));
            RSDataTable.Columns.Add("DELIVERY_DATE", typeof(string));
            RSDataTable.Columns.Add("ACTIVATE_DATE", typeof(string));
            RSDataTable.Columns.Add("CAR_NUMBER", typeof(string));
            RSDataTable.Columns.Add("DRIVER_TIN", typeof(string));
            RSDataTable.Columns.Add("TRANSPORT_COAST", typeof(string));
            RSDataTable.Columns.Add("W_NAME", typeof(string));
            RSDataTable.Columns.Add("BAR_CODE", typeof(string));
            RSDataTable.Columns.Add("AMOUNT", typeof(string));
            RSDataTable.Columns.Add("Quantity", typeof(string));


            DateTime startDateParam = new DateTime();
            DateTime endDateParam = new DateTime();
            startDateParam = startDate;

            while (startDateParam < EndDateForWS)
            {
                endDateParam = startDateParam.AddDays(3);

                if (endDateParam > EndDateForWS)
                {
                    endDateParam = EndDateForWS;
                }

                DataTable RSDataTable_part = oWayBill.get_waybill_goods_list(startDateParam, endDateParam, itypes, buyer_tin, statuses, "", "", out errorText);

                for (int i = 0; i < RSDataTable_part.Rows.Count; i++)
                {


                    DataRow RSDataRow = RSDataTable.Rows.Add();

                    RSDataRow["RowLinked"] = "N";

                    RSDataRow["ID"] = RSDataTable.Rows[i]["ID"].ToString();

                    RSDataRow["WAYBILL_NUMBER"] = RSDataTable_part.Rows[i]["WAYBILL_NUMBER"];

                    RSDataRow["FULL_AMOUNT"] = RSDataTable_part.Rows[i]["FULL_AMOUNT"];

                    RSDataRow["STATUS"] = RSDataTable_part.Rows[i]["STATUS"];

                    RSDataRow["TIN"] = RSDataTable_part.Rows[i]["SELLER_TIN"];

                    RSDataRow["NAME"] = RSDataTable_part.Rows[i]["SELLER_NAME"];

                    RSDataRow["BEGIN_DATE"] = RSDataTable_part.Rows[i]["BEGIN_DATE"];

                    RSDataRow["TYPE"] = RSDataTable_part.Rows[i]["TYPE"];

                    RSDataRow["START_ADDRESS"] = RSDataTable_part.Rows[i]["START_ADDRESS"];

                    RSDataRow["END_ADDRESS"] = RSDataTable_part.Rows[i]["END_ADDRESS"];

                    RSDataRow["DELIVERY_DATE"] = RSDataTable_part.Rows[i]["DELIVERY_DATE"];

                    RSDataRow["ACTIVATE_DATE"] = RSDataTable_part.Rows[i]["ACTIVATE_DATE"];

                    RSDataRow["CAR_NUMBER"] = RSDataTable_part.Rows[i]["CAR_NUMBER"];

                    RSDataRow["TRANSPORT_COAST"] = RSDataTable_part.Rows[i]["TRANSPORT_COAST"];

                    RSDataRow["W_NAME"] = RSDataTable_part.Rows[i]["W_NAME"];

                    RSDataRow["BAR_CODE"] = RSDataTable_part.Rows[i]["BAR_CODE"];

                    RSDataRow["AMOUNT"] = RSDataTable_part.Rows[i]["AMOUNT"];

                    RSDataRow["Quantity"] = RSDataTable_part.Rows[i]["Quantity"];

                }

                startDateParam = endDateParam;
            }

            //თუ არადეტალურია, ნომენკლატურის მონაცემები აღარ იქნება RS ის ცხრილში
            if (FOption == "1" && RSDataTable.Rows.Count > 0)
            {
                RSDataTable = RSDataTable.AsEnumerable().GroupBy(r => new
                {
                    Col1 = r["RowLinked"],
                    Col3 = r["ID"],
                    Col4 = r["WAYBILL_NUMBER"],
                    Col5 = r["FULL_AMOUNT"],
                    Col6 = r["STATUS"],
                    Col7 = r["TIN"],
                    Col8 = r["NAME"],
                    Col9 = r["BEGIN_DATE"],
                    Col10 = r["START_ADDRESS"],
                    Col12 = r["END_ADDRESS"],
                    Col13 = r["DELIVERY_DATE"],
                    Col14 = r["ACTIVATE_DATE"],
                    Col15 = r["CAR_NUMBER"],
                    Col16 = r["DRIVER_TIN"],
                    Col17 = r["TRANSPORT_COAST"]
                })
                                                       .Select(g => g.OrderBy(r => r["ID"]).First())
                                                       .CopyToDataTable();
            }



            SAPbouiCOM.DataTable oDataTable;

            oDataTable = oForm.DataSources.DataTables.Item("WbTable");
            oDataTable.Rows.Clear();

            string XML = "";
            XML = oDataTable.GetAsXML();
            XML = XML.Replace("<Rows/></DataTable>", "");

            StringBuilder Sbuilder = new StringBuilder();
            Sbuilder.Append(XML);
            Sbuilder.Append("<Rows>");

            int count = 0;

            string itemFindParameter = "";

            if (rsSettings["ItemCode"] == "0")
            {
                itemFindParameter = "ItemCode";
            }
            else if (rsSettings["ItemCode"] == "1")
            {
                itemFindParameter = "SWW";
            }
            else if (rsSettings["ItemCode"] == "2")
            {
                itemFindParameter = "CodeBars";
            }

            string FilterCompStat = oForm.DataSources.UserDataSources.Item("CompStat").ValueEx;

            int WBLD_Doc;
            int BaseDocNum;
            int DocEntry;
            string WBNum;
            string U_wbID;
            string wbtype;
            string barCode;
            double Amount;
            double RSAmount;
            double RSAmount_Full;
            double RSQuantity;
            string RS_W_NAME;
            string wbCompStat;
            string RS_STATUS;

            while (!oRecordSet.EoF)
            {
                WBNum = oRecordSet.Fields.Item("U_number").Value;
                U_wbID = oRecordSet.Fields.Item("U_wblID").Value;
                wbtype = oRecordSet.Fields.Item("type").Value;
                WBNum = WBNum.Trim();
                barCode = "";
                if (FOption != "1")
                {
                    barCode = oRecordSet.Fields.Item(itemFindParameter).Value;
                }

                Amount = oRecordSet.Fields.Item("Gtotal").Value;
                //Amount_Full = oRecordSet.Fields.Item("Gtotal").Value;
                RSAmount = 0;
                RSAmount_Full = 0;
                RSQuantity = 0;
                RS_W_NAME = "";
                wbCompStat = "";
                RS_STATUS = "";

                if (WBNum != "")
                {
                    bool foundonRS = false;

                    if (RSDataTable.Rows.Count > 0)
                    {
                        if (RSDataTable.Columns.Contains("WAYBILL_NUMBER"))
                        {
                            DataRow[] foundRows;

                            if (FOption != "1")
                            {
                                foundRows = RSDataTable.Select("WAYBILL_NUMBER = '" + WBNum + "'" + " and " + "BAR_CODE = '" + barCode.Replace("'", "''") + "'");
                            }
                            else
                            {
                                foundRows = RSDataTable.Select("WAYBILL_NUMBER = '" + WBNum + "'");
                            }

                            for (int i = 0; i < foundRows.Length; i++)
                            {
                                foundonRS = true;
                                foundRows[i]["RowLinked"] = "Y";

                                RSAmount = Convert.ToDouble(foundRows[i]["AMOUNT"], System.Globalization.CultureInfo.InvariantCulture);
                                RSAmount_Full = Convert.ToDouble(foundRows[i]["FULL_AMOUNT"], System.Globalization.CultureInfo.InvariantCulture);
                                RSQuantity = Convert.ToDouble(foundRows[i]["Quantity"], System.Globalization.CultureInfo.InvariantCulture);
                                RS_W_NAME = foundRows[i]["W_NAME"].ToString();
                                RS_STATUS = foundRows[i]["STATUS"].ToString();
                                if (wbtype == "5")
                                {
                                    RSAmount = RSAmount * (-1);
                                    RSAmount_Full = RSAmount_Full * (-1);
                                    RSQuantity = RSQuantity * (-1);
                                }

                            }
                        }
                    }

                    if (!foundonRS)
                    {
                        wbCompStat = "2";
                    }
                    else if (oRecordSet.Fields.Item("CANCELED").Value == "Y")
                    {
                        wbCompStat = "5";
                    }
                    else if (FOption == "1" && Amount != RSAmount_Full)
                    {
                        wbCompStat = "3";
                    }
                    else if (FOption != "1" && Amount != RSAmount)
                    {
                        wbCompStat = "3";
                    }
                    else
                    {
                        wbCompStat = "4";
                    }
                }
                else if (U_wbID == "")
                {
                    wbCompStat = "2";
                }
                else
                {
                    wbCompStat = "6";
                }

                if (FilterCompStat != "0" & FilterCompStat != "")
                {
                    if (FilterCompStat != wbCompStat)
                    {
                        oRecordSet.MoveNext();
                        continue;
                    }
                }

                Sbuilder.Append("<Row>");
                Sbuilder.Append("<Cell> <ColumnUid>BaseCard</ColumnUid> <Value>");
                Sbuilder = CommonFunctions.AppendXML(Sbuilder, oRecordSet.Fields.Item("BaseCard").Value);
                Sbuilder.Append("</Value></Cell>");

                Sbuilder.Append("<Cell> <ColumnUid>ComparStat</ColumnUid> <Value>");
                Sbuilder = CommonFunctions.AppendXML(Sbuilder, wbCompStat);
                Sbuilder.Append("</Value></Cell>");

                Sbuilder.Append("<Cell> <ColumnUid>WB_number</ColumnUid> <Value>");
                Sbuilder = CommonFunctions.AppendXML(Sbuilder, WBNum);
                Sbuilder.Append("</Value></Cell>");

                Sbuilder.Append("<Cell> <ColumnUid>LineNum</ColumnUid> <Value>");
                Sbuilder = CommonFunctions.AppendXML(Sbuilder, (count + 1).ToString());
                Sbuilder.Append("</Value></Cell>");

                Sbuilder.Append("<Cell> <ColumnUid>LicTradNum</ColumnUid> <Value>");
                Sbuilder = CommonFunctions.AppendXML(Sbuilder, oRecordSet.Fields.Item("LicTradNum").Value);
                Sbuilder.Append("</Value></Cell>");

                Sbuilder.Append("<Cell> <ColumnUid>CardName</ColumnUid> <Value>");
                Sbuilder = CommonFunctions.AppendXML(Sbuilder, oRecordSet.Fields.Item("CardName").Value);
                Sbuilder.Append("</Value></Cell>");

                Sbuilder.Append("<Cell> <ColumnUid>TYPE</ColumnUid> <Value>");
                Sbuilder = CommonFunctions.AppendXML(Sbuilder, wbtype);
                Sbuilder.Append("</Value></Cell>");

                if (wbCompStat != "2")
                {
                    Sbuilder.Append("<Cell> <ColumnUid>WB_begDate</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, oRecordSet.Fields.Item("U_begDate").Value.ToString("yyyyMMdd") == "18991230" ? "" : oRecordSet.Fields.Item("U_begDate").Value.ToString("yyyyMMdd"));
                    Sbuilder.Append("</Value></Cell>");

                    Sbuilder.Append("<Cell> <ColumnUid>WB_status</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, oRecordSet.Fields.Item("WB_Status").Value);
                    Sbuilder.Append("</Value></Cell>");

                    Sbuilder.Append("<Cell> <ColumnUid>WB_strAddrs</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, oRecordSet.Fields.Item("U_strAddrs").Value);
                    Sbuilder.Append("</Value></Cell>");

                    Sbuilder.Append("<Cell> <ColumnUid>WB_endAddrs</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, oRecordSet.Fields.Item("U_endAddrs").Value);
                    Sbuilder.Append("</Value></Cell>");

                    Sbuilder.Append("<Cell> <ColumnUid>WB_delvDate</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, oRecordSet.Fields.Item("U_delvDate").Value.ToString("yyyyMMdd") == "18991230" ? "" : oRecordSet.Fields.Item("U_delvDate").Value.ToString("yyyyMMdd"));
                    Sbuilder.Append("</Value></Cell>");

                    Sbuilder.Append("<Cell> <ColumnUid>WB_ID</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, oRecordSet.Fields.Item("U_wblID").Value);
                    Sbuilder.Append("</Value></Cell>");

                    Sbuilder.Append("<Cell> <ColumnUid>WB_actDate</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, oRecordSet.Fields.Item("U_actDate").Value.ToString("yyyyMMdd") == "18991230" ? "" : oRecordSet.Fields.Item("U_actDate").Value.ToString("yyyyMMdd"));
                    Sbuilder.Append("</Value></Cell>");

                    Sbuilder.Append("<Cell> <ColumnUid>WB_vehicNum</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, oRecordSet.Fields.Item("U_vehicNum").Value);
                    Sbuilder.Append("</Value></Cell>");

                    Sbuilder.Append("<Cell> <ColumnUid>WB_drivTin</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, oRecordSet.Fields.Item("U_drivTin").Value);
                    Sbuilder.Append("</Value></Cell>");

                    Sbuilder.Append("<Cell> <ColumnUid>WB_trnsExpn</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, FormsB1.ConvertDecimalToString(Convert.ToDecimal(oRecordSet.Fields.Item("U_trnsExpn").Value)));
                    Sbuilder.Append("</Value></Cell>");

                    WBLD_Doc = (int)oRecordSet.Fields.Item("WBLD_Doc").Value;
                    Sbuilder.Append("<Cell> <ColumnUid>WBLD_Doc</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, (WBLD_Doc == 0 ? "" : WBLD_Doc.ToString()));
                    Sbuilder.Append("</Value></Cell>");
                }
                Sbuilder.Append("<Cell> <ColumnUid>WB_begDate</ColumnUid> <Value>");
                Sbuilder = CommonFunctions.AppendXML(Sbuilder, oRecordSet.Fields.Item("BaseDocDate").Value.ToString("yyyyMMdd") == "18991230" ? "" : oRecordSet.Fields.Item("BaseDocDate").Value.ToString("yyyyMMdd"));
                Sbuilder.Append("</Value></Cell>");

                BaseDocNum = (int)oRecordSet.Fields.Item("BaseDocNum").Value;
                Sbuilder.Append("<Cell> <ColumnUid>DocNum</ColumnUid> <Value>");
                Sbuilder = CommonFunctions.AppendXML(Sbuilder, (BaseDocNum == 0 ? "" : BaseDocNum.ToString()));
                Sbuilder.Append("</Value></Cell>");

                DocEntry = (int)oRecordSet.Fields.Item("DocEntry").Value;
                Sbuilder.Append("<Cell> <ColumnUid>BaseDoc</ColumnUid> <Value>");
                Sbuilder = CommonFunctions.AppendXML(Sbuilder, (DocEntry == 0 ? "" : DocEntry.ToString()));
                Sbuilder.Append("</Value></Cell>");

                Sbuilder.Append("<Cell> <ColumnUid>WhsTo</ColumnUid> <Value>");
                Sbuilder = CommonFunctions.AppendXML(Sbuilder, oRecordSet.Fields.Item("WhsTo").Value);
                Sbuilder.Append("</Value></Cell>");

                Sbuilder.Append("<Cell> <ColumnUid>WhsFrom</ColumnUid> <Value>");
                Sbuilder = CommonFunctions.AppendXML(Sbuilder, oRecordSet.Fields.Item("WhsFrom").Value);
                Sbuilder.Append("</Value></Cell>");

                Sbuilder.Append("<Cell> <ColumnUid>RS_Status</ColumnUid> <Value>");
                Sbuilder = CommonFunctions.AppendXML(Sbuilder, RS_STATUS);
                Sbuilder.Append("</Value></Cell>");

                if (FOption != "1")
                {
                    Sbuilder.Append("<Cell> <ColumnUid>RSAmount</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, FormsB1.ConvertDecimalToString(Convert.ToDecimal(RSAmount)));
                    Sbuilder.Append("</Value></Cell>");

                    Sbuilder.Append("<Cell> <ColumnUid>Gtotal</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, FormsB1.ConvertDecimalToString(Convert.ToDecimal(Amount)));
                    Sbuilder.Append("</Value></Cell>");

                    Sbuilder.Append("<Cell> <ColumnUid>ItemCode</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, oRecordSet.Fields.Item("ItemCode").Value);
                    Sbuilder.Append("</Value></Cell>");

                    Sbuilder.Append("<Cell> <ColumnUid>InvntryUom</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, oRecordSet.Fields.Item("InvntryUom").Value);
                    Sbuilder.Append("</Value></Cell>");

                    Sbuilder.Append("<Cell> <ColumnUid>BarCode</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, barCode);
                    Sbuilder.Append("</Value></Cell>");

                    Sbuilder.Append("<Cell> <ColumnUid>Quantity</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, FormsB1.ConvertDecimalToString(Convert.ToDecimal(oRecordSet.Fields.Item("Quantity").Value)));
                    Sbuilder.Append("</Value></Cell>");

                    Sbuilder.Append("<Cell> <ColumnUid>RSQuantity</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, FormsB1.ConvertDecimalToString(Convert.ToDecimal(RSQuantity)));
                    Sbuilder.Append("</Value></Cell>");

                    Sbuilder.Append("<Cell> <ColumnUid>RSName</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, RS_W_NAME);
                    Sbuilder.Append("</Value></Cell>");
                }
                else
                {
                    Sbuilder.Append("<Cell> <ColumnUid>RSAmount</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, FormsB1.ConvertDecimalToString(Convert.ToDecimal(RSAmount_Full)));
                    Sbuilder.Append("</Value></Cell>");

                    Sbuilder.Append("<Cell> <ColumnUid>Gtotal</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, FormsB1.ConvertDecimalToString(Convert.ToDecimal(Amount)));
                    Sbuilder.Append("</Value></Cell>");
                }

                Sbuilder.Append("<Cell> <ColumnUid>BaseType</ColumnUid> <Value>");
                Sbuilder = CommonFunctions.AppendXML(Sbuilder, oRecordSet.Fields.Item("BaseType").Value);
                Sbuilder.Append("</Value></Cell>");

                Sbuilder.Append("</Row>");

                count++;

                oRecordSet.MoveNext();
            }


            if (RSDataTable.Rows.Count > 0 & (FilterCompStat == "1" || FilterCompStat == "0" || FilterCompStat == ""))
            {
                DataRow[] RemainingRows;
                RemainingRows = RSDataTable.Select("RowLinked = 'N'");

                DateTime DeliveryDate;
                DateTime BegDate;
                DateTime ActivateDate;
                double TranspCost;

                string cardName;
                string WAYBILL_NUMBER;
                string cTIN;
                string cCode;
                string strBEGIN_DATE;
                string wbSTATUS;
                string wbSTART_ADDRESS;
                string wbEND_ADDRESS;
                string strDELIVERY_DATE;
                string wbID;
                string strACTIVATE_DATE;
                string wbCAR_NUMBER;
                string wbDRIVER_TIN;
                string wbTRANSPORT_COAST;
                string BAR_CODE_RS;
                string W_NAME;
                double RSFULL_AMOUNT = 0;

                for (int i = 0; i < RemainingRows.Length; i++)
                {
                    Sbuilder.Append("<Row>");

                    cardName = "";
                    WAYBILL_NUMBER = "";
                    cTIN = "";
                    cCode = "";

                    cTIN = RemainingRows[i]["TIN"].ToString().Trim();

                    if (cTIN != "")
                    {
                        cCode = BusinessPartners.GetCardCodeByTin( cTIN, "C", out cardName);
                    }

                    //if (String.IsNullOrWhiteSpace(cCode) == false)
                    //{
                        Sbuilder.Append("<Cell> <ColumnUid>BaseCard</ColumnUid> <Value>");
                        Sbuilder = CommonFunctions.AppendXML(Sbuilder, cCode);
                        Sbuilder.Append("</Value></Cell>");
                    //}

                    Sbuilder.Append("<Cell> <ColumnUid>ComparStat</ColumnUid> <Value>1</Value></Cell>");

                    WAYBILL_NUMBER = RemainingRows[i]["WAYBILL_NUMBER"].ToString();
                    Sbuilder.Append("<Cell> <ColumnUid>WB_number</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, WAYBILL_NUMBER);
                    Sbuilder.Append("</Value></Cell>");

                    Sbuilder.Append("<Cell> <ColumnUid>LineNum</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, (count + 1).ToString());
                    Sbuilder.Append("</Value></Cell>");

                    Sbuilder.Append("<Cell> <ColumnUid>LicTradNum</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, cTIN);
                    Sbuilder.Append("</Value></Cell>");

                    cardName = RemainingRows[i]["NAME"].ToString();
                    Sbuilder.Append("<Cell> <ColumnUid>CardName</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, cardName);
                    Sbuilder.Append("</Value></Cell>");

                    strBEGIN_DATE = RemainingRows[i]["BEGIN_DATE"].ToString();
                    if (strBEGIN_DATE != "")
                    {
                        Sbuilder.Append("<Cell> <ColumnUid>WB_begDate</ColumnUid> <Value>");
                        Sbuilder = CommonFunctions.AppendXML(Sbuilder, (DateTime.TryParse(strBEGIN_DATE, out BegDate) == false ? DateTime.MinValue : BegDate).ToString("yyyyMMdd"));
                        Sbuilder.Append("</Value></Cell>");
                    }

                    wbtype = RemainingRows[i]["TYPE"].ToString();
                    Sbuilder.Append("<Cell> <ColumnUid>TYPE</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, wbtype);
                    Sbuilder.Append("</Value></Cell>");

                    wbSTATUS = RemainingRows[i]["STATUS"].ToString();
                    Sbuilder.Append("<Cell> <ColumnUid>WB_status</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, wbSTATUS);
                    Sbuilder.Append("</Value></Cell>");

                    wbSTART_ADDRESS = RemainingRows[i]["START_ADDRESS"].ToString();
                    Sbuilder.Append("<Cell> <ColumnUid>WB_strAddrs</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, wbSTART_ADDRESS);
                    Sbuilder.Append("</Value></Cell>");

                    wbEND_ADDRESS = RemainingRows[i]["END_ADDRESS"].ToString();
                    Sbuilder.Append("<Cell> <ColumnUid>WB_endAddrs</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, wbEND_ADDRESS);
                    Sbuilder.Append("</Value></Cell>");

                    strDELIVERY_DATE = RemainingRows[i]["DELIVERY_DATE"].ToString();
                    if (strDELIVERY_DATE != "")
                    {
                        Sbuilder.Append("<Cell> <ColumnUid>WB_delvDate</ColumnUid> <Value>");
                        Sbuilder = CommonFunctions.AppendXML(Sbuilder, (DateTime.TryParse(strDELIVERY_DATE, out DeliveryDate) == false ? DateTime.MinValue : DeliveryDate).ToString("yyyyMMdd"));
                        Sbuilder.Append("</Value></Cell>");
                    }

                    wbID = RemainingRows[i]["ID"].ToString();
                    Sbuilder.Append("<Cell> <ColumnUid>WB_ID</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, wbID);
                    Sbuilder.Append("</Value></Cell>");

                    strACTIVATE_DATE = RemainingRows[i]["ACTIVATE_DATE"].ToString();
                    if (strACTIVATE_DATE != "")
                    {
                        Sbuilder.Append("<Cell> <ColumnUid>WB_actDate</ColumnUid> <Value>");
                        Sbuilder = CommonFunctions.AppendXML(Sbuilder, (DateTime.TryParse(strACTIVATE_DATE, out ActivateDate) == false ? DateTime.MinValue : ActivateDate).ToString("yyyyMMdd"));
                        Sbuilder.Append("</Value></Cell>");
                    }

                    wbCAR_NUMBER = RemainingRows[i]["CAR_NUMBER"].ToString();
                    Sbuilder.Append("<Cell> <ColumnUid>WB_vehicNum</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, wbCAR_NUMBER);
                    Sbuilder.Append("</Value></Cell>");

                    wbDRIVER_TIN = RemainingRows[i]["DRIVER_TIN"].ToString();
                    Sbuilder.Append("<Cell> <ColumnUid>WB_drivTin</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, wbDRIVER_TIN);
                    Sbuilder.Append("</Value></Cell>");

                    wbTRANSPORT_COAST = RemainingRows[i]["TRANSPORT_COAST"].ToString();
                    Sbuilder.Append("<Cell> <ColumnUid>WB_trnsExpn</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, FormsB1.ConvertDecimalToString(Convert.ToDecimal(Double.TryParse(wbTRANSPORT_COAST, out TranspCost) == false ? 0 : TranspCost)));
                    Sbuilder.Append("</Value></Cell>");

                    Sbuilder.Append("<Cell> <ColumnUid>RS_Status</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, RemainingRows[i]["STATUS"].ToString());
                    Sbuilder.Append("</Value></Cell>");

                    //დეტალური როცა არის
                    if (FOption != "1")
                    {
                        RSAmount = Convert.ToDouble(RemainingRows[i]["AMOUNT"].ToString(), System.Globalization.CultureInfo.InvariantCulture);
                        Sbuilder.Append("<Cell> <ColumnUid>RSAmount</ColumnUid> <Value>");
                        Sbuilder = CommonFunctions.AppendXML(Sbuilder, FormsB1.ConvertDecimalToString(Convert.ToDecimal(wbtype == "5" ? RSAmount * (-1) : RSAmount)));
                        Sbuilder.Append("</Value></Cell>");

                        BAR_CODE_RS = RemainingRows[i]["BAR_CODE"].ToString();
                        Sbuilder.Append("<Cell> <ColumnUid>BarCode</ColumnUid> <Value>");
                        Sbuilder = CommonFunctions.AppendXML(Sbuilder, BAR_CODE_RS);
                        Sbuilder.Append("</Value></Cell>");

                        RSQuantity = Convert.ToDouble(RemainingRows[i]["Quantity"].ToString(), System.Globalization.CultureInfo.InvariantCulture);
                        Sbuilder.Append("<Cell> <ColumnUid>RSQuantity</ColumnUid> <Value>");
                        Sbuilder = CommonFunctions.AppendXML(Sbuilder, FormsB1.ConvertDecimalToString(Convert.ToDecimal(wbtype == "5" ? RSQuantity * (-1) : RSQuantity)));
                        Sbuilder.Append("</Value></Cell>");

                        W_NAME = RemainingRows[i]["W_NAME"].ToString();
                        Sbuilder.Append("<Cell> <ColumnUid>RSName</ColumnUid> <Value>");
                        Sbuilder = CommonFunctions.AppendXML(Sbuilder, W_NAME);
                        Sbuilder.Append("</Value></Cell>");

                    }
                    else
                    {
                        RSFULL_AMOUNT = Convert.ToDouble(RemainingRows[i]["FULL_AMOUNT"].ToString(), System.Globalization.CultureInfo.InvariantCulture);
                        Sbuilder.Append("<Cell> <ColumnUid>RSAmount</ColumnUid> <Value>");
                        Sbuilder = CommonFunctions.AppendXML(Sbuilder, FormsB1.ConvertDecimalToString(Convert.ToDecimal(wbtype == "5" ? RSFULL_AMOUNT * (-1) : RSFULL_AMOUNT)));
                        Sbuilder.Append("</Value></Cell>");
                    }

                    Sbuilder.Append("</Row>");
                    count++;
                }
            }

            Sbuilder.Append("</Rows>");
            Sbuilder.Append("</DataTable>");

            XML = Sbuilder.ToString();
            oDataTable.LoadFromXML(XML);

            SAPbouiCOM.Grid oGrid = ((SAPbouiCOM.Grid)(oForm.Items.Item("WbTable").Specific));
            SAPbouiCOM.GridColumns oColumns = oGrid.Columns;

            oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;

            oGrid.DataTable = oDataTable;
            oGrid.Columns.Item("BaseCard").TitleObject.Caption = BDOSResources.getTranslate("BP");
            oGrid.Columns.Item("ComparStat").TitleObject.Caption = BDOSResources.getTranslate("CompareStatus");
            oGrid.Columns.Item("WB_number").TitleObject.Caption = BDOSResources.getTranslate("WaybillNumber");

            oGrid.Columns.Item("LineNum").TitleObject.Caption = "#";
            oGrid.Columns.Item("LicTradNum").TitleObject.Caption = BDOSResources.getTranslate("Tin");
            oGrid.Columns.Item("CardName").TitleObject.Caption = BDOSResources.getTranslate("Name");
            oGrid.Columns.Item("WB_begDate").TitleObject.Caption = BDOSResources.getTranslate("TransBeginTime");
            oGrid.Columns.Item("TYPE").TitleObject.Caption = BDOSResources.getTranslate("Type");
            oGrid.Columns.Item("WB_status").TitleObject.Caption = BDOSResources.getTranslate("Status");
            oGrid.Columns.Item("WB_strAddrs").TitleObject.Caption = BDOSResources.getTranslate("StartAddress");
            oGrid.Columns.Item("WB_endAddrs").TitleObject.Caption = BDOSResources.getTranslate("EndAddress");
            oGrid.Columns.Item("WB_delvDate").TitleObject.Caption = BDOSResources.getTranslate("DeliveryDate");

            oGrid.Columns.Item("WB_ID").TitleObject.Caption = BDOSResources.getTranslate("WaybillID");
            oGrid.Columns.Item("WB_actDate").TitleObject.Caption = BDOSResources.getTranslate("ActivateDate");
            oGrid.Columns.Item("WB_vehicNum").TitleObject.Caption = BDOSResources.getTranslate("Vehicle");
            oGrid.Columns.Item("WB_drivTin").TitleObject.Caption = BDOSResources.getTranslate("TransporterTin");
            oGrid.Columns.Item("WB_trnsExpn").TitleObject.Caption = BDOSResources.getTranslate("TransportationExpense");
            oGrid.Columns.Item("WBLD_Doc").TitleObject.Caption = BDOSResources.getTranslate("WaybillDocEntry");

            oGrid.Columns.Item("DocDate").TitleObject.Caption = BDOSResources.getTranslate("Date");
            oGrid.Columns.Item("DocNum").TitleObject.Caption = BDOSResources.getTranslate("DocNum");
            oGrid.Columns.Item("BaseDoc").TitleObject.Caption = BDOSResources.getTranslate("BaseDocument");
            oGrid.Columns.Item("WhsTo").TitleObject.Caption = BDOSResources.getTranslate("ToWarehouse");
            oGrid.Columns.Item("WhsFrom").TitleObject.Caption = BDOSResources.getTranslate("FromWarehouse");

            oGrid.Columns.Item("RS_Status").TitleObject.Caption = BDOSResources.getTranslate("Status") + " RS";
            oGrid.Columns.Item("RSAmount").TitleObject.Caption = BDOSResources.getTranslate("Amount") + " RS";
            oGrid.Columns.Item("Gtotal").TitleObject.Caption = BDOSResources.getTranslate("Amount") + " " + BDOSResources.getTranslate("Document");

            oGrid.Columns.Item("ItemCode").TitleObject.Caption = BDOSResources.getTranslate("ItemCode");
            oGrid.Columns.Item("InvntryUom").TitleObject.Caption = BDOSResources.getTranslate("UomEntry");
            oGrid.Columns.Item("BarCode").TitleObject.Caption = BDOSResources.getTranslate("Code");
            oGrid.Columns.Item("Quantity").TitleObject.Caption = BDOSResources.getTranslate("Quantity");
            oGrid.Columns.Item("RSQuantity").TitleObject.Caption = BDOSResources.getTranslate("Quantity") + " RS";
            oGrid.Columns.Item("RSName").TitleObject.Caption = BDOSResources.getTranslate("Name") + " RS";

            //GTotal                 
            SAPbouiCOM.GridColumn oGC = oGrid.Columns.Item(25);
            oGC.Type = SAPbouiCOM.BoGridColumnType.gct_EditText;
            SAPbouiCOM.EditTextColumn oEditGC = (SAPbouiCOM.EditTextColumn)oGC;
            SAPbouiCOM.BoColumnSumType oST = oEditGC.ColumnSetting.SumType;
            oEditGC.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;

            //RSAmount                 
            SAPbouiCOM.GridColumn oAC = oGrid.Columns.Item(24);
            oGC.Type = SAPbouiCOM.BoGridColumnType.gct_EditText;
            SAPbouiCOM.EditTextColumn oEditAC = (SAPbouiCOM.EditTextColumn)oAC;
            SAPbouiCOM.BoColumnSumType oAST = oEditGC.ColumnSetting.SumType;
            oEditAC.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;

            //ComparStat
            oGrid.Columns.Item(1).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;

            SAPbouiCOM.ComboBoxColumn oComboComparStat = (SAPbouiCOM.ComboBoxColumn)oGrid.Columns.Item(1);
            oComboComparStat.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description;
            oComboComparStat.ValidValues.Add("1", BDOSResources.getTranslate("OnlyOnSite"));
            oComboComparStat.ValidValues.Add("2", BDOSResources.getTranslate("OnlyOnProgram"));
            oComboComparStat.ValidValues.Add("3", BDOSResources.getTranslate("AmountsNotEqual"));
            oComboComparStat.ValidValues.Add("4", BDOSResources.getTranslate("EqualAmounts"));
            oComboComparStat.ValidValues.Add("5", BDOSResources.getTranslate("Linked") + " " + BDOSResources.getTranslate("Document") + " " + BDOSResources.getTranslate("NotPosted"));
            oComboComparStat.ValidValues.Add("6", BDOSResources.getTranslate("SavedStatus"));
            oComboComparStat.ValidValues.Add("", "");

            //TYPE
            oGrid.Columns.Item(7).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;

            SAPbouiCOM.ComboBoxColumn oComboTYPE = (SAPbouiCOM.ComboBoxColumn)oGrid.Columns.Item(7);
            oComboTYPE.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description;
            oComboTYPE.ValidValues.Add("1", BDOSResources.getTranslate("InternalShipment"));
            oComboTYPE.ValidValues.Add("2", BDOSResources.getTranslate("WithTransport"));
            oComboTYPE.ValidValues.Add("3", BDOSResources.getTranslate("WithoutTransport"));
            oComboTYPE.ValidValues.Add("4", BDOSResources.getTranslate("Distribution"));
            oComboTYPE.ValidValues.Add("5", BDOSResources.getTranslate("Return"));
            oComboTYPE.ValidValues.Add("6", BDOSResources.getTranslate("SubWaybill"));
            oComboTYPE.ValidValues.Add("", "");

            //WB_Status
            oGrid.Columns.Item(8).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;

            SAPbouiCOM.ComboBoxColumn oWB_Status = (SAPbouiCOM.ComboBoxColumn)oGrid.Columns.Item(8);
            oWB_Status.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description;
            oWB_Status.ValidValues.Add("-2", BDOSResources.getTranslate("Canceled"));
            oWB_Status.ValidValues.Add("-1", BDOSResources.getTranslate("deleted"));
            oWB_Status.ValidValues.Add("0", BDOSResources.getTranslate("Saved"));
            oWB_Status.ValidValues.Add("1", BDOSResources.getTranslate("Active"));
            oWB_Status.ValidValues.Add("2", BDOSResources.getTranslate("finished"));
            oWB_Status.ValidValues.Add("8", BDOSResources.getTranslate("SentToTransporter"));
            oWB_Status.ValidValues.Add("", "");

            //RS_Status
            oGrid.Columns.Item(23).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;

            SAPbouiCOM.ComboBoxColumn oRS_Status = (SAPbouiCOM.ComboBoxColumn)oGrid.Columns.Item(23);
            oRS_Status.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description;
            oRS_Status.ValidValues.Add("-2", BDOSResources.getTranslate("Canceled"));
            oRS_Status.ValidValues.Add("-1", BDOSResources.getTranslate("deleted"));
            oRS_Status.ValidValues.Add("0", BDOSResources.getTranslate("Saved"));
            oRS_Status.ValidValues.Add("1", BDOSResources.getTranslate("Active"));
            oRS_Status.ValidValues.Add("2", BDOSResources.getTranslate("finished"));
            oRS_Status.ValidValues.Add("8", BDOSResources.getTranslate("SentToTransporter"));
            oRS_Status.ValidValues.Add("", "");

            //BaseCard
            SAPbouiCOM.EditTextColumn oBaseCard = (SAPbouiCOM.EditTextColumn)oGrid.Columns.Item("BaseCard");
            oBaseCard.LinkedObjectType = "2";

            //WhsTo
            SAPbouiCOM.EditTextColumn oWhsTo = (SAPbouiCOM.EditTextColumn)oGrid.Columns.Item("WhsTo");
            oWhsTo.LinkedObjectType = "64";

            //WhsFrom
            SAPbouiCOM.EditTextColumn oWhsFrom = (SAPbouiCOM.EditTextColumn)oGrid.Columns.Item("WhsFrom");
            oWhsFrom.LinkedObjectType = "64";

            //U_vehicNum
            SAPbouiCOM.EditTextColumn oU_vehicNum = (SAPbouiCOM.EditTextColumn)oGrid.Columns.Item("WB_vehicNum");
            oU_vehicNum.LinkedObjectType = "UDO_F_BDO_VECL_D";

            //U_baseDocT
            SAPbouiCOM.EditTextColumn obaseDocT = (SAPbouiCOM.EditTextColumn)oGrid.Columns.Item("BaseDoc");
            obaseDocT.LinkedObjectType = "13";

            //WBLD_Doc
            SAPbouiCOM.EditTextColumn oWBLD_Doc = (SAPbouiCOM.EditTextColumn)oGrid.Columns.Item("WBLD_Doc");
            oWBLD_Doc.LinkedObjectType = "UDO_F_BDO_WBLD_D";

            //ItemCode
            SAPbouiCOM.EditTextColumn oItemCode = (SAPbouiCOM.EditTextColumn)oGrid.Columns.Item("ItemCode");
            oItemCode.LinkedObjectType = "4";

            for (int i = 0; i < oColumns.Count; i++)
            {
                oColumns.Item(i).Editable = false;
            }

            oColumns.Item(3).Visible = false;
            if (FOption == "1")
            {
                oColumns.Item(26).Visible = false;
                oColumns.Item(27).Visible = false;
                oColumns.Item(28).Visible = false;
                oColumns.Item(29).Visible = false;
                oColumns.Item(30).Visible = false;
                oColumns.Item(31).Visible = false;
                oColumns.Item(32).Visible = false;
            }

            if (FOption == "1")
            {
                oGrid.CollapseLevel = 2;
            }
            else
            {
                oGrid.CollapseLevel = 3;
            }

            oGrid.AutoResizeColumns();

            SetGridColor(oForm, false, out errorText);

        }

        public static void SetGridColor(SAPbouiCOM.Form oForm, Boolean itemPressed, out string errorText)
        {
            errorText = "";

            string oSetGridColor = oForm.DataSources.UserDataSources.Item("StGrColor").ValueEx;
            if (oSetGridColor != "Y" && itemPressed == false)
            {
                return;
            }


            SAPbouiCOM.Grid oGrid = ((SAPbouiCOM.Grid)(oForm.Items.Item("WbTable").Specific));
            string FOption = oForm.DataSources.UserDataSources.Item("option").ValueEx;

            //Compare status color
            int lastComparStat = 0;
            Boolean isleaf;

            for (int i = 0; i < oGrid.Rows.Count; i++)
            {
                if (oSetGridColor != "Y" && itemPressed == true)
                {
                    oGrid.CommonSetting.SetCellFontColor(i + 1, 2, FormsB1.getLongIntRGB(0, 0, 0));
                }
                else
                {
                    if (lastComparStat != oGrid.Rows.GetParent(i))
                    {
                        isleaf = oGrid.Rows.IsLeaf(i);
                        if (isleaf != false)
                        {
                            int dTableRow = oGrid.GetDataTableRowIndex(i);
                            string CompStat = oGrid.DataTable.GetValue("ComparStat", dTableRow);

                            lastComparStat = oGrid.Rows.GetParent(i);
                            if (FOption != "1")
                            {
                                lastComparStat = oGrid.Rows.GetParent(lastComparStat);
                            }
                            if (lastComparStat < 0)
                            {
                                lastComparStat = i;
                            }

                            if (CompStat == "1") //ლურჯი - არსებობს მხოლოდ საიტზე
                            {
                                oGrid.CommonSetting.SetCellFontColor(lastComparStat + 1, 2, FormsB1.getLongIntRGB(0, 0, 255));
                            }
                            else if (CompStat == "2") // ნარინჯისფერი - არსებობს მხოლოდ პროგრამაში
                            {
                                oGrid.CommonSetting.SetCellFontColor(lastComparStat + 1, 2, FormsB1.getLongIntRGB(255, 127, 80));
                            }
                            else if (CompStat == "3") // წითელი - თანხები განსხვავებულია
                            {
                                oGrid.CommonSetting.SetCellFontColor(lastComparStat + 1, 2, FormsB1.getLongIntRGB(255, 0, 0));
                            }
                            else if (CompStat == "4") // მწვანე - თანხები შეესაბამება
                            {
                                oGrid.CommonSetting.SetCellFontColor(lastComparStat + 1, 2, FormsB1.getLongIntRGB(0, 128, 0));
                            }
                            else if (CompStat == "5") // ნაცრისფერი - მიბმული დოკ.გატარებული არ არის
                            {
                                oGrid.CommonSetting.SetCellFontColor(lastComparStat + 1, 2, FormsB1.getLongIntRGB(128, 128, 128));
                            }
                            else // შავი
                            {
                                oGrid.CommonSetting.SetCellFontColor(lastComparStat + 1, 2, FormsB1.getLongIntRGB(0, 0, 0));
                            }

                        }
                    }
                }
            }
        }

        public static void collapseGrid(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            SAPbouiCOM.Grid oGrid = ((SAPbouiCOM.Grid)(oForm.Items.Item("WbTable").Specific));
            oGrid.Rows.CollapseAll();
        }

        public static void expandGrid(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            SAPbouiCOM.Grid oGrid = ((SAPbouiCOM.Grid)(oForm.Items.Item("WbTable").Specific));
            oGrid.Rows.ExpandAll();
        }

        public static void gridColumnSetCfl( SAPbouiCOM.Form oForm, SAPbouiCOM.ItemEvent pVal, out string errorText)
        {
            errorText = null;

            try
            {
                if (pVal.ColUID == "BaseDoc")
                {
                    if ((pVal.EventType == SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED & pVal.BeforeAction == true) || (pVal.EventType == SAPbouiCOM.BoEventTypes.et_GOT_FOCUS & pVal.BeforeAction == false))
                    {
                        SAPbouiCOM.Grid oGrid = ((SAPbouiCOM.Grid)(oForm.Items.Item("WbTable").Specific));

                        int dTableRow = oGrid.GetDataTableRowIndex(pVal.Row);

                        SAPbouiCOM.DataTable oDataTable = (SAPbouiCOM.DataTable)oForm.DataSources.DataTables.Item("WbTable");
                        string DocType = oDataTable.GetValue("TYPE", dTableRow);
                        string BaseType = oDataTable.GetValue("BaseType", dTableRow);

                        //UbaseDocT
                        SAPbouiCOM.EditTextColumn obaseDocT = (SAPbouiCOM.EditTextColumn)oGrid.Columns.Item("BaseDoc");

                        if (BaseType == "60") //ჩამოწერა
                        {
                            obaseDocT.LinkedObjectType = "60";
                        }
                        else if (BaseType == "UDO_F_BDOSFASTRD_D") //Fixed Asset Transfer
                        {
                            obaseDocT.LinkedObjectType = "UDO_F_BDOSFASTRD_D";
                        }
                        else if (DocType == "1") //გადაადგილება
                        {
                            obaseDocT.LinkedObjectType = "67";
                        }
                        else if (DocType == "5") //დაბრუნება
                        {
                            obaseDocT.LinkedObjectType = "14";
                        }

                        else if (DocType == "165") // AR Correction Invoice
                        {
                            obaseDocT.LinkedObjectType = "165";
                        }

                        else if (BaseType == "15") //მიწოდება
                        {
                            obaseDocT.LinkedObjectType = "15";
                        }

                        else
                        {
                            obaseDocT.LinkedObjectType = "13";
                        }
                    }
                }
                else if (pVal.ColUID == "WhsFrom" || pVal.ColUID == "WhsTo")
                {
                    if ((pVal.EventType == SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED & pVal.BeforeAction == true) || (pVal.EventType == SAPbouiCOM.BoEventTypes.et_GOT_FOCUS & pVal.BeforeAction == false))
                    {
                        SAPbouiCOM.Grid oGrid = ((SAPbouiCOM.Grid)(oForm.Items.Item("WbTable").Specific));

                        int dTableRow = oGrid.GetDataTableRowIndex(pVal.Row);

                        SAPbouiCOM.DataTable oDataTable = (SAPbouiCOM.DataTable)oForm.DataSources.DataTables.Item("WbTable");
                        string BaseType = oDataTable.GetValue("BaseType", dTableRow);
                        
                        if (BaseType == "UDO_F_BDOSFASTRD_D") //Fixed Asset Transfer
                        {
                            SAPbouiCOM.EditTextColumn oWhsFrom = (SAPbouiCOM.EditTextColumn)oGrid.Columns.Item("WhsFrom");
                            SAPbouiCOM.EditTextColumn oWhsTo = (SAPbouiCOM.EditTextColumn)oGrid.Columns.Item("WhsTo");
                            oWhsFrom.LinkedObjectType = "144";
                            oWhsTo.LinkedObjectType = "144";
                        }
                        
                    }
                }
                else
                {

                }
            }
            catch (Exception ex)
            {
                int errCode;
                string errMsg;

                Program.oCompany.GetLastError(out errCode, out errMsg);
                errorText = BDOSResources.getTranslate("ErrorDescription") + " : " + errMsg + "! " + BDOSResources.getTranslate("Code") + " : " + errCode + "! " + BDOSResources.getTranslate("OtherInfo") + " : " + ex.Message;
            }
            finally
            {
                GC.Collect();
            }
        }


        public static void addMenus()
        {
            SAPbouiCOM.Menus moduleMenus;
            SAPbouiCOM.MenuItem menuItem;
            SAPbouiCOM.MenuCreationParams oCreationPackage;

            try
            {
                // Find the id of the menu into wich you want to add your menu item
                menuItem = Program.uiApp.Menus.Item("12800");

                // Get the menu collection of SAP Business One
                moduleMenus = menuItem.SubMenus;

                // Add a pop-up menu item
                oCreationPackage = (SAPbouiCOM.MenuCreationParams)Program.uiApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                oCreationPackage.Checked = false;
                oCreationPackage.Enabled = true;
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "BDO_WBSA";
                oCreationPackage.String = BDOSResources.getTranslate("SentWaybillsAnalysis");
                oCreationPackage.Position = -1;

                menuItem = moduleMenus.AddEx(oCreationPackage);
            }
            catch
            {
                
            }
        }

        public static string getQueryText(DateTime startDate, DateTime endDate, string cardCode, string itypes, string statuses, string foption)
        {
            string tempQuery = @"
         SELECT
	         ""@BDO_WBLD"".*,
	         ""BASEDOCS"".*,
	         ""@BDO_WBLD"".""DocEntry"" AS ""WBLD_Doc"",
	         (CASE 
            WHEN ""BaseType"" = '13' OR ""BaseType"" = '15'
	        THEN (CASE WHEN ""U_type"" = '0' 
		        THEN '2' 
		        ELSE '3' 
		        END) 
            
            WHEN ""BaseType"" = '67' OR ""BaseType"" = '60' OR ""BaseType"" = 'UDO_F_BDOSFASTRD_D'
                THEN '1' 
            WHEN ""BaseType"" = '14' 
	        THEN '5' 
WHEN ""BaseType"" = '165' THEN '165' 

            ELSE '5' 
	        END) AS ""TYPE"",
	         (CASE WHEN ""U_status"" = '1' 
	        THEN '0' WHEN ""U_status"" = '2' 
	        THEN '1' WHEN ""U_status"" = '3' 
	        THEN '2' WHEN ""U_status"" = '4' 
	        THEN '-1' WHEN ""U_status"" = '5' 
	        THEN '-2' 
	        ELSE '8' 
	        END) AS ""WB_Status"" 
        FROM 
	         (SELECT
	         ""OCRD"".""CardName"",
	         ""OCRD"".""LicTradNum"",
	         ""BASEDOCGDS"".""BaseType"", " +
          ((foption != "1") ? @"
	         ""BASEDOCGDS"".""LineNum"",
	         ""BASEDOCGDS"".""Dscription"" AS ItemDesc,
	         ""BASEDOCGDS"".""ItemCode"",
	         ""OITM"".""SWW"",
	         ""OITM"".""InvntryUom"",
	         ""OITM"".""CodeBars"", " : " ") + @"
	         ""BASEDOCGDS"".""BaseCard"",
	         ""BASEDOCGDS"".""DocEntry"",
	         ""BASEDOCGDS"".""WhsFrom"",
	         ""BASEDOCGDS"".""WhsTo"",
	         ""BASEDOCGDS"".""DocNum"" AS BaseDocNum,
	         ""BASEDOCGDS"".""DocDate"" AS BaseDocDate,
	         ""BASEDOCGDS"".""CANCELED"" AS CANCELED,
	         SUM(""BASEDOCGDS"".""Quantity"") AS Quantity,
	         MAX(""BASEDOCGDS"".""DocTotal"") AS DocTotal,
	         SUM(""BASEDOCGDS"".""GTotal"") AS GTotal 
	        FROM 
	         (SELECT
	         ""INV"".""BaseType"",
	         ""INV"".""Quantity"",
	         ""INV"".""GTotal"",
	         ""INV"".""LineNum"",
	         ""INV"".""Dscription"",
	         ""INV"".""ItemCode"",
	         ""INV"".""BaseCard"",
	         ""INV"".""DocEntry"",
	         ""INV"".""WhsTo"",
	         ""INV"".""WhsFrom"",
	         ""OINV"".""DocNum"",
	         ""OINV"".""DocDate"",
	         ""INV"".""DocTotal"",
	         ""OINV"".""CANCELED"" 
		        FROM (SELECT
	         '13' AS ""BaseType"",
	         ""INV1"".""Quantity"" * ""INV1"".""NumPerMsr"" AS ""Quantity"",
	         ""INV1"".""GTotal"",
             ""OINV"".""DocTotal"" + ""OINV"".""DpmAmnt"" AS ""DocTotal"",
	         ""INV1"".""LineNum"",
	         ""INV1"".""Dscription"",
	         ""INV1"".""ItemCode"",
	         ""INV1"".""BaseCard"",
	         ""INV1"".""DocEntry"",
	         ""INV1"".""WhsCode"" AS ""WhsFrom"",
	         NULL AS ""WhsTo"" 
			        FROM ""INV1"" 
                    INNER JOIN ""OINV"" ON ""OINV"".""DocEntry"" = ""INV1"".""DocEntry""
			    UNION ALL 
            SELECT '13',
	         ""RIN1"".""Quantity"" * (-1) * (CASE WHEN ""RIN1"".""NoInvtryMv"" = 'Y' 
				        THEN 0 
				        ELSE 1 
				        END) * ""RIN1"".""NumPerMsr"",
	         ""RIN1"".""GTotal"" * (-1),
             ""ORIN"".""DocTotal"" * (-1) + ""ORIN"".""DpmAmnt"" * (-1),
	         ""RIN1"".""BaseLine"",
	         ""RIN1"".""Dscription"",
	         ""RIN1"".""ItemCode"",
	         ""RIN1"".""BaseCard"",
	         ""RIN1"".""BaseEntry"",
	         ""RIN1"".""WhsCode"" AS ""WhsFrom"",
	         NULL AS ""WhsTo"" 
			        FROM ""RIN1"" 

			        INNER JOIN ""ORIN"" ON ""ORIN"".""DocEntry"" = ""RIN1"".""DocEntry"" 
			        WHERE ""RIN1"".""TargetType"" < 0 

			        AND ""ORIN"".""U_BDO_CNTp"" <> 1) AS ""INV"" 
		        INNER JOIN ""OINV"" ON ""INV"".""DocEntry"" = ""OINV"".""DocEntry"" 
		       
         UNION ALL
            SELECT
	         ""DEL"".""BaseType"",
	         ""DEL"".""Quantity"",
	         ""DEL"".""GTotal"",
	         ""DEL"".""LineNum"",
	         ""DEL"".""Dscription"",
	         ""DEL"".""ItemCode"",
	         ""DEL"".""BaseCard"",
	         ""DEL"".""DocEntry"",
	         ""DEL"".""WhsTo"",
	         ""DEL"".""WhsFrom"",
	         ""ODLN"".""DocNum"",
	         ""ODLN"".""DocDate"",
	         ""DEL"".""DocTotal"",
	         ""ODLN"".""CANCELED"" 
		        FROM (SELECT
	         '15' AS ""BaseType"",
	         ""DLN1"".""Quantity"" * ""DLN1"".""NumPerMsr"" AS ""Quantity"",
	         ""DLN1"".""GTotal"",
             ""ODLN"".""DocTotal"" + ""ODLN"".""DpmAmnt"" AS ""DocTotal"",
	         ""DLN1"".""LineNum"",
	         ""DLN1"".""Dscription"",
	         ""DLN1"".""ItemCode"",
	         ""DLN1"".""BaseCard"",
	         ""DLN1"".""DocEntry"",
	         ""DLN1"".""WhsCode"" AS ""WhsFrom"",
	         NULL AS ""WhsTo"" 
			        FROM ""DLN1"" 
                    INNER JOIN ""ODLN"" ON ""ODLN"".""DocEntry"" = ""DLN1"".""DocEntry""
			    UNION ALL 
            SELECT '15',
	         ""RIN1"".""Quantity"" * (-1) * (CASE WHEN ""RIN1"".""NoInvtryMv"" = 'Y' 
				        THEN 0 
				        ELSE 1 
				        END) * ""RIN1"".""NumPerMsr"",
	         ""RIN1"".""GTotal"" * (-1),
             ""ORIN"".""DocTotal"" * (-1) + ""ORIN"".""DpmAmnt"" * (-1),
	         ""RIN1"".""BaseLine"",
	         ""RIN1"".""Dscription"",
	         ""RIN1"".""ItemCode"",
	         ""RIN1"".""BaseCard"",
	         ""RIN1"".""BaseEntry"",
	         ""RIN1"".""WhsCode"" AS ""WhsFrom"",
	         NULL AS ""WhsTo"" 
			        FROM ""RIN1"" 

			        INNER JOIN ""ORIN"" ON ""ORIN"".""DocEntry"" = ""RIN1"".""DocEntry"" 
			        WHERE ""RIN1"".""TargetType"" < 0 

			        AND ""ORIN"".""U_BDO_CNTp"" <> 1) AS ""DEL"" 
		        INNER JOIN ""ODLN"" ON ""DEL"".""DocEntry"" = ""ODLN"".""DocEntry"" 



UNION ALL SELECT
	         '14',
	         ""RIN1"".""Quantity"" * (-1) * (CASE WHEN ""RIN1"".""NoInvtryMv"" = 'Y' 
			        THEN 0 
			        ELSE 1 
			        END) * ""RIN1"".""NumPerMsr"",
	         ""RIN1"".""GTotal"" * (-1),
	         ""RIN1"".""BaseLine"",
	         ""RIN1"".""Dscription"",
	         ""RIN1"".""ItemCode"",
	         ""RIN1"".""BaseCard"",
	         ""RIN1"".""DocEntry"",
	         NULL AS ""WhsFrom"",
	         ""RIN1"".""WhsCode"" AS ""WhsTo"",
	         ""ORIN"".""DocNum"",
	         ""ORIN"".""DocDate"",
	         ""ORIN"".""DocTotal"" * (-1) + ""ORIN"".""DpmAmnt"" * (-1),
	         ""ORIN"".""CANCELED"" 
		        FROM ""RIN1"" 
		        INNER JOIN ""ORIN"" ON ""ORIN"".""DocEntry"" = ""RIN1"".""DocEntry"" 
		        WHERE ""RIN1"".""TargetType"" < 0 
		        AND ""ORIN"".""U_BDO_CNTp"" = 1 
		        

UNION ALL 
                SELECT '165', 
                       ""CSI1"".""Quantity"" * ( -1 ) * ( CASE 
            WHEN
            ""CSI1"".""NoInvtryMv"" = 'Y'
            THEN 0
            ELSE 1
            END ) *
                ""CSI1"".""NumPerMsr"", 
            ""CSI1"".""GTotal"" * (-1), 
            ""CSI1"".""BaseLine"", 
            ""CSI1"".""Dscription"", 
            ""CSI1"".""ItemCode"", 
            ""CSI1"".""BaseCard"", 
            ""CSI1"".""DocEntry"", 
            NULL AS ""WhsFrom"", 
            ""CSI1"".""WhsCode"" AS ""WhsTo"", 
            ""OCSI"".""DocNum"", 
            ""OCSI"".""DocDate"", 
            ""OCSI"".""DocTotal"" * (-1) + ""OCSI"".""DpmAmnt"" * (-1), 
            ""OCSI"".""CANCELED""
            FROM   ""CSI1""
            INNER JOIN ""OCSI""
            ON ""OCSI"".""DocEntry"" = ""CSI1"".""DocEntry""
            WHERE  ""CSI1"".""TargetType"" < 0
            AND ""OCSI"".""U_BDOSCITp"" = 1



                UNION ALL SELECT
	         '67',
	         ""WTR1"".""Quantity"" * (CASE WHEN ""WTR1"".""NoInvtryMv"" = 'Y' 
			        THEN 0 
			        ELSE 1 
			        END) * ""WTR1"".""NumPerMsr"",
	         0,
	         ""WTR1"".""BaseLine"",
	         ""WTR1"".""Dscription"",
	         ""WTR1"".""ItemCode"",
	         ""WTR1"".""BaseCard"",
	         ""WTR1"".""DocEntry"",
	         ""WTR1"".""WhsCode"" AS ""WhsTo"",
	         ""WTR1"".""FromWhsCod"" AS ""WhsFrom"",
	         ""OWTR"".""DocNum"",
	         ""OWTR"".""DocDate"",
	         0,
	         ""OWTR"".""CANCELED"" 
		        FROM ""WTR1"" 
		        INNER JOIN ""OWTR"" ON ""OWTR"".""DocEntry"" = ""WTR1"".""DocEntry""

            

        
        UNION ALL SELECT
	         'UDO_F_BDOSFASTRD_D',

            (CASE WHEN ""@BDOSFASTR1"".""U_Quantity"" is null or ""@BDOSFASTR1"".""U_Quantity"" = '0' THEN '1' ELSE ""@BDOSFASTR1"".""U_Quantity"" END) * (CASE WHEN ""@BDOSFASTRD"".""Transfered"" = 'Y' 
			        THEN 0 
			        ELSE 1 
			        END),
	         0,
             0,
	         ""@BDOSFASTR1"".""U_ItemName"",
	         ""@BDOSFASTR1"".""U_ItemCode"",
	         ""@BDOSFASTRD"".""U_CardCode"",
	         ""@BDOSFASTR1"".""DocEntry"",
	         ""@BDOSFASTRD"".""U_TLocCode"" AS ""WhsTo"",
	         ""@BDOSFASTRD"".""U_FLocCode"" AS ""WhsFrom"",
	         ""@BDOSFASTRD"".""DocNum"",
	         ""@BDOSFASTRD"".""U_DocDate"",
	         0,
	         ""@BDOSFASTRD"".""Canceled"" 
		        FROM ""@BDOSFASTR1"" 
		        INNER JOIN ""@BDOSFASTRD"" ON ""@BDOSFASTRD"".""DocEntry"" = ""@BDOSFASTR1"".""DocEntry""

            


             UNION ALL SELECT   
		     '60',
	         ""IGE1"".""Quantity"" * (CASE WHEN ""IGE1"".""NoInvtryMv"" = 'Y' 
			        THEN 0 
			        ELSE 1 
			        END) * ""IGE1"".""NumPerMsr"",
	         0,
	         ""IGE1"".""BaseLine"",
	         ""IGE1"".""Dscription"",
	         ""IGE1"".""ItemCode"",
	         ""IGE1"".""BaseCard"",
	         ""IGE1"".""DocEntry"",
	         NULL AS ""WhsFrom"",
	         ""IGE1"".""WhsCode"" AS ""WhsTo"",
	         ""OIGE"".""DocNum"",
	         ""OIGE"".""DocDate"",
	         0,
	         ""OIGE"".""CANCELED"" 
		        FROM ""IGE1"" 
		        INNER JOIN ""OIGE"" ON ""OIGE"".""DocEntry"" = ""IGE1"".""DocEntry""


            ) AS ""BASEDOCGDS"" 
	        LEFT JOIN ""OCRD"" AS ""OCRD"" ON ""BASEDOCGDS"".""BaseCard"" = ""OCRD"".""CardCode"" 
	        LEFT JOIN ""OITM"" ON ""BASEDOCGDS"".""ItemCode"" = ""OITM"".""ItemCode"" 
	        WHERE ((""OITM"".""ItemType"" = 'I' AND ""OITM"".""InvntItem"" = 'Y') OR ""OITM"".""ItemType"" = 'F') " +
             ((startDate != new DateTime(1, 1, 1)) ? @" AND ""BASEDOCGDS"".""DocDate"" >= '" + startDate.ToString("yyyyMMdd") + "' " : " ") +
              ((endDate != new DateTime(1, 1, 1)) ? @" AND ""BASEDOCGDS"".""DocDate"" <= '" + endDate.ToString("yyyyMMdd") + "' " : " ") +
              ((cardCode != "") ? @" AND ""BASEDOCGDS"".""BaseCard"" = N'" + cardCode.Replace("'", "''") + "' " : " ") +
             @"GROUP BY ""OCRD"".""CardName"",
	         ""OCRD"".""LicTradNum"",
	         ""BASEDOCGDS"".""BaseType"",
	         ""BASEDOCGDS"".""BaseCard"",
	         ""BASEDOCGDS"".""DocEntry"",
	         ""BASEDOCGDS"".""WhsFrom"",
	         ""BASEDOCGDS"".""WhsTo"", " +
        ((foption != "1") ? @" 
	         ""BASEDOCGDS"".""LineNum"",
	         ""BASEDOCGDS"".""Dscription"",
	         ""BASEDOCGDS"".""ItemCode"",
	         ""OITM"".""SWW"",
	         ""OITM"".""InvntryUom"",
	         ""OITM"".""CodeBars""," : " ") + @" 
	         ""BASEDOCGDS"".""DocDate"",
	         ""BASEDOCGDS"".""DocNum"",
	         ""BASEDOCGDS"".""CANCELED""
	         ) AS ""BASEDOCS"" 
	       		    LEFT JOIN ""@BDO_WBLD"" ON (""BASEDOCS"".""DocEntry"" = ""@BDO_WBLD"".""U_baseDoc""
	        AND  ""BASEDOCS"".""BaseType"" = ""@BDO_WBLD"".""U_baseDocT"" AND ""@BDO_WBLD"".""Status"" != 'C') 
        WHERE 1 = 1 " +
       ((itypes != "" && itypes != "0") ? @" AND 
        (CASE 
            WHEN ""BaseType"" = '13' 
	        THEN (CASE WHEN ""U_type"" = '0' 
		        THEN '2' 
		        ELSE '3' 
		        END)
            WHEN ""BaseType"" = '15' 
	        THEN (CASE WHEN ""U_type"" = '0' 
		        THEN '2' 
		        ELSE '3' 
		        END)
            WHEN ""BaseType"" = '67' OR ""BaseType"" = '60' OR ""BaseType"" = 'UDO_F_BDOSFASTRD_D'
	        THEN '1' 
            WHEN ""BaseType"" = '14' 
	        THEN '5' 
WHEN ""BaseType"" = '165' THEN '165'
	        ELSE '5' 
	        END) ='" + itypes + @"'  OR ""BaseType"" = '165' " : " ") +
      ((statuses != "" && statuses != "-99") ? @"AND (CASE WHEN ""U_status"" = '1' 
	    THEN '0' WHEN ""U_status""= '2' 
	    THEN '1' WHEN ""U_status"" = '3' 
	    THEN '2' WHEN ""U_status"" = '4' 
	    THEN '-1' WHEN ""U_status"" = '5' 
	    THEN '-2' 
	    ELSE '8' 
	    END) = '" + statuses + "' " : " ");

            return tempQuery;
        }

        public static void reArrangeFormItems(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.Item oItem = null;

            oItem = oForm.Items.Item("Option");
            oItem.Left = 395;
        }

        public static void resizeForm( SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            try
            {
                reArrangeFormItems(oForm);
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
        }
        public static void chooseFromList( SAPbouiCOM.Form oForm, SAPbouiCOM.IChooseFromListEvent oCFLEvento, string itemUID, bool beforeAction, out string errorText)
        {
            errorText = null;

            try
            {
                string sCFL_ID = oCFLEvento.ChooseFromListUID;
                SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(sCFL_ID);

                if (beforeAction == false)
                {
                    SAPbouiCOM.DataTable oDataTable = null;
                    oDataTable = oCFLEvento.SelectedObjects;

                    if (oDataTable != null)
                    {
                        if (sCFL_ID == "BusinessPartner_CFL")
                        {
                            string businessPartnerCode = Convert.ToString(oDataTable.GetValue("CardCode", 0));
                            oForm.DataSources.UserDataSources.Item("ClientID").Value = businessPartnerCode;

                        }
                    }
                }
            }
            catch
            { }

        }

        public static void uiApp_ItemEvent(  string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if ((pVal.ItemUID == "ClientID") & pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                {
                    SAPbouiCOM.IChooseFromListEvent oCFLEvento = null;
                    oCFLEvento = ((SAPbouiCOM.IChooseFromListEvent)(pVal));

                    chooseFromList( oForm, oCFLEvento, pVal.ItemUID, pVal.BeforeAction, out errorText);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK & pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "WbFillTb")
                    {
                        oForm.Freeze(true);
                        updateGrid(  oForm, out errorText);
                        oForm.Update();
                        oForm.Freeze(false);
                    }
                    if (pVal.ItemUID == "btnColl")
                    {
                        collapseGrid(oForm, out errorText);
                    }
                    if (pVal.ItemUID == "btnExp")
                    {
                        expandGrid(oForm, out errorText);
                    }
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_COMBO_SELECT & pVal.BeforeAction == false)
                {
                    oForm.DataSources.UserDataSources.Item(pVal.ItemUID).ValueEx = oForm.Items.Item(pVal.ItemUID).Specific.Value;
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED)
                {
                    gridColumnSetCfl( oForm, pVal, out errorText);
                }

                if ((pVal.ItemUID == "StGrColor") & pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED & pVal.BeforeAction == false)
                {
                    SetGridColor(oForm, true, out errorText);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_RESIZE & pVal.BeforeAction == false)
                {
                    oForm.Freeze(true);
                    resizeForm( oForm, out errorText);
                    oForm.Freeze(false);
                }
            }
        }
    }
}
