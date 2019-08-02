﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Globalization;
using System.Text.RegularExpressions;

namespace BDO_Localisation_AddOn
{
    static partial class BDOSWaybillsAnalysisReceived
    {
        public static void createForm(  out string errorText)
        {
            errorText = null;

            //ფორმის აუცილებელი თვისებები
            Dictionary<string, object> formProperties = new Dictionary<string, object>();
            formProperties.Add("UniqueID", "BDOSWBRAn");
            formProperties.Add("BorderStyle", SAPbouiCOM.BoFormBorderStyle.fbs_Sizable);
            formProperties.Add("Title", BDOSResources.getTranslate("ReceivedWaybillsAnalysis"));
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
                    oDataTable.Columns.Add("Whs", SAPbouiCOM.BoFieldsType.ft_Text, 50);//21
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
                    oDataTable.Columns.Add("BaseDType", SAPbouiCOM.BoFieldsType.ft_Text, 50);//32----


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
                    oCon.CondVal = "S"; //მომწოდებელი
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
                    listValidValuesDict.Add("7", BDOSResources.getTranslate("NotFoundOnSiteOnThisPeriod"));

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

            string Fstatus = oForm.DataSources.UserDataSources.Item("WBStatus").ValueEx;
            string FOption = oForm.DataSources.UserDataSources.Item("option").ValueEx;

            string query = getQueryText(BeginDate, EndDate, cardCode, itypes, Fstatus, FOption);


            if (itypes == "0" || itypes == "")
            {
                itypes = "1,2,3,4,5,6";
            }

            string statuses;
            if (Fstatus == "-99" || Fstatus == "")
            {
                statuses = ",,1,2,-1,-2,7,";
            }
            else if (Fstatus == "0")
            {
                statuses = ",,";
            }
            else
            {
                statuses = "," + Fstatus + ",";
            }

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

            RSDataTable.Columns.Add("ItemCode", typeof(string));
            RSDataTable.Columns.Add("W_NAME", typeof(string));
            RSDataTable.Columns.Add("BAR_CODE", typeof(string));
            RSDataTable.Columns.Add("AMOUNT", typeof(string));
            RSDataTable.Columns.Add("Quantity", typeof(string));


            string car_number = "";
            DateTime begin_date_s = startDate;
            DateTime begin_date_e = endDate;
            DateTime create_date_s = startDate;
            DateTime create_date_e = endDate;
            string driver_tin = null;
            DateTime delivery_date_s = startDate;
            DateTime delivery_date_e = endDate;
            decimal full_amount = 0;
            string waybill_number = "";
            DateTime close_date_s = startDate;
            DateTime close_date_e = endDate;
            string s_user_id = "";
            string comment = null;
            string seller_id = buyer_tin;
            string StartAddress = "";
            string EndAddress = "";

            DateTime startDateParam = new DateTime();
            DateTime endDateParam = new DateTime();
            startDateParam = startDate;

            while (startDateParam < EndDateForWS)
            {
                endDateParam = startDateParam.AddDays(2);

                if (endDateParam > EndDateForWS)
                {
                    endDateParam = EndDateForWS;
                }

                Dictionary<string, Dictionary<string, string>> waybills_map_part = oWayBill.get_buyer_waybills(itypes, seller_id, ",,1,2,-1,-2,7,", car_number, startDateParam, endDateParam, startDateParam, endDateParam, driver_tin, startDateParam, endDateParam, full_amount, waybill_number, startDateParam, endDateParam, s_user_id, comment, StartAddress, EndAddress, out errorText);
                foreach (KeyValuePair<string, Dictionary<string, string>> map_record in waybills_map_part)
                {
                    Dictionary<string, string> Waybill_Header = map_record.Value;

                    string SELLER_TIN = Waybill_Header["SELLER_TIN"];
                    string SELLER_NAME = Waybill_Header["SELLER_NAME"];
                    string WBID = Waybill_Header["ID"];

                    if (FOption != "0")
                    {
                        DataRow taxDataRow = RSDataTable.Rows.Add();
                        taxDataRow["RowLinked"] = "N";
                        taxDataRow["ID"] = WBID;
                        taxDataRow["WAYBILL_NUMBER"] = Waybill_Header["WAYBILL_NUMBER"];
                        taxDataRow["FULL_AMOUNT"] = Waybill_Header["FULL_AMOUNT"];
                        taxDataRow["STATUS"] = Waybill_Header["STATUS"];
                        taxDataRow["TIN"] = SELLER_TIN;
                        taxDataRow["NAME"] = SELLER_NAME;
                        taxDataRow["BEGIN_DATE"] = Waybill_Header["BEGIN_DATE"];
                        taxDataRow["TYPE"] = Waybill_Header["TYPE"];
                        taxDataRow["START_ADDRESS"] = Waybill_Header["START_ADDRESS"];
                        taxDataRow["END_ADDRESS"] = Waybill_Header["END_ADDRESS"];
                        taxDataRow["DELIVERY_DATE"] = Waybill_Header["DELIVERY_DATE"];
                        taxDataRow["ACTIVATE_DATE"] = Waybill_Header["ACTIVATE_DATE"];
                        taxDataRow["CAR_NUMBER"] = Waybill_Header["CAR_NUMBER"];
                        taxDataRow["DRIVER_TIN"] = Waybill_Header["DRIVER_TIN"];
                        taxDataRow["TRANSPORT_COAST"] = Waybill_Header["TRANSPORT_COAST"];
                    }
                    else
                    {
                        //ცხრილი
                        string[] array_HEADER;
                        string[][] array_GOODS, array_SUB_WAYBILLS;
                        int returnCode = oWayBill.get_waybill(Convert.ToInt32(WBID), out array_HEADER, out array_GOODS, out array_SUB_WAYBILLS, out errorText);

                        int rowCounter = 1;
                        int rowIndex = 0;

                        foreach (string[] goodsRow in array_GOODS)
                        {
                            string WBBarcode = goodsRow[6] == null ? "" : Regex.Replace(goodsRow[6], @"\t|\n|\r|'", "").Trim();
                            string WBItmName = goodsRow[1];

                            string ItmCode = "";
                            string cardName;
                            string Cardcode = BusinessPartners.GetCardCodeByTin( SELLER_TIN, "S", out cardName);
                            if (Cardcode != null)
                            {
                                ItmCode = BDO_WaybillsJournalReceived.findItemByNameOITM( WBItmName, WBBarcode, Cardcode, out errorText);

                                SAPbobsCOM.Recordset CatalogEntry = BDO_BPCatalog.getCatalogEntryByBPBarcode(Cardcode, WBItmName, WBBarcode, out errorText);

                                if (CatalogEntry != null)
                                {
                                    ItmCode = CatalogEntry.Fields.Item("ItemCode").Value;
                                }
                            }

                            DataRow taxDataRow = RSDataTable.Rows.Add();
                            taxDataRow["RowLinked"] = "N";
                            taxDataRow["ID"] = Waybill_Header["ID"];
                            taxDataRow["WAYBILL_NUMBER"] = Waybill_Header["WAYBILL_NUMBER"];
                            taxDataRow["FULL_AMOUNT"] = Waybill_Header["FULL_AMOUNT"];
                            taxDataRow["STATUS"] = Waybill_Header["STATUS"];
                            taxDataRow["TIN"] = Waybill_Header["SELLER_TIN"];
                            taxDataRow["NAME"] = Waybill_Header["SELLER_NAME"];
                            taxDataRow["BEGIN_DATE"] = Waybill_Header["BEGIN_DATE"];
                            taxDataRow["TYPE"] = Waybill_Header["TYPE"];
                            taxDataRow["START_ADDRESS"] = Waybill_Header["START_ADDRESS"];
                            taxDataRow["END_ADDRESS"] = Waybill_Header["END_ADDRESS"];
                            taxDataRow["DELIVERY_DATE"] = Waybill_Header["DELIVERY_DATE"];
                            taxDataRow["ACTIVATE_DATE"] = Waybill_Header["ACTIVATE_DATE"];
                            taxDataRow["CAR_NUMBER"] = Waybill_Header["CAR_NUMBER"];
                            taxDataRow["DRIVER_TIN"] = Waybill_Header["DRIVER_TIN"];
                            taxDataRow["TRANSPORT_COAST"] = Waybill_Header["TRANSPORT_COAST"];

                            taxDataRow["ItemCode"] = ItmCode;
                            taxDataRow["W_NAME"] = WBItmName;
                            taxDataRow["BAR_CODE"] = WBBarcode;
                            taxDataRow["AMOUNT"] = goodsRow[5];
                            taxDataRow["Quantity"] = goodsRow[3];

                            rowCounter++;
                            rowIndex++;
                        }
                    }
                }
                startDateParam = endDateParam;
            }

            int count = 0;

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRecordSet.DoQuery(query);

            SAPbouiCOM.DataTable oDataTable;

            oDataTable = oForm.DataSources.DataTables.Item("WbTable");
            oDataTable.Rows.Clear();

            string XML = "";
            XML = oDataTable.GetAsXML();
            XML = XML.Replace("<Rows/></DataTable>", "");

            StringBuilder Sbuilder = new StringBuilder();
            Sbuilder.Append(XML);
            Sbuilder.Append("<Rows>");

            string FilterCompStat = oForm.DataSources.UserDataSources.Item("CompStat").ValueEx;

            string WBNum;
            int BaseDocNum;
            string BaseDType;
            int DocEntry;
            string ItemCode;
            double Amount;
            double Amount_Full;
            double RSAmount;
            double RSAmount_Full;
            double RSQuantity;
            string RS_W_NAME;
            string wbCompStat;
            string RS_STATUS;
            string TYPE;
            string START_ADDRESS;
            string END_ADDRESS;
            string CAR_NUMBER;
            string DRIVER_TIN;
            double TRANSPORT_COAST = 0;
            bool foundonRS;
            DateTime DeliveryDate;
            DateTime BegDate;
            DateTime ActivateDate;
            string strDELIVERY_DATE;
            string strACTIVATE_DATE;
            string strBEGIN_DATE;

            while (!oRecordSet.EoF)
            {
                WBNum = oRecordSet.Fields.Item("WBNo").Value;
                WBNum = WBNum.Trim();

                ItemCode = "";
                if (FOption != "1")
                {
                    ItemCode = oRecordSet.Fields.Item("ItemCode").Value;
                }

                Amount = oRecordSet.Fields.Item("Gtotal").Value;
                Amount_Full = oRecordSet.Fields.Item("DocTotal").Value;
                RSAmount = 0;
                RSAmount_Full = 0;
                RSQuantity = 0;
                RS_W_NAME = "";
                wbCompStat = "";
                RS_STATUS = "";
                TYPE = "";
                START_ADDRESS = "";
                END_ADDRESS = "";
                CAR_NUMBER = "";
                DRIVER_TIN = "";
                TRANSPORT_COAST = 0;
                strDELIVERY_DATE = "";
                strACTIVATE_DATE = "";
                strBEGIN_DATE = "";

                if (WBNum != "")
                {
                    foundonRS = false;

                    if (RSDataTable.Rows.Count > 0)
                    {
                        if (RSDataTable.Columns.Contains("WAYBILL_NUMBER"))
                        {
                            DataRow[] foundRows;
                            if (FOption != "1")
                            {
                                foundRows = RSDataTable.Select("WAYBILL_NUMBER = '" + WBNum + "'" + " and " + "ItemCode = '" + ItemCode.Replace("'", "''") + "'");
                            }
                            else
                            {
                                foundRows = RSDataTable.Select("WAYBILL_NUMBER = '" + WBNum + "'");
                            }

                            for (int i = 0; i < foundRows.Length; i++)
                            {
                                foundonRS = true;
                                foundRows[i]["RowLinked"] = "Y";

                                if (FOption != "1")
                                {
                                    RSAmount = Convert.ToDouble(foundRows[i]["AMOUNT"], System.Globalization.CultureInfo.InvariantCulture);
                                    RSQuantity = Convert.ToDouble(foundRows[i]["Quantity"], System.Globalization.CultureInfo.InvariantCulture);
                                    RS_W_NAME = foundRows[i]["W_NAME"].ToString();
                                }

                                RSAmount_Full = Convert.ToDouble(foundRows[i]["FULL_AMOUNT"], System.Globalization.CultureInfo.InvariantCulture);
                                RS_STATUS = foundRows[i]["STATUS"].ToString();
                                TYPE = foundRows[i]["TYPE"].ToString();
                                START_ADDRESS = foundRows[i]["START_ADDRESS"].ToString();
                                END_ADDRESS = foundRows[i]["END_ADDRESS"].ToString();
                                CAR_NUMBER = foundRows[i]["CAR_NUMBER"].ToString();
                                DRIVER_TIN = foundRows[i]["DRIVER_TIN"].ToString();

                                if (String.IsNullOrEmpty((string)foundRows[i]["TRANSPORT_COAST"]))
                                {
                                    TRANSPORT_COAST = 0;
                                }
                                else
                                {
                                TRANSPORT_COAST = Convert.ToDouble(foundRows[i]["TRANSPORT_COAST"], System.Globalization.CultureInfo.InvariantCulture);
                                }

                                strDELIVERY_DATE = foundRows[i]["DELIVERY_DATE"].ToString();
                                strACTIVATE_DATE = foundRows[i]["ACTIVATE_DATE"].ToString();
                                strBEGIN_DATE = foundRows[i]["BEGIN_DATE"].ToString();
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
                    else if (Amount_Full != RSAmount_Full)
                    {
                        wbCompStat = "3";
                    }
                    else
                    {
                        wbCompStat = "4";
                    }
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

                //სტატუსის ფილტრი
                if (((Fstatus == "-99" || Fstatus == "") || statuses.IndexOf(RS_STATUS) > 0) == false)
                {
                    oRecordSet.MoveNext();
                    continue;
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

                if (strBEGIN_DATE != "")
                {
                    Sbuilder.Append("<Cell> <ColumnUid>WB_begDate</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, (DateTime.TryParse(strBEGIN_DATE, out BegDate) == false ? DateTime.MinValue : BegDate).ToString("yyyyMMdd"));
                    Sbuilder.Append("</Value></Cell>");
                }

                Sbuilder.Append("<Cell> <ColumnUid>TYPE</ColumnUid> <Value>");
                Sbuilder = CommonFunctions.AppendXML(Sbuilder, TYPE);
                Sbuilder.Append("</Value></Cell>");

                Sbuilder.Append("<Cell> <ColumnUid>WB_strAddrs</ColumnUid> <Value>");
                Sbuilder = CommonFunctions.AppendXML(Sbuilder, START_ADDRESS);
                Sbuilder.Append("</Value></Cell>");

                Sbuilder.Append("<Cell> <ColumnUid>WB_endAddrs</ColumnUid> <Value>");
                Sbuilder = CommonFunctions.AppendXML(Sbuilder, END_ADDRESS);
                Sbuilder.Append("</Value></Cell>");

                if (strDELIVERY_DATE != "")
                {
                    Sbuilder.Append("<Cell> <ColumnUid>WB_delvDate</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, (DateTime.TryParse(strDELIVERY_DATE, out DeliveryDate) == false ? DateTime.MinValue : DeliveryDate).ToString("yyyyMMdd"));
                    Sbuilder.Append("</Value></Cell>");
                }

                Sbuilder.Append("<Cell> <ColumnUid>WB_ID</ColumnUid> <Value>");
                Sbuilder = CommonFunctions.AppendXML(Sbuilder, oRecordSet.Fields.Item("WBID").Value);
                Sbuilder.Append("</Value></Cell>");

                if (strACTIVATE_DATE != "")
                {
                    Sbuilder.Append("<Cell> <ColumnUid>WB_actDate</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, (DateTime.TryParse(strACTIVATE_DATE, out ActivateDate) == false ? DateTime.MinValue : ActivateDate).ToString("yyyyMMdd"));
                    Sbuilder.Append("</Value></Cell>");
                }

                Sbuilder.Append("<Cell> <ColumnUid>WB_vehicNum</ColumnUid> <Value>");
                Sbuilder = CommonFunctions.AppendXML(Sbuilder, CAR_NUMBER);
                Sbuilder.Append("</Value></Cell>");

                Sbuilder.Append("<Cell> <ColumnUid>WB_drivTin</ColumnUid> <Value>");
                Sbuilder = CommonFunctions.AppendXML(Sbuilder, DRIVER_TIN);
                Sbuilder.Append("</Value></Cell>");

                Sbuilder.Append("<Cell> <ColumnUid>WB_trnsExpn</ColumnUid> <Value>");
                Sbuilder = CommonFunctions.AppendXML(Sbuilder, FormsB1.ConvertDecimalToString(Convert.ToDecimal(TRANSPORT_COAST)));
                Sbuilder.Append("</Value></Cell>");

                Sbuilder.Append("<Cell> <ColumnUid>WB_begDate</ColumnUid> <Value>");
                Sbuilder = CommonFunctions.AppendXML(Sbuilder, oRecordSet.Fields.Item("BaseDocDate").Value.ToString("yyyyMMdd") == "18991230" ? "" : oRecordSet.Fields.Item("BaseDocDate").Value.ToString("yyyyMMdd"));
                Sbuilder.Append("</Value></Cell>");

                BaseDocNum = (int)oRecordSet.Fields.Item("BaseDocNum").Value;
                Sbuilder.Append("<Cell> <ColumnUid>DocNum</ColumnUid> <Value>");
                Sbuilder = CommonFunctions.AppendXML(Sbuilder, (BaseDocNum == 0 ? "" : BaseDocNum.ToString()));
                Sbuilder.Append("</Value></Cell>");

                BaseDType = (string)oRecordSet.Fields.Item("BaseDType").Value;
                Sbuilder.Append("<Cell> <ColumnUid>BaseDType</ColumnUid> <Value>");
                Sbuilder = CommonFunctions.AppendXML(Sbuilder, (BaseDType == "" ? "" : BaseDType.ToString()));
                Sbuilder.Append("</Value></Cell>");

                DocEntry = (int)oRecordSet.Fields.Item("DocEntry").Value;
                Sbuilder.Append("<Cell> <ColumnUid>BaseDoc</ColumnUid> <Value>");
                Sbuilder = CommonFunctions.AppendXML(Sbuilder, (DocEntry == 0 ? "" : DocEntry.ToString()));
                Sbuilder.Append("</Value></Cell>");

                Sbuilder.Append("<Cell> <ColumnUid>Whs</ColumnUid> <Value>");
                Sbuilder = CommonFunctions.AppendXML(Sbuilder, oRecordSet.Fields.Item("Whs").Value);
                Sbuilder.Append("</Value></Cell>");

                Sbuilder.Append("<Cell> <ColumnUid>RS_Status</ColumnUid> <Value>");
                Sbuilder = CommonFunctions.AppendXML(Sbuilder, RS_STATUS);
                Sbuilder.Append("</Value></Cell>");

                //დეტალური როცა არის
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
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, FormsB1.ConvertDecimalToString(Convert.ToDecimal(Amount_Full)));
                    Sbuilder.Append("</Value></Cell>");
                }

                Sbuilder.Append("</Row>");

                count++;

                oRecordSet.MoveNext();
            }

            if (RSDataTable.Rows.Count > 0 & (FilterCompStat == "1" || FilterCompStat == "0" || FilterCompStat == ""))
            {
                DataRow[] RemainingRows;
                RemainingRows = RSDataTable.Select("RowLinked = 'N'");

                double TranspCost;
                string cardName;
                string WAYBILL_NUMBER;
                string cTIN;
                string cCode;
                double RSFULL_AMOUNT = 0;
                string wbTRANSPORT_COAST;
                string RS_st;

                for (int i = 0; i < RemainingRows.Length; i++)
                {
                    RS_st = RemainingRows[i]["STATUS"].ToString();
                    //სტატუსის ფილტრი
                    if (((Fstatus == "-99" || Fstatus == "") || statuses.IndexOf(RS_st) > 0) == false)
                    {
                        continue;
                    }

                    Sbuilder.Append("<Row>");

                    cardName = "";
                    WAYBILL_NUMBER = "";
                    cTIN = RemainingRows[i]["TIN"].ToString().Trim();
                    cCode = BusinessPartners.GetCardCodeByTin( cTIN, "S", out cardName);

                    if (String.IsNullOrWhiteSpace(cCode) == false)
                    {
                        Sbuilder.Append("<Cell> <ColumnUid>BaseCard</ColumnUid> <Value>");
                        Sbuilder = CommonFunctions.AppendXML(Sbuilder, cCode);
                        Sbuilder.Append("</Value></Cell>");
                    }

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

                    Sbuilder.Append("<Cell> <ColumnUid>TYPE</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, RemainingRows[i]["TYPE"].ToString());
                    Sbuilder.Append("</Value></Cell>");

                    Sbuilder.Append("<Cell> <ColumnUid>WB_status</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, RS_st);
                    Sbuilder.Append("</Value></Cell>");

                    Sbuilder.Append("<Cell> <ColumnUid>WB_strAddrs</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, RemainingRows[i]["START_ADDRESS"].ToString());
                    Sbuilder.Append("</Value></Cell>");

                    Sbuilder.Append("<Cell> <ColumnUid>WB_endAddrs</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, RemainingRows[i]["END_ADDRESS"].ToString());
                    Sbuilder.Append("</Value></Cell>");

                    strDELIVERY_DATE = RemainingRows[i]["DELIVERY_DATE"].ToString();
                    if (strDELIVERY_DATE != "")
                    {
                        Sbuilder.Append("<Cell> <ColumnUid>WB_delvDate</ColumnUid> <Value>");
                        Sbuilder = CommonFunctions.AppendXML(Sbuilder, (DateTime.TryParse(strDELIVERY_DATE, out DeliveryDate) == false ? DateTime.MinValue : DeliveryDate).ToString("yyyyMMdd"));
                        Sbuilder.Append("</Value></Cell>");
                    }

                    Sbuilder.Append("<Cell> <ColumnUid>WB_ID</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, RemainingRows[i]["ID"].ToString());
                    Sbuilder.Append("</Value></Cell>");

                    strACTIVATE_DATE = RemainingRows[i]["ACTIVATE_DATE"].ToString();
                    if (strACTIVATE_DATE != "")
                    {
                        Sbuilder.Append("<Cell> <ColumnUid>WB_actDate</ColumnUid> <Value>");
                        Sbuilder = CommonFunctions.AppendXML(Sbuilder, (DateTime.TryParse(strACTIVATE_DATE, out ActivateDate) == false ? DateTime.MinValue : ActivateDate).ToString("yyyyMMdd"));
                        Sbuilder.Append("</Value></Cell>");
                    }

                    Sbuilder.Append("<Cell> <ColumnUid>WB_vehicNum</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, RemainingRows[i]["CAR_NUMBER"].ToString());
                    Sbuilder.Append("</Value></Cell>");

                    Sbuilder.Append("<Cell> <ColumnUid>WB_drivTin</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, RemainingRows[i]["DRIVER_TIN"].ToString());
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
                        Sbuilder = CommonFunctions.AppendXML(Sbuilder, FormsB1.ConvertDecimalToString(Convert.ToDecimal(RSAmount)));
                        Sbuilder.Append("</Value></Cell>");

                        Sbuilder.Append("<Cell> <ColumnUid>BarCode</ColumnUid> <Value>");
                        Sbuilder = CommonFunctions.AppendXML(Sbuilder, RemainingRows[i]["BAR_CODE"].ToString());
                        Sbuilder.Append("</Value></Cell>");

                        RSQuantity = Convert.ToDouble(RemainingRows[i]["Quantity"].ToString(), System.Globalization.CultureInfo.InvariantCulture);
                        Sbuilder.Append("<Cell> <ColumnUid>RSQuantity</ColumnUid> <Value>");
                        Sbuilder = CommonFunctions.AppendXML(Sbuilder, FormsB1.ConvertDecimalToString(Convert.ToDecimal(RSQuantity)));
                        Sbuilder.Append("</Value></Cell>");

                        Sbuilder.Append("<Cell> <ColumnUid>RSName</ColumnUid> <Value>");
                        Sbuilder = CommonFunctions.AppendXML(Sbuilder, RemainingRows[i]["W_NAME"].ToString());
                        Sbuilder.Append("</Value></Cell>");
                    }
                    else
                    {
                        RSFULL_AMOUNT = Convert.ToDouble(RemainingRows[i]["FULL_AMOUNT"].ToString(), System.Globalization.CultureInfo.InvariantCulture);
                        Sbuilder.Append("<Cell> <ColumnUid>RSAmount</ColumnUid> <Value>");
                        Sbuilder = CommonFunctions.AppendXML(Sbuilder, FormsB1.ConvertDecimalToString(Convert.ToDecimal(RSFULL_AMOUNT)));
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
            oGrid.Columns.Item("Whs").TitleObject.Caption = BDOSResources.getTranslate("Warehouse");
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
            oComboComparStat.ValidValues.Add("7", BDOSResources.getTranslate("NotFoundOnSiteOnThisPeriod"));

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

            //Whs
            SAPbouiCOM.EditTextColumn oWhs = (SAPbouiCOM.EditTextColumn)oGrid.Columns.Item("Whs");
            oWhs.LinkedObjectType = "64";

            //WhsFrom
            SAPbouiCOM.EditTextColumn oWhsFrom = (SAPbouiCOM.EditTextColumn)oGrid.Columns.Item("WhsFrom");
            oWhsFrom.LinkedObjectType = "64";

            //U_vehicNum
            SAPbouiCOM.EditTextColumn oU_vehicNum = (SAPbouiCOM.EditTextColumn)oGrid.Columns.Item("WB_vehicNum");
            oU_vehicNum.LinkedObjectType = "UDO_F_BDO_VECL_D";

            //U_baseDocT
            SAPbouiCOM.EditTextColumn obaseDocT = (SAPbouiCOM.EditTextColumn)oGrid.Columns.Item("BaseDoc");
            obaseDocT.LinkedObjectType = "14";

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
            oColumns.Item(8).Visible = false;
            oColumns.Item(17).Visible = false;
            oColumns.Item(22).Visible = false;
            oColumns.Item(28).Visible = false;
            oColumns.Item(32).Visible = false;
            if (FOption == "1")
            {
                oColumns.Item(26).Visible = false;
                oColumns.Item(27).Visible = false;
                oColumns.Item(28).Visible = false;
                oColumns.Item(29).Visible = false;
                oColumns.Item(30).Visible = false;
                oColumns.Item(31).Visible = false;
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
                        Boolean isleaf = oGrid.Rows.IsLeaf(i);
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
                        //string DocType = oDataTable.GetValue("TYPE", dTableRow);
                        string BaseDType = oDataTable.GetValue("BaseDType", dTableRow);

                        //UbaseDocT
                        SAPbouiCOM.EditTextColumn obaseDocT = (SAPbouiCOM.EditTextColumn)oGrid.Columns.Item("BaseDoc");

                        if (BaseDType == "18")
                        {
                            obaseDocT.LinkedObjectType = "18";
                        }
                        else if (BaseDType == "20")
                        {
                            obaseDocT.LinkedObjectType = "20";
                        }
                        else if (BaseDType == "19") //დაბრუნება
                        {
                            obaseDocT.LinkedObjectType = "19";
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

        public static void addMenus( out string errorText)
        {
            errorText = null;

            SAPbouiCOM.Menus moduleMenus;
            SAPbouiCOM.MenuItem menuItem;
            SAPbouiCOM.MenuCreationParams oCreationPackage;

            try
            {
                // Find the id of the menu into wich you want to add your menu item
                menuItem = Program.uiApp.Menus.Item("43534");

                // Get the menu collection of SAP Business One
                moduleMenus = menuItem.SubMenus;

                // Add a pop-up menu item
                oCreationPackage = (SAPbouiCOM.MenuCreationParams)Program.uiApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                oCreationPackage.Checked = false;
                oCreationPackage.Enabled = true;
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "BDO_WBRA";
                oCreationPackage.String = BDOSResources.getTranslate("ReceivedWaybillsAnalysis");
                oCreationPackage.Position = -1;

                menuItem = moduleMenus.AddEx(oCreationPackage);
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
        }

        public static string getQueryText(DateTime startDate, DateTime endDate, string cardCode, string itypes, string statuses, string foption)
        {
            string tempQuery = @"
         SELECT
             ""BASEDOCGDS"".""Type"",
	         ""OCRD"".""CardName"",
	         ""OCRD"".""LicTradNum"", " +
             ((foption != "1") ? @"
	         ""BASEDOCGDS"".""LineNum"",
	         ""BASEDOCGDS"".""Dscription"" AS ItemDesc,
	         ""BASEDOCGDS"".""ItemCode"",
	         ""OITM"".""SWW"",
	         ""OITM"".""InvntryUom"",
	         ""OITM"".""CodeBars"", " : " ") + @"
	         ""BASEDOCGDS"".""BaseCard"",
	         ""BASEDOCGDS"".""BaseDType"",
	         ""BASEDOCGDS"".""DocEntry"",
	         ""BASEDOCGDS"".""Whs"",
	         ""BASEDOCGDS"".""DocNum"" AS BaseDocNum,
	         ""BASEDOCGDS"".""DocDate"" AS BaseDocDate,
	         ""BASEDOCGDS"".""U_BDO_WBID"" AS WBID,
	         ""BASEDOCGDS"".""U_BDO_WBNo"" AS WBNo,
	         ""BASEDOCGDS"".""U_BDO_WBSt"" AS WBSt,
	         ""BASEDOCGDS"".""CANCELED"" AS CANCELED,
	         SUM(""BASEDOCGDS"".""Quantity"") AS Quantity,
	         MAX(""BASEDOCGDS"".""DocTotal"") AS DocTotal,
	         SUM(""BASEDOCGDS"".""GTotal"") AS GTotal 
	        FROM  (SELECT
            '2' AS ""Type"",
	         ""PCH"".""Quantity"",
	         ""PCH"".""GTotal"",
	         ""PCH"".""LineNum"",
	         ""PCH"".""Dscription"",
	         ""PCH"".""ItemCode"",
	         ""PCH"".""BaseCard"",
            '18' AS ""BaseDType"",
	         ""PCH"".""DocEntry"",
	         ""PCH"".""Whs"",
	         ""OPCH"".""DocNum"",
	         ""OPCH"".""DocDate"",
	         ""OPCH"".""U_BDO_WBID"",
	         ""OPCH"".""U_BDO_WBNo"",
	         ""OPCH"".""U_BDO_WBSt"",
	         ""PCH"".""DocTotal"",
	         ""OPCH"".""CANCELED"" 
		        FROM (SELECT
	         ""PCH1"".""Quantity"" * ""PCH1"".""NumPerMsr"" AS ""Quantity"",
	         ""PCH1"".""GTotal"",
	         ""OPCH"".""DocTotal"" + ""OPCH"".""DpmAmnt"" AS ""DocTotal"",
	         ""PCH1"".""LineNum"",
	         ""PCH1"".""Dscription"",
	         ""PCH1"".""ItemCode"",
	         ""PCH1"".""BaseCard"",
	         ""OPCH"".""U_BDO_WBID"",
	         ""PCH1"".""DocEntry"",
	         ""PCH1"".""WhsCode"" AS ""Whs""
			        FROM ""PCH1""  INNER JOIN ""OPCH"" ON ""OPCH"".""DocEntry"" = ""PCH1"".""DocEntry""
                    WHERE ""PCH1"".""BaseType"" <> 20
			        UNION ALL SELECT
	         ""RPC1"".""Quantity"" * (-1) * (CASE WHEN ""RPC1"".""NoInvtryMv"" = 'Y' 
				        THEN 0 
				        ELSE 1 
				        END) * ""RPC1"".""NumPerMsr"",
	         ""RPC1"".""GTotal"" * (-1),
			 ""ORPC"".""DocTotal"" * (-1) + ""ORPC"".""DpmAmnt"" * (-1),
	         ""RPC1"".""BaseLine"",
	         ""RPC1"".""Dscription"",
	         ""RPC1"".""ItemCode"",
	         ""RPC1"".""BaseCard"",
	         ""ORPC"".""U_BDO_WBID"",
	         ""RPC1"".""BaseEntry"",
	         ""RPC1"".""WhsCode"" AS ""Whs""
			        FROM ""RPC1"" 
			        INNER JOIN ""ORPC"" ON ""ORPC"".""DocEntry"" = ""RPC1"".""DocEntry"" 
			        WHERE ""RPC1"".""TargetType"" < 0
			        AND ""ORPC"".""U_BDO_WBID"" IN (SELECT ""OPCH"".""U_BDO_WBID"" FROM ""OPCH"" WHERE ""OPCH"".""CANCELED"" = 'N') ) AS ""PCH"" 
		        
		    INNER JOIN ""OPCH"" ON ""PCH"".""DocEntry"" = ""OPCH"".""DocEntry"" 
		        
		    UNION ALL 

            SELECT
            '2',
	         ""PDN"".""Quantity"",
	         ""PDN"".""GTotal"",
	         ""PDN"".""LineNum"",
	         ""PDN"".""Dscription"",
	         ""PDN"".""ItemCode"",
	         ""PDN"".""BaseCard"",
            '20' AS ""BaseDType"",
	         ""PDN"".""DocEntry"",
	         ""PDN"".""Whs"",
	         ""OPDN"".""DocNum"",
	         ""OPDN"".""DocDate"",
	         ""OPDN"".""U_BDO_WBID"",
	         ""OPDN"".""U_BDO_WBNo"",
	         ""OPDN"".""U_BDO_WBSt"",
	         ""PDN"".""DocTotal"",
	         ""OPDN"".""CANCELED"" 
		        FROM (SELECT
	         ""PDN1"".""Quantity"" * ""PDN1"".""NumPerMsr"" AS ""Quantity"",
	         ""PDN1"".""GTotal"",
	         ""OPDN"".""DocTotal"" + ""OPDN"".""DpmAmnt"" AS ""DocTotal"",
	         ""PDN1"".""LineNum"",
	         ""PDN1"".""Dscription"",
	         ""PDN1"".""ItemCode"",
	         ""PDN1"".""BaseCard"",
	         ""OPDN"".""U_BDO_WBID"",
	         ""PDN1"".""DocEntry"",
	         ""PDN1"".""WhsCode"" AS ""Whs""
			        FROM ""PDN1""  INNER JOIN ""OPDN"" ON ""OPDN"".""DocEntry"" = ""PDN1"".""DocEntry""
			        UNION ALL SELECT
	         ""RPC1"".""Quantity"" * (-1) * (CASE WHEN ""RPC1"".""NoInvtryMv"" = 'Y' 
				        THEN 0 
				        ELSE 1 
				        END) * ""RPC1"".""NumPerMsr"",
	         ""RPC1"".""GTotal"" * (-1),
			 ""ORPC"".""DocTotal"" * (-1) + ""ORPC"".""DpmAmnt"" * (-1),
	         ""RPC1"".""BaseLine"",
	         ""RPC1"".""Dscription"",
	         ""RPC1"".""ItemCode"",
	         ""RPC1"".""BaseCard"",
	         ""ORPC"".""U_BDO_WBID"",
	         ""RPC1"".""BaseEntry"",
	         ""RPC1"".""WhsCode"" AS ""Whs""
			        FROM ""RPC1"" 
			        INNER JOIN ""ORPC"" ON ""ORPC"".""DocEntry"" = ""RPC1"".""DocEntry"" 
			        WHERE ""RPC1"".""TargetType"" < 0
			        AND ""ORPC"".""U_BDO_WBID"" IN (SELECT ""OPDN"".""U_BDO_WBID"" FROM ""OPDN"" WHERE ""OPDN"".""CANCELED"" = 'N') ) AS ""PDN"" 
		        
		    INNER JOIN ""OPDN"" ON ""PDN"".""DocEntry"" = ""OPDN"".""DocEntry""
 
		    UNION ALL 
		    SELECT
             '5',
	         ""RPC1"".""Quantity"" * (CASE WHEN ""RPC1"".""NoInvtryMv"" = 'Y' 
			 THEN 0 
			 ELSE 1 
			 END) * ""RPC1"".""NumPerMsr"",
	         ""RPC1"".""GTotal"",
	         ""RPC1"".""BaseLine"",
	         ""RPC1"".""Dscription"",
	         ""RPC1"".""ItemCode"",
	         ""RPC1"".""BaseCard"",
            '19' AS ""BaseDType"",
	         ""RPC1"".""DocEntry"",
	         ""RPC1"".""WhsCode"" AS ""Whs"",
	         ""ORPC"".""DocNum"",
	         ""ORPC"".""DocDate"",
	         ""ORPC"".""U_BDO_WBID"",
	         ""ORPC"".""U_BDO_WBNo"",
	         ""ORPC"".""U_BDO_WBSt"",
	         ""ORPC"".""DocTotal"" + ""ORPC"".""DpmAmnt"",
	         ""ORPC"".""CANCELED"" 
		        FROM ""RPC1"" 
		        INNER JOIN ""ORPC"" ON ""ORPC"".""DocEntry"" = ""RPC1"".""DocEntry"" 
		        WHERE ""RPC1"".""TargetType"" < 0 AND NOT ""ORPC"".""U_BDO_WBID"" IN (SELECT ""OPCH"".""U_BDO_WBID"" FROM ""OPCH"" WHERE ""OPCH"".""CANCELED"" = 'Y')
		        ) AS ""BASEDOCGDS""   
	        LEFT JOIN ""OCRD"" AS ""OCRD"" ON ""BASEDOCGDS"".""BaseCard"" = ""OCRD"".""CardCode"" 
	        LEFT JOIN ""OITM"" ON ""BASEDOCGDS"".""ItemCode"" = ""OITM"".""ItemCode"" 
	        WHERE ((""OITM"".""ItemType"" = 'I' 
	        AND ""OITM"".""InvntItem"" = 'Y') OR ""OITM"".""ItemType"" = 'F') AND ""BASEDOCGDS"".""CANCELED"" = 'N' " +

            ((startDate != new DateTime(1, 1, 1)) ? @" AND ""BASEDOCGDS"".""DocDate"" >= '" + startDate.ToString("yyyyMMdd") + "' " : " ") +
              ((endDate != new DateTime(1, 1, 1)) ? @" AND ""BASEDOCGDS"".""DocDate"" <= '" + endDate.ToString("yyyyMMdd") + "' " : " ") +
              ((cardCode != "") ? @" AND ""BASEDOCGDS"".""BaseCard"" = N'" + cardCode.Replace("'", "''") + "' " : " ") +
              ((itypes == "2" || itypes == "3" || itypes == "5") ? @" AND ""BASEDOCGDS"".""Type"" ='" + ((itypes == "5") ? "5" : "2") + "' " : " ") +

//   ((statuses != "" && statuses != "-99") ? @"AND (CASE WHEN ""U_BDO_WBSt"" = '1' 
                //	    THEN '0' WHEN ""U_BDO_WBSt""= '2' 
                //	    THEN '1' WHEN ""U_BDO_WBSt"" = '3' 
                //	    THEN '2' WHEN ""U_BDO_WBSt"" = '4' 
                //	    THEN '-1' WHEN ""U_BDO_WBSt"" = '5' 
                //	    THEN '-2' 
                //	    ELSE '8' 
                //	    END) = '" + statuses + "' " : " ") +

    @"GROUP BY 
             ""BASEDOCGDS"".""Type"",
             ""OCRD"".""CardName"",
	         ""OCRD"".""LicTradNum"",
	         ""BASEDOCGDS"".""BaseCard"",
	         ""BASEDOCGDS"".""DocEntry"", ""BASEDOCGDS"".""BaseDType"",
	         ""BASEDOCGDS"".""Whs"", " +
 ((foption != "1") ? @"			 
	         ""BASEDOCGDS"".""LineNum"",
	         ""BASEDOCGDS"".""Dscription"",
	         ""BASEDOCGDS"".""ItemCode"",
	         ""OITM"".""SWW"",
	         ""OITM"".""InvntryUom"",
	         ""OITM"".""CodeBars""," : " ") + @"  
	         ""BASEDOCGDS"".""DocDate"",
	         ""BASEDOCGDS"".""U_BDO_WBID"",
	         ""BASEDOCGDS"".""U_BDO_WBNo"",
	         ""BASEDOCGDS"".""U_BDO_WBSt"",
	         ""BASEDOCGDS"".""DocNum"",
	         ""BASEDOCGDS"".""CANCELED"" ";

            return tempQuery;
        }

        public static void reArrangeFormItems(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.Item oItem = null;

            oItem = oForm.Items.Item("Option");
            oItem.Left = 395;

            //oItem = oForm.Items.Item("33_U_BC");
            //oItem.Left = oForm.ClientWidth - 6 - oItem.Width;
            //oItem.Top = oForm.ClientHeight - 25;
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
