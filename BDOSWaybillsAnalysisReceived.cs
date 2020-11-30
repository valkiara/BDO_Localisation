using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Collections;

namespace BDO_Localisation_AddOn
{
    static partial class BDOSWaybillsAnalysisReceived
    {
        static Hashtable selectedBusinessPartners = new Hashtable();
        static Hashtable selectedTypes = new Hashtable();
        static Hashtable selectedStatuses = new Hashtable();
        static Hashtable selectedCompareStatuses = new Hashtable();
        static string buttonType = null;

        public static void createForm(out string errorText)
        {
            selectedBusinessPartners = new Hashtable();
            selectedTypes = new Hashtable();
            selectedStatuses = new Hashtable();
            selectedCompareStatuses = new Hashtable();
            buttonType = null;

            errorText = null;
            int formHeight = Program.uiApp.Desktop.Height;
            int formWidth = Program.uiApp.Desktop.Width;

            //ფორმის აუცილებელი თვისებები
            Dictionary<string, object> formProperties = new Dictionary<string, object>();
            formProperties.Add("UniqueID", "BDOSWBRAn");
            formProperties.Add("BorderStyle", SAPbouiCOM.BoFormBorderStyle.fbs_Sizable);
            formProperties.Add("Title", BDOSResources.getTranslate("ReceivedWaybillsAnalysis"));
            formProperties.Add("ClientWidth", formWidth);
            formProperties.Add("ClientHeight", formHeight);

            SAPbouiCOM.Form oForm;
            bool newForm;
            bool formExist = FormsB1.createForm(formProperties, out oForm, out newForm, out errorText);

            if (errorText != null)
            {
                return;
            }

            if (formExist)
            {
                if (newForm)
                {
                    Dictionary<string, object> formItems;

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
                    oDataTable.Columns.Add("PrjCode", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20);//22 როცა დეტალურია
                    oDataTable.Columns.Add("PrjName", SAPbouiCOM.BoFieldsType.ft_Text, 100);//23 როცა დეტალურია
                    oDataTable.Columns.Add("WhsFrom", SAPbouiCOM.BoFieldsType.ft_Text, 50);//24

                    oDataTable.Columns.Add("RS_Status", SAPbouiCOM.BoFieldsType.ft_Text, 50);//25
                    oDataTable.Columns.Add("RSAmount", SAPbouiCOM.BoFieldsType.ft_Sum, 50);//26
                    oDataTable.Columns.Add("Gtotal", SAPbouiCOM.BoFieldsType.ft_Sum, 50);//27

                    //დეტალური როცა არის
                    oDataTable.Columns.Add("ItemCode", SAPbouiCOM.BoFieldsType.ft_Text, 50);//28
                    oDataTable.Columns.Add("ItemName", SAPbouiCOM.BoFieldsType.ft_Text, 100); //29
                    oDataTable.Columns.Add("InvntryUom", SAPbouiCOM.BoFieldsType.ft_Text, 50);//30
                    oDataTable.Columns.Add("BarCode", SAPbouiCOM.BoFieldsType.ft_Text, 50);//31
                    oDataTable.Columns.Add("Quantity", SAPbouiCOM.BoFieldsType.ft_Quantity, 50);//32
                    oDataTable.Columns.Add("RSUom", SAPbouiCOM.BoFieldsType.ft_Text, 50);//33
                    oDataTable.Columns.Add("RSQuantity", SAPbouiCOM.BoFieldsType.ft_Quantity, 50);//34
                    oDataTable.Columns.Add("RSName", SAPbouiCOM.BoFieldsType.ft_Text, 50);//35
                    //დეტალური როცა არის
                    oDataTable.Columns.Add("BaseDType", SAPbouiCOM.BoFieldsType.ft_Text, 50);//36----
                    oDataTable.Columns.Add("WB_trnsSide", SAPbouiCOM.BoFieldsType.ft_Text, 50);//37
                    oDataTable.Columns.Add("DocCanld", SAPbouiCOM.BoFieldsType.ft_Text, 50);//38
                    string itemName;

                    int left_s = 6;
                    int width_s = 115;
                    int left_e = left_s + width_s + 1;
                    int height = 15;
                    int top = 6;
                    int width_e = 200;
                    int left_s2 = left_e + width_e + width_e / 6 + 15;
                    int left_e2 = left_s2 + width_s + 1;
                    int top2 = top + height + 1;
                    int left_s3 = left_e2 + width_e + width_e / 6 + 15;
                    int left_e3 = left_s3 + width_s + 1;
                    int top3 = top + height + 1;

                    formItems = new Dictionary<string, object>();
                    itemName = "PeriodS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("Period"));
                    formItems.Add("LinkTo", "DateFromE");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    string startOfMonthStr = DateTime.Today.ToString("yyyyMMdd");

                    formItems = new Dictionary<string, object>();
                    itemName = "DateFromE";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_DATE);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Left", left_e);
                    formItems.Add("Width", width_e / 2);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("ValueEx", startOfMonthStr);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    string endOfMonthStr = DateTime.Today.ToString("yyyyMMdd");

                    formItems = new Dictionary<string, object>();
                    itemName = "DateToE";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_DATE);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Left", left_e + width_e / 2);
                    formItems.Add("Width", width_e / 2);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("ValueEx", endOfMonthStr);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    top += height + 1;

                    formItems = new Dictionary<string, object>();
                    itemName = "WBDocTpSt";
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Caption", BDOSResources.getTranslate("Type"));
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "WBDocTp";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_LONG_TEXT);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EXTEDIT);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Left", left_e);
                    formItems.Add("Width", width_e);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height * 3);
                    formItems.Add("UID", itemName);
                    formItems.Add("Enabled", false);
                    formItems.Add("ScrollBars", SAPbouiCOM.BoScrollBars.sb_Vertical);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "TPChooseB";
                    formItems.Add("Caption", "...");
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left_e + width_e + 1);
                    formItems.Add("Width", width_e / 6);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    top += height * 3 + 1;

                    formItems = new Dictionary<string, object>();
                    itemName = "ClientIDSt";
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Caption", BDOSResources.getTranslate("CardCode"));
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    FormsB1.addChooseFromList(oForm, true, "2", "BusinessPartner_CFL");

                    //პირობის დადება ბიზნესპარტნიორის არჩევის სიაზე
                    SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item("BusinessPartner_CFL");
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
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_LONG_TEXT);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EXTEDIT);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Left", left_e);
                    formItems.Add("Width", width_e);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height * 3);
                    formItems.Add("UID", itemName);
                    formItems.Add("Enabled", false);
                    formItems.Add("ScrollBars", SAPbouiCOM.BoScrollBars.sb_Vertical);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "BPChooseB";
                    formItems.Add("Caption", "...");
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left_e + width_e + 1);
                    formItems.Add("Width", width_e / 6);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("ChooseFromListUID", "BusinessPartner_CFL");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    top += height * 3 + 10;

                    formItems = new Dictionary<string, object>();
                    itemName = "WbFillTb";
                    formItems.Add("Caption", BDOSResources.getTranslate("Fill"));
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "WBStatusSt";
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s2);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top2);
                    formItems.Add("Caption", BDOSResources.getTranslate("Status"));
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "WBStatus";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_LONG_TEXT);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EXTEDIT);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Left", left_e2);
                    formItems.Add("Width", width_e);
                    formItems.Add("Top", top2);
                    formItems.Add("Height", height * 3);
                    formItems.Add("UID", itemName);
                    formItems.Add("Enabled", false);
                    formItems.Add("ScrollBars", SAPbouiCOM.BoScrollBars.sb_Vertical);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "STChooseB";
                    formItems.Add("Caption", "...");
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left_e2 + width_e + 1);
                    formItems.Add("Width", width_e / 6);
                    formItems.Add("Top", top2);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    top2 += height * 3 + 1;

                    formItems = new Dictionary<string, object>();
                    itemName = "CompStatSt";
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s2);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top2);
                    formItems.Add("Caption", BDOSResources.getTranslate("CompSt"));
                    formItems.Add("Description", BDOSResources.getTranslate("CompareStatus"));
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "CompStat";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_LONG_TEXT);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EXTEDIT);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Left", left_e2);
                    formItems.Add("Width", width_e);
                    formItems.Add("Top", top2);
                    formItems.Add("Height", height * 3);
                    formItems.Add("UID", itemName);
                    formItems.Add("Enabled", false);
                    formItems.Add("ScrollBars", SAPbouiCOM.BoScrollBars.sb_Vertical);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "CSTChooseB";
                    formItems.Add("Caption", "...");
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left_e2 + width_e + 1);
                    formItems.Add("Width", width_e / 6);
                    formItems.Add("Top", top2);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "OptionSt";
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s3);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top3);
                    formItems.Add("Caption", BDOSResources.getTranslate("Option"));
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    Dictionary<string, string> listValidValuesDict = new Dictionary<string, string>();
                    listValidValuesDict.Add("0", BDOSResources.getTranslate("Details"));
                    listValidValuesDict.Add("1", BDOSResources.getTranslate("General"));

                    formItems = new Dictionary<string, object>();
                    itemName = "Option";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("Length", 30);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                    formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("Left", left_e3);
                    formItems.Add("Width", width_e);
                    formItems.Add("Top", top3);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("ValidValues", listValidValuesDict);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    SAPbouiCOM.ComboBox oComboBox_Option = (SAPbouiCOM.ComboBox)oForm.Items.Item("Option").Specific;
                    oComboBox_Option.Select("1", SAPbouiCOM.BoSearchKey.psk_ByValue);

                    top3 += height + 1;

                    formItems = new Dictionary<string, object>();
                    itemName = "TranSideS";
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s3);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top3);
                    formItems.Add("Caption", BDOSResources.getTranslate("TransportationSide"));
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    listValidValuesDict = new Dictionary<string, string>();
                    listValidValuesDict.Add("0", BDOSResources.getTranslate("All"));
                    listValidValuesDict.Add("1", BDOSResources.getTranslate("Buyer"));
                    listValidValuesDict.Add("2", BDOSResources.getTranslate("Seller"));

                    formItems = new Dictionary<string, object>();
                    itemName = "TranSide";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("Length", 30);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                    formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("Left", left_e3);
                    formItems.Add("Width", width_e);
                    formItems.Add("Top", top3);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("ValidValues", listValidValuesDict);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    oComboBox_Option = (SAPbouiCOM.ComboBox)oForm.Items.Item("TranSide").Specific;
                    oComboBox_Option.Select("0", SAPbouiCOM.BoSearchKey.psk_ByValue);

                    top3 += height + 1;

                    formItems = new Dictionary<string, object>();
                    itemName = "WBNumberS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s3);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top3);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("WaybillNumber"));

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "WBNumber";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("Length", 30);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Left", left_e3);
                    formItems.Add("Width", width_e);
                    formItems.Add("Top", top3);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    top3 += height + 1;

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
                    formItems.Add("Left", left_s3);
                    formItems.Add("Width", width_s * 2);
                    formItems.Add("Top", top3);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("SetGridColor"));
                    formItems.Add("ValOff", "N");
                    formItems.Add("ValOn", "Y");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    //Grid
                    top = top + 30;

                    itemName = "WbTable";
                    formItems = new Dictionary<string, object>();
                    formItems.Add("isDataSource", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_GRID);
                    formItems.Add("Left", left_s);
                    formItems.Add("Top", top);
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "btnColl";
                    formItems.Add("Caption", BDOSResources.getTranslate("Collapse"));
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "btnExp";
                    formItems.Add("Caption", BDOSResources.getTranslate("Expand"));
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left_e);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }
                }

                oForm.Visible = true;
                oForm.Select();
            }
        }

        public static void updateGrid(SAPbouiCOM.Form oForm)
        {
            try
            {
                oForm.Freeze(true);

                string errorText;
                Dictionary<string, string> rsSettings = CompanyDetails.getRSSettings(out errorText);
                if (errorText != null)
                {
                    throw new Exception(errorText);
                }

                string su = rsSettings["SU"];
                string sp = rsSettings["SP"];
                WayBill oWayBill = new WayBill(su, sp, rsSettings["ProtocolType"]);

                bool chek_service_user = oWayBill.chek_service_user(su, sp, out errorText);
                if (!chek_service_user)
                {
                    errorText = BDOSResources.getTranslate("ServiceUserPasswordNotCorrect");
                    throw new Exception(errorText);
                }

                DateTime startDate;
                string startDateStr = oForm.DataSources.UserDataSources.Item("DateFromE").ValueEx;
                DateTime BeginDate = new DateTime(1, 1, 1);

                if (DateTime.TryParseExact(startDateStr, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out startDate))
                    BeginDate = startDate;

                DateTime endDate;
                string endDateStr = oForm.DataSources.UserDataSources.Item("DateToE").ValueEx;
                DateTime EndDate = DateTime.Today;

                if (DateTime.TryParseExact(endDateStr, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out endDate))
                    EndDate = endDate;

                //string itypes = oForm.DataSources.UserDataSources.Item("WBDocTp").ValueEx;
                //string cardCode = oForm.DataSources.UserDataSources.Item("ClientID").Value;
                //cardCode = cardCode.Trim();
                //string buyer_tin = "";

                //if (cardCode != "")
                //{
                //    SAPbobsCOM.BusinessPartners oBP;
                //    oBP = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);
                //    oBP.GetByKey(cardCode);

                //    buyer_tin = oBP.UserFields.Fields.Item("LicTradNum").Value;
                //}

                //string Fstatus = oForm.DataSources.UserDataSources.Item("WBStatus").ValueEx;
                string FOption = oForm.DataSources.UserDataSources.Item("option").ValueEx;
                string FNUMBER = oForm.DataSources.UserDataSources.Item("WBNumber").ValueEx;
                string TranSide = oForm.DataSources.UserDataSources.Item("TranSide").ValueEx;

                List<string> buyerTinsList = new List<string>();
                List<string> cardCodesList = new List<string>();
                foreach (DictionaryEntry item in selectedBusinessPartners)
                {
                    cardCodesList.Add("N'" + item.Key + "'");
                    if (!string.IsNullOrEmpty(item.Value.ToString()))
                        buyerTinsList.Add(item.Value.ToString());
                }

                List<string> selectedTypesList = new List<string>();
                foreach (DictionaryEntry item in selectedTypes)
                {
                    selectedTypesList.Add("'" + item.Key + "'");
                }

                string query = getQueryText(BeginDate, EndDate, string.Join(",", cardCodesList), string.Join(",", selectedTypesList), FOption);

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
                RSDataTable.Columns.Add("ItemName", typeof(string));
                RSDataTable.Columns.Add("W_NAME", typeof(string));
                RSDataTable.Columns.Add("BAR_CODE", typeof(string));
                RSDataTable.Columns.Add("AMOUNT", typeof(string));
                RSDataTable.Columns.Add("Quantity", typeof(string));
                RSDataTable.Columns.Add("RSUom", typeof(string));
                RSDataTable.Columns.Add("PrjCode", typeof(string));
                RSDataTable.Columns.Add("PrjName", typeof(string));
                RSDataTable.Columns.Add("TRAN_COST_PAYER", typeof(string));

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
                string WBGUntCode = "";
                string PrjCode = "";
                string PrjName = "";
                string ItmCode = "";
                string ItemName = "";

                string StartAddress = "";
                string EndAddress = "";

                DateTime startDateParam = new DateTime();
                DateTime endDateParam = new DateTime();
                startDateParam = startDate;

                while (startDateParam < EndDateForWS)
                {
                    endDateParam = startDateParam.AddDays(2);

                    if (endDateParam > EndDateForWS)
                        endDateParam = EndDateForWS;

                    bool isBpFilter = buyerTinsList.Count > 0;
                    int countBpFilter = isBpFilter ? buyerTinsList.Count : 1;

                    for (int i = 0; i < countBpFilter; i++)
                    {
                        string seller_id = isBpFilter ? buyerTinsList[i] : "";
                        Dictionary<string, Dictionary<string, string>> waybills_map_part = oWayBill.get_buyer_waybills(string.Join(",", selectedTypesList), seller_id, ",,1,2,-1,-2,7,", car_number, startDateParam, endDateParam, startDateParam, endDateParam, driver_tin, startDateParam, endDateParam, full_amount, waybill_number, startDateParam, endDateParam, s_user_id, comment, StartAddress, EndAddress, out errorText);
                        foreach (KeyValuePair<string, Dictionary<string, string>> map_record in waybills_map_part)
                        {
                            Dictionary<string, string> Waybill_Header = map_record.Value;
                            SAPbouiCOM.EditText WBNUM = (SAPbouiCOM.EditText)(oForm.Items.Item("WBNumber").Specific);
                            if (Waybill_Header["WAYBILL_NUMBER"] == WBNUM.Value || WBNUM.Value == "")
                            {
                                string SELLER_TIN = Waybill_Header["SELLER_TIN"];
                                string SELLER_NAME = Waybill_Header["SELLER_NAME"];
                                string WBID = Waybill_Header["ID"];

                                if (FOption != "0")
                                {
                                    string[] array_HEADER;
                                    string[][] array_GOODS, array_SUB_WAYBILLS;
                                    int returnCode = oWayBill.get_waybill(Convert.ToInt32(WBID), out array_HEADER, out array_GOODS, out array_SUB_WAYBILLS, out errorText);

                                    //ხარჯის გამწევით ფილტრი
                                    if (TranSide != "0" && array_HEADER[26] != TranSide)
                                        continue;


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
                                    taxDataRow["TRAN_COST_PAYER"] = array_HEADER[26];
                                    
                                }
                                else
                                {
                                    //ცხრილი
                                    string[] array_HEADER;
                                    string[][] array_GOODS, array_SUB_WAYBILLS;
                                    int returnCode = oWayBill.get_waybill(Convert.ToInt32(WBID), out array_HEADER, out array_GOODS, out array_SUB_WAYBILLS, out errorText);

                                    //ხარჯის გამწევით ფილტრი
                                    if (TranSide != "0" && array_HEADER[26] != TranSide)
                                        continue;

                                    int rowCounter = 1;
                                    int rowIndex = 0;

                                    foreach (string[] goodsRow in array_GOODS)
                                    {
                                        string WBBarcode = goodsRow[6] == null ? "" : Regex.Replace(goodsRow[6], @"\t|\n|\r|'", "").Trim();
                                        string WBItmName = goodsRow[1];

                                        
                                        string cardName;
                                        string Cardcode = BusinessPartners.GetCardCodeByTin(SELLER_TIN, "S", out cardName);
                                        if (Cardcode != null)
                                        {
                                            ItmCode = BDO_WaybillsJournalReceived.findItemByNameOITM(WBItmName, WBBarcode, Cardcode, out ItemName);
                                            if (ItemName == null) ItemName = "";

                                            SAPbobsCOM.Recordset CatalogEntry = BDO_BPCatalog.getCatalogEntryByBPBarcode(Cardcode, WBItmName, WBBarcode);

                                            if (CatalogEntry != null)
                                            {
                                                WBGUntCode = CatalogEntry.Fields.Item("U_BDO_UoMCod").Value;
                                                ItmCode = CatalogEntry.Fields.Item("ItemCode").Value;
                                            }
                                        }


                                        SAPbobsCOM.Recordset oRecordsetbyRSCODE = BDO_RSUoM.getUomByRSCode(ItmCode, goodsRow[2], out errorText);

                                        if (oRecordsetbyRSCODE != null)
                                        {
                                            if (WBGUntCode == "")
                                            {
                                                WBGUntCode = oRecordsetbyRSCODE.Fields.Item("UomCode").Value;
                                            }
                                        }
                                        string WbUntNmRS = string.IsNullOrEmpty(goodsRow[13]) ? oWayBill.get_waybill_unit_name_by_code(goodsRow[2]) : goodsRow[13];
                                        
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
                                        taxDataRow["TRAN_COST_PAYER"] = array_HEADER[26];
                                        taxDataRow["ItemCode"] = ItmCode;
                                        taxDataRow["W_NAME"] = WBItmName;
                                        taxDataRow["BAR_CODE"] = WBBarcode;
                                        taxDataRow["AMOUNT"] = goodsRow[5];
                                        taxDataRow["Quantity"] = goodsRow[3];
                                        taxDataRow["ItemName"] = ItemName;
                                        taxDataRow["RSUom"]= WbUntNmRS;
                                        taxDataRow["PrjCode"] = PrjCode;
                                        taxDataRow["PrjName"] = PrjName;

                                        rowCounter++;
                                        rowIndex++;
                                    }
                                }
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

                string WBNum;
                int BaseDocNum;
                string BaseDType;
                int DocEntry;
                string DocCanld;
                string ItemCode;
                decimal Amount;
                decimal Amount_Full;
                decimal RSAmount;
                decimal RSAmount_Full;
                decimal RSQuantity;
                string RS_W_NAME;
                string wbCompStat;
                string RS_STATUS;
                string TYPE;
                string START_ADDRESS;
                string END_ADDRESS;
                string CAR_NUMBER;
                string DRIVER_TIN;
                decimal TRANSPORT_COAST = 0;
                string TRAN_COST_PAYER;
                bool foundonRS;
                DateTime DeliveryDate;
                DateTime BegDate;
                DateTime ActivateDate;
                string strDELIVERY_DATE;
                string strACTIVATE_DATE;
                string strBEGIN_DATE;

                List<string> selectedStatusesList = new List<string>();
                foreach (DictionaryEntry item in selectedStatuses)
                    selectedStatusesList.Add(item.Key.ToString());

                List<string> selectedCompareStatusesList = new List<string>();
                foreach (DictionaryEntry item in selectedCompareStatuses)
                    selectedCompareStatusesList.Add(item.Key.ToString());

                while (!oRecordSet.EoF)
                {
                    WBNum = oRecordSet.Fields.Item("WBNo").Value;
                    WBNum = WBNum.Trim();

                    ItemCode = "";
                    ItemName = "";
                    if (FOption != "1")
                    {
                        ItemCode = oRecordSet.Fields.Item("ItemCode").Value;
                        ItemName = oRecordSet.Fields.Item("ItemDesc").Value;
                        PrjCode = (string)oRecordSet.Fields.Item("PrjCode").Value;
                        PrjName = (string)oRecordSet.Fields.Item("PrjName").Value;
                    }

                    BaseDType = (string)oRecordSet.Fields.Item("BaseDType").Value;
                    
                    Amount = Convert.ToDecimal(oRecordSet.Fields.Item("Gtotal").Value, CultureInfo.InvariantCulture);
                    Amount_Full = Convert.ToDecimal(oRecordSet.Fields.Item("DocTotal").Value, CultureInfo.InvariantCulture);
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
                    TRAN_COST_PAYER = "";
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
                                    foundRows = RSDataTable.Select("WAYBILL_NUMBER = '" + WBNum + "'" + " and " + "ItemCode = '" + ItemCode.Replace("'", "''") + "'");
                                else
                                    foundRows = RSDataTable.Select("WAYBILL_NUMBER = '" + WBNum + "'");

                                for (int i = 0; i < foundRows.Length; i++)
                                {
                                    foundonRS = true;
                                    foundRows[i]["RowLinked"] = "Y";

                                    if (FOption != "1")
                                    {
                                        RSAmount = FormsB1.cleanStringOfNonDigits(foundRows[i]["AMOUNT"].ToString());
                                        RSQuantity = FormsB1.cleanStringOfNonDigits(foundRows[i]["Quantity"].ToString());
                                        RS_W_NAME = foundRows[i]["W_NAME"].ToString();
                                    }

                                    RSAmount_Full = FormsB1.cleanStringOfNonDigits(foundRows[i]["FULL_AMOUNT"].ToString());
                                    RS_STATUS = foundRows[i]["STATUS"].ToString();
                                    TYPE = foundRows[i]["TYPE"].ToString();
                                    START_ADDRESS = foundRows[i]["START_ADDRESS"].ToString();
                                    END_ADDRESS = foundRows[i]["END_ADDRESS"].ToString();
                                    CAR_NUMBER = foundRows[i]["CAR_NUMBER"].ToString();
                                    DRIVER_TIN = foundRows[i]["DRIVER_TIN"].ToString();

                                    if (string.IsNullOrEmpty(foundRows[i]["TRANSPORT_COAST"].ToString()))
                                        TRANSPORT_COAST = 0;
                                    else
                                        TRANSPORT_COAST = FormsB1.cleanStringOfNonDigits(foundRows[i]["TRANSPORT_COAST"].ToString());

                                    TRAN_COST_PAYER = foundRows[i]["TRAN_COST_PAYER"].ToString();

                                    strDELIVERY_DATE = foundRows[i]["DELIVERY_DATE"].ToString();
                                    strACTIVATE_DATE = foundRows[i]["ACTIVATE_DATE"].ToString();
                                    strBEGIN_DATE = foundRows[i]["BEGIN_DATE"].ToString();
                                }
                            }
                        }

                        if (!foundonRS)
                            wbCompStat = "2";
                        else if (oRecordSet.Fields.Item("CANCELED").Value == "Y")
                            wbCompStat = "5";
                        else if (Amount_Full != RSAmount_Full)
                            wbCompStat = "3";
                        else
                            wbCompStat = "4";
                    }
                    else
                    {
                        wbCompStat = "6";
                    }

                    if (selectedCompareStatusesList.Count > 0)
                    {
                        if (!selectedCompareStatusesList.Contains(wbCompStat))
                        {
                            oRecordSet.MoveNext();
                            continue;
                        }
                    }

                    if (selectedStatusesList.Count > 0)
                    {
                        if (!selectedStatusesList.Contains(RS_STATUS))
                        {
                            oRecordSet.MoveNext();
                            continue;
                        }
                    }

                    if (FNUMBER != "" && FNUMBER != WBNum)
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
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, TRANSPORT_COAST == 0 ? "" : FormsB1.ConvertDecimalToString(TRANSPORT_COAST));
                    Sbuilder.Append("</Value></Cell>");

                    Sbuilder.Append("<Cell> <ColumnUid>WB_trnsSide</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, TRAN_COST_PAYER=="1"? BDOSResources.getTranslate("Buyer") : BDOSResources.getTranslate("Buyer"));
                    Sbuilder.Append("</Value></Cell>");
                                        

                    Sbuilder.Append("<Cell> <ColumnUid>WB_begDate</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, oRecordSet.Fields.Item("BaseDocDate").Value.ToString("yyyyMMdd") == "18991230" ? "" : oRecordSet.Fields.Item("BaseDocDate").Value.ToString("yyyyMMdd"));
                    Sbuilder.Append("</Value></Cell>");

                    BaseDocNum = (int)oRecordSet.Fields.Item("BaseDocNum").Value;
                    Sbuilder.Append("<Cell> <ColumnUid>DocNum</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, (BaseDocNum == 0 ? "" : BaseDocNum.ToString()));
                    Sbuilder.Append("</Value></Cell>");
                    
                    Sbuilder.Append("<Cell> <ColumnUid>BaseDType</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, (BaseDType == "" ? "" : BaseDType.ToString()));
                    Sbuilder.Append("</Value></Cell>");

                    DocEntry = (int)oRecordSet.Fields.Item("DocEntry").Value;
                    Sbuilder.Append("<Cell> <ColumnUid>BaseDoc</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, (DocEntry == 0 ? "" : DocEntry.ToString()));
                    Sbuilder.Append("</Value></Cell>");

                    DocCanld = (string)oRecordSet.Fields.Item("CANCELED").Value;
                    Sbuilder.Append("<Cell> <ColumnUid>DocCanld</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, DocCanld == "N" ? "No" : "YES");
                    Sbuilder.Append("</Value></Cell>");

                    Sbuilder.Append("<Cell> <ColumnUid>Whs</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, oRecordSet.Fields.Item("Whs").Value);
                    Sbuilder.Append("</Value></Cell>");

                    Sbuilder.Append("<Cell> <ColumnUid>RS_Status</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, RS_STATUS);
                    Sbuilder.Append("</Value></Cell>");

                    if (TYPE == "5") //დაბრუნება
                    {
                        RSAmount *= -1;
                        Amount *= -1;
                        Amount_Full *= -1;
                        RSAmount_Full *= -1;
                    }

                    //დეტალური როცა არის
                    if (FOption != "1")
                    {
                        Sbuilder.Append("<Cell> <ColumnUid>RSAmount</ColumnUid> <Value>");
                        Sbuilder = CommonFunctions.AppendXML(Sbuilder, FormsB1.ConvertDecimalToString(RSAmount));
                        Sbuilder.Append("</Value></Cell>");

                        Sbuilder.Append("<Cell> <ColumnUid>Gtotal</ColumnUid> <Value>");
                        Sbuilder = CommonFunctions.AppendXML(Sbuilder, FormsB1.ConvertDecimalToString(Amount));
                        Sbuilder.Append("</Value></Cell>");

                        Sbuilder.Append("<Cell> <ColumnUid>ItemCode</ColumnUid> <Value>");
                        Sbuilder = CommonFunctions.AppendXML(Sbuilder, oRecordSet.Fields.Item("ItemCode").Value);
                        Sbuilder.Append("</Value></Cell>");

                        Sbuilder.Append("<Cell> <ColumnUid>RSUom</ColumnUid> <Value>");
                        Sbuilder = CommonFunctions.AppendXML(Sbuilder, WBGUntCode);
                        Sbuilder.Append("</Value></Cell>");
                        
                        Sbuilder.Append("<Cell> <ColumnUid>ItemName</ColumnUid> <Value>");
                        Sbuilder = CommonFunctions.AppendXML(Sbuilder, ItemName);
                        Sbuilder.Append("</Value></Cell>");

                        Sbuilder.Append("<Cell> <ColumnUid>InvntryUom</ColumnUid> <Value>");
                        Sbuilder = CommonFunctions.AppendXML(Sbuilder, oRecordSet.Fields.Item("InvntryUom").Value);
                        Sbuilder.Append("</Value></Cell>");

                        Sbuilder.Append("<Cell> <ColumnUid>Quantity</ColumnUid> <Value>");
                        Sbuilder = CommonFunctions.AppendXML(Sbuilder, FormsB1.ConvertDecimalToString(Convert.ToDecimal(oRecordSet.Fields.Item("Quantity").Value, CultureInfo.InvariantCulture)));
                        Sbuilder.Append("</Value></Cell>");

                        Sbuilder.Append("<Cell> <ColumnUid>RSQuantity</ColumnUid> <Value>");
                        Sbuilder = CommonFunctions.AppendXML(Sbuilder, FormsB1.ConvertDecimalToString(RSQuantity));
                        Sbuilder.Append("</Value></Cell>");

                        Sbuilder.Append("<Cell> <ColumnUid>RSName</ColumnUid> <Value>");
                        Sbuilder = CommonFunctions.AppendXML(Sbuilder, RS_W_NAME);
                        Sbuilder.Append("</Value></Cell>");

                        Sbuilder.Append("<Cell> <ColumnUid>PrjCode</ColumnUid> <Value>");
                        Sbuilder = CommonFunctions.AppendXML(Sbuilder, PrjCode);
                        Sbuilder.Append("</Value></Cell>");

                        Sbuilder.Append("<Cell> <ColumnUid>PrjName</ColumnUid> <Value>");
                        Sbuilder = CommonFunctions.AppendXML(Sbuilder, PrjName);
                        Sbuilder.Append("</Value></Cell>");

                    }
                    else
                    {
                        Sbuilder.Append("<Cell> <ColumnUid>RSAmount</ColumnUid> <Value>");
                        Sbuilder = CommonFunctions.AppendXML(Sbuilder, FormsB1.ConvertDecimalToString(RSAmount_Full));
                        Sbuilder.Append("</Value></Cell>");

                        Sbuilder.Append("<Cell> <ColumnUid>Gtotal</ColumnUid> <Value>");
                        Sbuilder = CommonFunctions.AppendXML(Sbuilder, FormsB1.ConvertDecimalToString(Amount_Full));
                        Sbuilder.Append("</Value></Cell>");
                    }

                    Sbuilder.Append("</Row>");

                    count++;

                    oRecordSet.MoveNext();
                }

                if (RSDataTable.Rows.Count > 0 && (selectedCompareStatusesList.Count == 0 || selectedCompareStatusesList.Contains("1")))
                {
                    DataRow[] RemainingRows;
                    RemainingRows = RSDataTable.Select("RowLinked = 'N'");

                    string cardName;
                    string WAYBILL_NUMBER;
                    string cTIN;
                    string cCode;
                    string RS_st;
                    string rsType;
                    decimal fullAmount;
                    decimal transportCoast;
                    //string TRAN_COST_PAYER;
                    decimal amount;
                    decimal quantity;

                    for (int i = 0; i < RemainingRows.Length; i++)
                    {
                        rsType = RemainingRows[i]["TYPE"].ToString();
                        fullAmount = FormsB1.cleanStringOfNonDigits(RemainingRows[i]["FULL_AMOUNT"].ToString());
                        transportCoast = FormsB1.cleanStringOfNonDigits(RemainingRows[i]["TRANSPORT_COAST"].ToString());
                        TRAN_COST_PAYER = RemainingRows[i]["TRAN_COST_PAYER"].ToString();
                        amount = FormsB1.cleanStringOfNonDigits(RemainingRows[i]["AMOUNT"].ToString());
                        quantity = FormsB1.cleanStringOfNonDigits(RemainingRows[i]["Quantity"].ToString());
                        WAYBILL_NUMBER = RemainingRows[i]["WAYBILL_NUMBER"].ToString();

                        if (rsType == "5") //დაბრუნება
                        {
                            fullAmount *= -1;
                            amount *= -1;
                        }

                        RS_st = RemainingRows[i]["STATUS"].ToString();
                        //სტატუსის ფილტრი
                        if (selectedStatusesList.Count > 0 && !selectedStatusesList.Contains(RS_st))
                            continue;

                        Sbuilder.Append("<Row>");

                        //cardName = "";
                        //WAYBILL_NUMBER = "";
                        cTIN = RemainingRows[i]["TIN"].ToString().Trim();
                        cCode = BusinessPartners.GetCardCodeByTin(cTIN, "S", out cardName);

                        if (!string.IsNullOrWhiteSpace(cCode))
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
                        Sbuilder = CommonFunctions.AppendXML(Sbuilder, rsType);
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

                        Sbuilder.Append("<Cell> <ColumnUid>WB_trnsExpn</ColumnUid> <Value>");
                        Sbuilder = CommonFunctions.AppendXML(Sbuilder, transportCoast == 0 ? "" : FormsB1.ConvertDecimalToString(transportCoast));
                        Sbuilder.Append("</Value></Cell>");

                        Sbuilder.Append("<Cell> <ColumnUid>WB_trnsSide</ColumnUid> <Value>");
                        Sbuilder = CommonFunctions.AppendXML(Sbuilder, TRAN_COST_PAYER == "1" ? BDOSResources.getTranslate("Buyer") : BDOSResources.getTranslate("Seller"));
                        Sbuilder.Append("</Value></Cell>");


                        

                        Sbuilder.Append("<Cell> <ColumnUid>RS_Status</ColumnUid> <Value>");
                        Sbuilder = CommonFunctions.AppendXML(Sbuilder, RemainingRows[i]["STATUS"].ToString());
                        Sbuilder.Append("</Value></Cell>");

                        //დეტალური როცა არის
                        if (FOption != "1")
                        {
                            Sbuilder.Append("<Cell> <ColumnUid>RSAmount</ColumnUid> <Value>");
                            Sbuilder = CommonFunctions.AppendXML(Sbuilder, FormsB1.ConvertDecimalToString(amount));
                            Sbuilder.Append("</Value></Cell>");

                            Sbuilder.Append("<Cell> <ColumnUid>BarCode</ColumnUid> <Value>");
                            Sbuilder = CommonFunctions.AppendXML(Sbuilder, RemainingRows[i]["BAR_CODE"].ToString());
                            Sbuilder.Append("</Value></Cell>");

                            Sbuilder.Append("<Cell> <ColumnUid>PrjCode</ColumnUid> <Value>");
                            Sbuilder = CommonFunctions.AppendXML(Sbuilder, RemainingRows[i]["PrjCode"].ToString());
                            Sbuilder.Append("</Value></Cell>");

                            Sbuilder.Append("<Cell> <ColumnUid>RSUom</ColumnUid> <Value>");
                            Sbuilder = CommonFunctions.AppendXML(Sbuilder, RemainingRows[i]["RSUom"].ToString());
                            Sbuilder.Append("</Value></Cell>");

                            Sbuilder.Append("<Cell> <ColumnUid>ItemName</ColumnUid> <Value>");
                            Sbuilder = CommonFunctions.AppendXML(Sbuilder, RemainingRows[i]["ItemName"].ToString());
                            Sbuilder.Append("</Value></Cell>");

                            Sbuilder.Append("<Cell> <ColumnUid>RSQuantity</ColumnUid> <Value>");
                            Sbuilder = CommonFunctions.AppendXML(Sbuilder, FormsB1.ConvertDecimalToString(quantity));
                            Sbuilder.Append("</Value></Cell>");

                            Sbuilder.Append("<Cell> <ColumnUid>RSName</ColumnUid> <Value>");
                            Sbuilder = CommonFunctions.AppendXML(Sbuilder, RemainingRows[i]["W_NAME"].ToString());
                            Sbuilder.Append("</Value></Cell>");

                            Sbuilder.Append("<Cell> <ColumnUid>PrjName</ColumnUid> <Value>");
                            Sbuilder = CommonFunctions.AppendXML(Sbuilder, RemainingRows[i]["PrjName"].ToString());
                            Sbuilder.Append("</Value></Cell>");
                        }
                        else
                        {
                            Sbuilder.Append("<Cell> <ColumnUid>RSAmount</ColumnUid> <Value>");
                            Sbuilder = CommonFunctions.AppendXML(Sbuilder, FormsB1.ConvertDecimalToString(fullAmount));
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
                oGrid.Columns.Item("BaseCard").TitleObject.Caption = BDOSResources.getTranslate("CardCode");
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
                oGrid.Columns.Item("WB_trnsSide").TitleObject.Caption = BDOSResources.getTranslate("TransportationSide");
                oGrid.Columns.Item("WBLD_Doc").TitleObject.Caption = BDOSResources.getTranslate("WaybillDocEntry");

                oGrid.Columns.Item("DocDate").TitleObject.Caption = BDOSResources.getTranslate("Date");
                oGrid.Columns.Item("DocNum").TitleObject.Caption = BDOSResources.getTranslate("DocNum");
                oGrid.Columns.Item("BaseDoc").TitleObject.Caption = BDOSResources.getTranslate("BaseDocument");
                oGrid.Columns.Item("DocCanld").TitleObject.Caption = BDOSResources.getTranslate("Canceled");                    
                oGrid.Columns.Item("Whs").TitleObject.Caption = BDOSResources.getTranslate("Warehouse");
                oGrid.Columns.Item("PrjCode").TitleObject.Caption = BDOSResources.getTranslate("PrjCode");
                oGrid.Columns.Item("PrjName").TitleObject.Caption = BDOSResources.getTranslate("PrjName");
                oGrid.Columns.Item("WhsFrom").TitleObject.Caption = BDOSResources.getTranslate("FromWarehouse");

                oGrid.Columns.Item("RS_Status").TitleObject.Caption = BDOSResources.getTranslate("Status") + " RS";
                oGrid.Columns.Item("RSAmount").TitleObject.Caption = BDOSResources.getTranslate("Amount") + " RS";
                oGrid.Columns.Item("Gtotal").TitleObject.Caption = BDOSResources.getTranslate("Amount") + " " + BDOSResources.getTranslate("Document");

                oGrid.Columns.Item("ItemCode").TitleObject.Caption = BDOSResources.getTranslate("ItemCode");
                oGrid.Columns.Item("ItemName").TitleObject.Caption = BDOSResources.getTranslate("ItemName");
                oGrid.Columns.Item("InvntryUom").TitleObject.Caption = BDOSResources.getTranslate("UomEntry");
                oGrid.Columns.Item("BarCode").TitleObject.Caption = BDOSResources.getTranslate("Code");
                oGrid.Columns.Item("Quantity").TitleObject.Caption = BDOSResources.getTranslate("Quantity");
                oGrid.Columns.Item("RSUom").TitleObject.Caption = BDOSResources.getTranslate("UoM") + " RS";
                oGrid.Columns.Item("RSQuantity").TitleObject.Caption = BDOSResources.getTranslate("Quantity") + " RS";
                oGrid.Columns.Item("RSName").TitleObject.Caption = BDOSResources.getTranslate("Name") + " RS";

                //GTotal                 
                SAPbouiCOM.GridColumn oGC = oGrid.Columns.Item(27);
                oGC.Type = SAPbouiCOM.BoGridColumnType.gct_EditText;
                SAPbouiCOM.EditTextColumn oEditGC = (SAPbouiCOM.EditTextColumn)oGC;
                SAPbouiCOM.BoColumnSumType oST = oEditGC.ColumnSetting.SumType;
                oEditGC.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;

                //RSAmount                 
                SAPbouiCOM.GridColumn oAC = oGrid.Columns.Item(26);
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
                oGrid.Columns.Item(25).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;

                SAPbouiCOM.ComboBoxColumn oRS_Status = (SAPbouiCOM.ComboBoxColumn)oGrid.Columns.Item(25);
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

                //ProjectCode
                SAPbouiCOM.EditTextColumn oPrjCode = (SAPbouiCOM.EditTextColumn)oGrid.Columns.Item("PrjCode");
                oPrjCode.LinkedObjectType = "63";

                for (int i = 0; i < oColumns.Count; i++)
                {
                    oColumns.Item(i).Editable = false;
                }

                oColumns.Item(3).Visible = false;
                oColumns.Item(8).Visible = false;
                //oColumns.Item(17).Visible = false;
                oColumns.Item(24).Visible = false;
                oColumns.Item(31).Visible = false;
                oColumns.Item(36).Visible = false;
                if (FOption == "1")
                {
                    oColumns.Item(22).Visible = false;
                    oColumns.Item(23).Visible = false;
                    oColumns.Item(28).Visible = false;
                    oColumns.Item(29).Visible = false;
                    oColumns.Item(30).Visible = false;
                    oColumns.Item(31).Visible = false;
                    oColumns.Item(32).Visible = false;
                    oColumns.Item(33).Visible = false;
                    oColumns.Item(34).Visible = false;
                    oColumns.Item(35).Visible = false;
                }

                if (FOption == "1")
                    oGrid.CollapseLevel = 2;
                else
                    oGrid.CollapseLevel = 3;

                oGrid.AutoResizeColumns();

                SetGridColor(oForm, false);
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

        public static void SetGridColor(SAPbouiCOM.Form oForm, bool itemPressed)
        {
            string oSetGridColor = oForm.DataSources.UserDataSources.Item("StGrColor").ValueEx;
            if (oSetGridColor != "Y" && !itemPressed)
            {
                return;
            }

            SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("WbTable").Specific;
            string FOption = oForm.DataSources.UserDataSources.Item("option").ValueEx;

            //Compare status color
            int lastComparStat = 0;

            for (int i = 0; i < oGrid.Rows.Count; i++)
            {
                if (oSetGridColor != "Y" && itemPressed)
                {
                    oGrid.CommonSetting.SetCellFontColor(i + 1, 2, FormsB1.getLongIntRGB(0, 0, 0));
                }
                else
                {
                    if (lastComparStat != oGrid.Rows.GetParent(i))
                    {
                        bool isleaf = oGrid.Rows.IsLeaf(i);
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

        public static void collapseGrid(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("WbTable").Specific;
            oGrid.Rows.CollapseAll();
        }

        public static void expandGrid(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("WbTable").Specific;
            oGrid.Rows.ExpandAll();
        }

        public static void gridColumnSetCfl(SAPbouiCOM.Form oForm, SAPbouiCOM.ItemEvent pVal)
        {
            try
            {
                if (pVal.ColUID == "BaseDoc")
                {
                    if ((pVal.EventType == SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED && pVal.BeforeAction) || (pVal.EventType == SAPbouiCOM.BoEventTypes.et_GOT_FOCUS && !pVal.BeforeAction))
                    {
                        SAPbouiCOM.Grid oGrid = ((SAPbouiCOM.Grid)(oForm.Items.Item("WbTable").Specific));

                        int dTableRow = oGrid.GetDataTableRowIndex(pVal.Row);

                        SAPbouiCOM.DataTable oDataTable = (SAPbouiCOM.DataTable)oForm.DataSources.DataTables.Item("WbTable");
                        //string DocType = oDataTable.GetValue("TYPE", dTableRow);
                        string BaseDType = oDataTable.GetValue("BaseDType", dTableRow);

                        //UbaseDocT
                        SAPbouiCOM.EditTextColumn obaseDocT = (SAPbouiCOM.EditTextColumn)oGrid.Columns.Item("BaseDoc");

                        if (BaseDType == "18")
                            obaseDocT.LinkedObjectType = "18";
                        else if (BaseDType == "20")
                            obaseDocT.LinkedObjectType = "20";
                        else if (BaseDType == "19") //დაბრუნება
                            obaseDocT.LinkedObjectType = "19";
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

        public static void addMenus()
        {
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
            catch
            {

            }
        }

        public static string getQueryText(DateTime startDate, DateTime endDate, string cardCodes, string types, string foption)
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
	         ""OITM"".""CodeBars"",
             ""OPRJ"".""PrjName""," : " ") + @"
	         ""BASEDOCGDS"".""BaseCard"",
	         ""BASEDOCGDS"".""BaseDType"",
	         ""BASEDOCGDS"".""DocEntry"",
	         ""BASEDOCGDS"".""Whs"",
	         ""BASEDOCGDS"".""DocNum"" AS BaseDocNum,
	         ""BASEDOCGDS"".""DocDate"" AS BaseDocDate,
             ""BASEDOCGDS"".""Project"" AS PrjCode,
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
             ""OPCH"".""Project"",
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
             ""OPCH"".""Project"",
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
             ""ORPC"".""Project"",
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
             ""OPDN"".""Project"",
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
             ""OPDN"".""Project"",
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
	         ""ORPC"".""Project"",
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
	         ""ORPC"".""Project"",
	         ""ORPC"".""U_BDO_WBID"",
	         ""ORPC"".""U_BDO_WBNo"",
	         ""ORPC"".""U_BDO_WBSt"",
	         ""ORPC"".""DocTotal"" + ""ORPC"".""DpmAmnt"",
	         ""ORPC"".""CANCELED"" 
		        FROM ""RPC1"" 
		        INNER JOIN ""ORPC"" ON ""ORPC"".""DocEntry"" = ""RPC1"".""DocEntry"" 
		        WHERE ""RPC1"".""TargetType"" < 0 AND ""ORPC"".""CANCELED"" = 'N'



) AS ""BASEDOCGDS""   
	        LEFT JOIN ""OCRD"" AS ""OCRD"" ON ""BASEDOCGDS"".""BaseCard"" = ""OCRD"".""CardCode"" 
	        LEFT JOIN ""OITM"" ON ""BASEDOCGDS"".""ItemCode"" = ""OITM"".""ItemCode"" 
            LEFT JOIN ""OPRJ"" ON ""BASEDOCGDS"".""Project"" = ""OPRJ"".""PrjCode"" 
	        WHERE ((""OITM"".""ItemType"" = 'I' 
	        AND ""OITM"".""InvntItem"" = 'Y') OR ""OITM"".""ItemType"" = 'F') AND ""BASEDOCGDS"".""CANCELED"" = 'N' " +

            //NOT ""ORPC"".""U_BDO_WBID"" IN (SELECT ""OPCH"".""U_BDO_WBID"" FROM ""OPCH"" WHERE ""OPCH"".""CANCELED"" = 'Y')

            ((startDate != new DateTime(1, 1, 1)) ? @" AND ""BASEDOCGDS"".""DocDate"" >= '" + startDate.ToString("yyyyMMdd") + "' " : " ") +
              ((endDate != new DateTime(1, 1, 1)) ? @" AND ""BASEDOCGDS"".""DocDate"" <= '" + endDate.ToString("yyyyMMdd") + "' " : " ") +
              ((!string.IsNullOrEmpty(cardCodes)) ? @" AND ""BASEDOCGDS"".""BaseCard"" IN (" + cardCodes + ") " : " ") +
              ((!string.IsNullOrEmpty(types)) ? @" AND ""BASEDOCGDS"".""Type"" IN (" + types + ") " : " ") +

    //((types == "2" || types == "3" || types == "5") ? @" AND ""BASEDOCGDS"".""Type"" ='" + ((types == "5") ? "5" : "2") + "' " : " ") +

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
	         ""OITM"".""CodeBars"",
             ""OPRJ"".""PrjName""," : " ") + @"  
	         ""BASEDOCGDS"".""DocDate"",
	         ""BASEDOCGDS"".""Project"",
	         ""BASEDOCGDS"".""U_BDO_WBID"",
	         ""BASEDOCGDS"".""U_BDO_WBNo"",
	         ""BASEDOCGDS"".""U_BDO_WBSt"",
	         ""BASEDOCGDS"".""DocNum"",
	         ""BASEDOCGDS"".""CANCELED"" ";

            return tempQuery;
        }

        public static void resizeForm(SAPbouiCOM.Form oForm)
        {
            oForm.Freeze(true);
            try
            {
                oForm.Items.Item("WbTable").Width = oForm.ClientWidth;
                oForm.Items.Item("WbTable").Height = oForm.ClientHeight - 300;

                oForm.Items.Item("btnColl").Top = oForm.Items.Item("WbTable").Top + oForm.Items.Item("WbTable").Height + 15;
                oForm.Items.Item("btnExp").Top = oForm.Items.Item("WbTable").Top + oForm.Items.Item("WbTable").Height + 15;

                int left_s = 6;
                int width_s = 115;
                int left_e = left_s + width_s + 1;
                int width_e = 200;
                int left_s2 = left_e + width_e + width_e / 6 + 15;
                int left_e2 = left_s2 + width_s + 1;
                int left_s3 = left_e2 + width_e + width_e / 6 + 15;
                int left_e3 = left_s3 + width_s + 1;

                oForm.Items.Item("WBStatusSt").Left = left_s2;
                oForm.Items.Item("WBStatus").Left = left_e2;
                oForm.Items.Item("STChooseB").Left = left_e2 + width_e + 1;
                oForm.Items.Item("CompStatSt").Left = left_s2;
                oForm.Items.Item("CompStat").Left = left_e2;
                oForm.Items.Item("CSTChooseB").Left = left_e2 + width_e + 1;
                oForm.Items.Item("OptionSt").Left = left_s3;
                oForm.Items.Item("Option").Left = left_e3;
                oForm.Items.Item("WBNumberS").Left = left_s3;
                oForm.Items.Item("WBNumber").Left = left_e3;
                oForm.Items.Item("StGrColor").Left = left_e3;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        public static void chooseFromList(SAPbouiCOM.Form oForm, SAPbouiCOM.ItemEvent pVal, SAPbouiCOM.IChooseFromListEvent oCFLEvento)
        {
            try
            {
                oForm.Freeze(true);

                if (pVal.BeforeAction)
                {
                }
                else
                {
                    SAPbouiCOM.DataTable oDataTable = oCFLEvento.SelectedObjects;

                    if (oDataTable != null)
                    {
                        if (oCFLEvento.ChooseFromListUID == "BusinessPartner_CFL")
                        {
                            selectedBusinessPartners = new Hashtable();
                            List<string> cardCodes = new List<string>();
                            for (int i = 0; i < oDataTable.Rows.Count; i++)
                            {
                                selectedBusinessPartners.Add(oDataTable.GetValue("CardCode", i), oDataTable.GetValue("LicTradNum", i));
                                cardCodes.Add(oDataTable.GetValue("CardCode", i));
                            }
                            oForm.DataSources.UserDataSources.Item("ClientID").Value = string.Join(",", cardCodes);
                        }
                    }
                    else
                    {
                        selectedBusinessPartners = new Hashtable();
                        oForm.DataSources.UserDataSources.Item("ClientID").Value = "";
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        public static void uiApp_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (FormUID == "BDOSSelectValues")
            {
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && !pVal.BeforeAction)
                {
                    SAPbouiCOM.Form oModalForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                    if (pVal.ItemUID == "SelectB")
                    {
                        SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oModalForm.Items.Item("ValueMTR").Specific;
                        oMatrix.FlushToDataSource();

                        SAPbouiCOM.DataTable oDataTable = oModalForm.DataSources.DataTables.Item("ValueMTR");
                        string checkBox;
                        List<string> selectedTypesList = new List<string>();
                        List<string> selectedStatusesList = new List<string>();
                        List<string> selectedCompareStatusesList = new List<string>();

                        if (buttonType == "TPChooseB")
                            selectedTypes = new Hashtable();
                        else if (buttonType == "STChooseB")
                            selectedStatuses = new Hashtable();
                        else if (buttonType == "CSTChooseB")
                            selectedCompareStatuses = new Hashtable();

                        for (int i = 0; i < oDataTable.Rows.Count; i++)
                        {
                            checkBox = oDataTable.GetValue("CheckBox", i);
                            if (checkBox == "Y")
                            {
                                if (buttonType == "TPChooseB")
                                {
                                    selectedTypes.Add(oDataTable.GetValue("Key", i), oDataTable.GetValue("Value", i));
                                    selectedTypesList.Add(oDataTable.GetValue("Value", i));
                                }
                                else if (buttonType == "STChooseB")
                                {
                                    selectedStatuses.Add(oDataTable.GetValue("Key", i), oDataTable.GetValue("Value", i));
                                    selectedStatusesList.Add(oDataTable.GetValue("Value", i));
                                }
                                else if (buttonType == "CSTChooseB")
                                {
                                    selectedCompareStatuses.Add(oDataTable.GetValue("Key", i), oDataTable.GetValue("Value", i));
                                    selectedCompareStatusesList.Add(oDataTable.GetValue("Value", i));
                                }
                            }
                        }

                        SAPbouiCOM.Form oDocForm = Program.uiApp.Forms.GetForm("60004", 1);

                        if (buttonType == "TPChooseB")
                            oDocForm.DataSources.UserDataSources.Item("WBDocTp").Value = string.Join(",", selectedTypesList);
                        else if (buttonType == "STChooseB")
                            oDocForm.DataSources.UserDataSources.Item("WBStatus").Value = string.Join(",", selectedStatusesList);
                        else if (buttonType == "CSTChooseB")
                            oDocForm.DataSources.UserDataSources.Item("CompStat").Value = string.Join(",", selectedCompareStatusesList);

                        oModalForm.Close();
                    }
                }
            }
            else
            {
                if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
                {
                    SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_CLOSE && pVal.BeforeAction)
                    {
                        int answer = Program.uiApp.MessageBox(BDOSResources.getTranslate("CloseReceivedWaybillsAnalysis") + "?", 1, BDOSResources.getTranslate("Yes"), BDOSResources.getTranslate("No"), "");

                        if (answer != 1)
                            BubbleEvent = false;
                    }

                    else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                    {
                        SAPbouiCOM.IChooseFromListEvent oCFLEvento = (SAPbouiCOM.IChooseFromListEvent)pVal;
                        chooseFromList(oForm, pVal, oCFLEvento);
                    }

                    else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && !pVal.BeforeAction)
                    {
                        if (pVal.ItemUID == "WbFillTb")
                            updateGrid(oForm);
                        else if (pVal.ItemUID == "btnColl")
                            collapseGrid(oForm);
                        else if (pVal.ItemUID == "btnExp")
                            expandGrid(oForm);
                        else if (pVal.ItemUID == "TPChooseB" || pVal.ItemUID == "STChooseB" || pVal.ItemUID == "CSTChooseB")
                        {
                            buttonType = pVal.ItemUID;
                            createModalForm();
                        }
                    }

                    else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_COMBO_SELECT && !pVal.BeforeAction)
                    {
                        oForm.DataSources.UserDataSources.Item(pVal.ItemUID).ValueEx = oForm.Items.Item(pVal.ItemUID).Specific.Value;
                    }

                    else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED)
                    {
                        gridColumnSetCfl(oForm, pVal);
                    }

                    else if ((pVal.ItemUID == "StGrColor") && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && !pVal.BeforeAction)
                    {
                        SetGridColor(oForm, true);
                    }

                    else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_RESIZE && !pVal.BeforeAction)
                    {
                        resizeForm(oForm);
                    }
                }
            }
        }

        static void createModalForm()
        {
            string errorText;
            int formHeight = 208; //Program.uiApp.Desktop.Height / 5;
            int formWidth = 384; //Program.uiApp.Desktop.Width / 5;

            //ფორმის აუცილებელი თვისებები
            Dictionary<string, object> formProperties = new Dictionary<string, object>();
            formProperties.Add("UniqueID", "BDOSSelectValues");
            formProperties.Add("BorderStyle", SAPbouiCOM.BoFormBorderStyle.fbs_Fixed);
            formProperties.Add("Title", BDOSResources.getTranslate("SelectValues"));
            formProperties.Add("ClientWidth", formWidth);
            formProperties.Add("ClientHeight", formHeight);
            formProperties.Add("Modality", SAPbouiCOM.BoFormModality.fm_Modal);

            SAPbouiCOM.Form oForm;
            bool newForm;
            bool formExist = FormsB1.createForm(formProperties, out oForm, out newForm, out errorText);

            if (errorText != null)
            {
                return;
            }

            if (formExist)
            {
                if (newForm)
                {
                    Dictionary<string, object> formItems;
                    string itemName;

                    int top = 10;

                    formItems = new Dictionary<string, object>();
                    itemName = "ValueMTR"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
                    formItems.Add("Width", formWidth);
                    formItems.Add("Height", formHeight - 30);
                    formItems.Add("Top", top);
                    formItems.Add("UID", itemName);
                    formItems.Add("State", SAPbouiCOM.BoFormStateEnum.fs_Maximized);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("ValueMTR").Specific;
                    oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
                    SAPbouiCOM.Columns oColumns = oMatrix.Columns;
                    SAPbouiCOM.Column oColumn;

                    SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Add("ValueMTR");

                    oDataTable.Columns.Add("Key", SAPbouiCOM.BoFieldsType.ft_Integer);
                    oDataTable.Columns.Add("Value", SAPbouiCOM.BoFieldsType.ft_Text, 100);
                    oDataTable.Columns.Add("CheckBox", SAPbouiCOM.BoFieldsType.ft_Text, 1);

                    if (buttonType == "TPChooseB")
                    {
                        int i = 0;
                        oDataTable.Rows.Add();
                        oDataTable.SetValue("Key", i, i + 1);
                        oDataTable.SetValue("Value", i, BDOSResources.getTranslate("InternalShipment"));
                        oDataTable.SetValue("CheckBox", i, selectedTypes.ContainsKey(i + 1) ? "Y" : "N");
                        i++;
                        oDataTable.Rows.Add();
                        oDataTable.SetValue("Key", i, i + 1);
                        oDataTable.SetValue("Value", i, BDOSResources.getTranslate("ARInvoice") + " " + BDOSResources.getTranslate("WithTransport"));
                        oDataTable.SetValue("CheckBox", i, selectedTypes.ContainsKey(i + 1) ? "Y" : "N");
                        i++;
                        oDataTable.Rows.Add();
                        oDataTable.SetValue("Key", i, i + 1);
                        oDataTable.SetValue("Value", i, BDOSResources.getTranslate("ARInvoice") + " " + BDOSResources.getTranslate("WithoutTransport"));
                        oDataTable.SetValue("CheckBox", i, selectedTypes.ContainsKey(i + 1) ? "Y" : "N");
                        i++;
                        oDataTable.Rows.Add();
                        oDataTable.SetValue("Key", i, i + 1);
                        oDataTable.SetValue("Value", i, BDOSResources.getTranslate("Distribution"));
                        oDataTable.SetValue("CheckBox", i, selectedTypes.ContainsKey(i + 1) ? "Y" : "N");
                        i++;
                        oDataTable.Rows.Add();
                        oDataTable.SetValue("Key", i, i + 1);
                        oDataTable.SetValue("Value", i, BDOSResources.getTranslate("Return"));
                        oDataTable.SetValue("CheckBox", i, selectedTypes.ContainsKey(i + 1) ? "Y" : "N");
                        i++;
                        oDataTable.Rows.Add();
                        oDataTable.SetValue("Key", i, i + 1);
                        oDataTable.SetValue("Value", i, BDOSResources.getTranslate("SubWaybill"));
                        oDataTable.SetValue("CheckBox", i, selectedTypes.ContainsKey(i + 1) ? "Y" : "N");
                    }
                    else if (buttonType == "STChooseB")
                    {
                        int i = 0;
                        oDataTable.Rows.Add();
                        oDataTable.SetValue("Key", i, -2);
                        oDataTable.SetValue("Value", i, BDOSResources.getTranslate("Canceled"));
                        oDataTable.SetValue("CheckBox", i, selectedStatuses.ContainsKey(-2) ? "Y" : "N");
                        i++;
                        oDataTable.Rows.Add();
                        oDataTable.SetValue("Key", i, -1);
                        oDataTable.SetValue("Value", i, BDOSResources.getTranslate("deleted"));
                        oDataTable.SetValue("CheckBox", i, selectedStatuses.ContainsKey(-1) ? "Y" : "N");
                        i++;
                        oDataTable.Rows.Add();
                        oDataTable.SetValue("Key", i, 0);
                        oDataTable.SetValue("Value", i, BDOSResources.getTranslate("Saved"));
                        oDataTable.SetValue("CheckBox", i, selectedStatuses.ContainsKey(0) ? "Y" : "N");
                        i++;
                        oDataTable.Rows.Add();
                        oDataTable.SetValue("Key", i, 1);
                        oDataTable.SetValue("Value", i, BDOSResources.getTranslate("Active"));
                        oDataTable.SetValue("CheckBox", i, selectedStatuses.ContainsKey(1) ? "Y" : "N");
                        i++;
                        oDataTable.Rows.Add();
                        oDataTable.SetValue("Key", i, 2);
                        oDataTable.SetValue("Value", i, BDOSResources.getTranslate("finished"));
                        oDataTable.SetValue("CheckBox", i, selectedStatuses.ContainsKey(2) ? "Y" : "N");
                        i++;
                        oDataTable.Rows.Add();
                        oDataTable.SetValue("Key", i, 7);
                        oDataTable.SetValue("Value", i, BDOSResources.getTranslate("SentToTransporter"));
                        oDataTable.SetValue("CheckBox", i, selectedStatuses.ContainsKey(7) ? "Y" : "N");
                    }
                    else if (buttonType == "CSTChooseB")
                    {
                        int i = 0;
                        oDataTable.Rows.Add();
                        oDataTable.SetValue("Key", i, i + 1);
                        oDataTable.SetValue("Value", i, BDOSResources.getTranslate("OnlyOnSite"));
                        oDataTable.SetValue("CheckBox", i, selectedCompareStatuses.ContainsKey(i + 1) ? "Y" : "N");
                        i++;
                        oDataTable.Rows.Add();
                        oDataTable.SetValue("Key", i, i + 1);
                        oDataTable.SetValue("Value", i, BDOSResources.getTranslate("OnlyOnProgram"));
                        oDataTable.SetValue("CheckBox", i, selectedCompareStatuses.ContainsKey(i + 1) ? "Y" : "N");
                        i++;
                        oDataTable.Rows.Add();
                        oDataTable.SetValue("Key", i, i + 1);
                        oDataTable.SetValue("Value", i, BDOSResources.getTranslate("AmountsNotEqual"));
                        oDataTable.SetValue("CheckBox", i, selectedCompareStatuses.ContainsKey(i + 1) ? "Y" : "N");
                        i++;
                        oDataTable.Rows.Add();
                        oDataTable.SetValue("Key", i, i + 1);
                        oDataTable.SetValue("Value", i, BDOSResources.getTranslate("EqualAmounts"));
                        oDataTable.SetValue("CheckBox", i, selectedCompareStatuses.ContainsKey(i + 1) ? "Y" : "N");
                        i++;
                        oDataTable.Rows.Add();
                        oDataTable.SetValue("Key", i, i + 1);
                        oDataTable.SetValue("Value", i, BDOSResources.getTranslate("Linked") + " " + BDOSResources.getTranslate("Document") + " " + BDOSResources.getTranslate("NotPosted"));
                        oDataTable.SetValue("CheckBox", i, selectedCompareStatuses.ContainsKey(i + 1) ? "Y" : "N");
                        i++;
                        oDataTable.Rows.Add();
                        oDataTable.SetValue("Key", i, i + 1);
                        oDataTable.SetValue("Value", i, BDOSResources.getTranslate("SavedStatus"));
                        oDataTable.SetValue("CheckBox", i, selectedCompareStatuses.ContainsKey(i + 1) ? "Y" : "N");
                        i++;
                        oDataTable.Rows.Add();
                        oDataTable.SetValue("Key", i, i + 1);
                        oDataTable.SetValue("Value", i, BDOSResources.getTranslate("NotFoundOnSiteOnThisPeriod"));
                        oDataTable.SetValue("CheckBox", i, selectedCompareStatuses.ContainsKey(i + 1) ? "Y" : "N");
                    }

                    string UID = "ValueMTR";

                    oColumn = oColumns.Add("Key", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("Key");
                    oColumn.Editable = false;
                    oColumn.Visible = false;
                    oColumn.DataBind.Bind(UID, "Key");

                    oColumn = oColumns.Add("Value", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("Value");
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind(UID, "Value");

                    oColumn = oColumns.Add("CheckBox", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
                    oColumn.TitleObject.Caption = "";
                    oColumn.Editable = true;
                    oColumn.ValOff = "N";
                    oColumn.ValOn = "Y";
                    oColumn.DataBind.Bind(UID, "CheckBox");

                    oMatrix.Clear();
                    oMatrix.LoadFromDataSource();
                    oMatrix.AutoResizeColumns();

                    int left_s = 6;
                    int height = 15;

                    formItems = new Dictionary<string, object>();
                    itemName = "SelectB";
                    formItems.Add("Caption", "OK");
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", formHeight - 20);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }
                }
                oForm.Visible = true;
            }
        }
    }
}
