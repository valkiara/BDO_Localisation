using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BDO_Localisation_AddOn
{
    static partial class BDO_WaybillsJournalSent
    {
        public static void createForm( out string errorText)
        {
            errorText = null;

            //ფორმის აუცილებელი თვისებები
            Dictionary<string, object> formProperties = new Dictionary<string, object>();
            formProperties.Add("UniqueID", "BDO_WaybillsSentForm");
            formProperties.Add("BorderStyle", SAPbouiCOM.BoFormBorderStyle.fbs_Sizable);
            formProperties.Add("Title", BDOSResources.getTranslate("WaybillsSent"));
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
                    //ფორმის ელემენტების თვისებები
                    Dictionary<string, object> formItems = null;

                    string itemName = "";
                    int left = 6;
                    int Top = 5;

                    //რიგი 1
                    //თარიღები
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
                    formItems.Add("Caption", BDOSResources.getTranslate("To"));
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
                    formItems.Add("Caption", BDOSResources.getTranslate("DocumentType"));
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left = left + 50 + 10;

                    List<string> listValidValues = new List<string>();
                    listValidValues.Add(BDOSResources.getTranslate("Sale"));
                    listValidValues.Add(BDOSResources.getTranslate("Transfer"));
                    listValidValues.Add(BDOSResources.getTranslate("Return"));
                    listValidValues.Add(BDOSResources.getTranslate("GoodsIssue"));
                    listValidValues.Add(BDOSResources.getTranslate("Delivery"));
                    listValidValues.Add(BDOSResources.getTranslate("FixedAssetTransferDocument"));

                    formItems = new Dictionary<string, object>();
                    itemName = "WBDocTp";
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                    formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("ValidValues", listValidValues);

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

                    listValidValues = new List<string>();
                    listValidValues.Add(" ");
                    listValidValues.Add(BDOSResources.getTranslate("EmptyStatus"));
                    listValidValues.Add(BDOSResources.getTranslate("Saved"));
                    listValidValues.Add(BDOSResources.getTranslate("Active"));
                    listValidValues.Add(BDOSResources.getTranslate("finished"));
                    listValidValues.Add(BDOSResources.getTranslate("deleted"));
                    listValidValues.Add(BDOSResources.getTranslate("Canceled"));
                    listValidValues.Add(BDOSResources.getTranslate("SentToTransporter"));

                    formItems = new Dictionary<string, object>();
                    itemName = "WBStatus";
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                    formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("ValidValues", listValidValues);

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
                    formItems.Add("Caption", BDOSResources.getTranslate("BPTin"));
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left = left + 50 + 10;

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

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    //პანელი
                    Top = Top + 30;
                    left = 6;

                    formItems = new Dictionary<string, object>();
                    itemName = "WbCheck";
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 19);
                    formItems.Add("Top", Top + 1);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("Image", "HANA_CHECKBOX_CH");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left = left + 20 + 1;

                    formItems = new Dictionary<string, object>();
                    itemName = "WbUncheck";
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 19);
                    formItems.Add("Top", Top + 1);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("Image", "HANA_CHECKBOX_UH");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left = left + 20 + 5;

                    //ზედნადებების ცხრილი
                    Top = Top + 30;
                    left = 6;

                    itemName = "WBMatrix";
                    formItems = new Dictionary<string, object>();
                    formItems.Add("isDataSource", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
                    //formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_GRID);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 750);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 260);
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    Top = Top + 260 + 6;
                    left = 750 - 120;

                    //RS ოპერაციები
                    Dictionary<string, object> fieldskeysMap = new Dictionary<string, object>();
                    listValidValues = new List<string>();
                    listValidValues.Add(BDOSResources.getTranslate("RSCreate"));
                    listValidValues.Add(BDOSResources.getTranslate("RSActivation"));
                    listValidValues.Add(BDOSResources.getTranslate("RSSendToTransporter"));
                    listValidValues.Add(BDOSResources.getTranslate("RSCorrection"));
                    listValidValues.Add(BDOSResources.getTranslate("RSFinish"));
                    listValidValues.Add(BDOSResources.getTranslate("RSCancel"));
                    listValidValues.Add(BDOSResources.getTranslate("RSUpdateStatus"));

                    formItems = new Dictionary<string, object>();
                    itemName = "WbSentRS";
                    formItems.Add("Size", 20);
                    formItems.Add("Caption", BDOSResources.getTranslate("Operations"));
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON_COMBO);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 110);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 20);
                    formItems.Add("UID", itemName);
                    formItems.Add("ValidValues", listValidValues);
                    formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
                    formItems.Add("AffectsFormMode", false);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left = 750 - 120 - 115;

                    formItems = new Dictionary<string, object>();
                    itemName = "WbSentPR";
                    formItems.Add("Size", 20);
                    formItems.Add("Caption", BDOSResources.getTranslate("Print"));
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

                    SAPbouiCOM.LinkedButton oLink;
                    SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("WBMatrix").Specific));
                    SAPbouiCOM.Columns oColumns = oMatrix.Columns;
                    SAPbouiCOM.Column oColumn;

                    oColumn = oColumns.Add("LineNum", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = "#";
                    oColumn.Width = 100;
                    oColumn.Editable = false;

                    oColumn = oColumns.Add("WbChkBx", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
                    oColumn.TitleObject.Caption = "";
                    oColumn.Width = 100;
                    oColumn.Editable = true;

                    oColumn = oColumns.Add("DocType", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("DocumentType");
                    oColumn.Width = 100;
                    oColumn.Editable = false;

                    oColumn = oColumns.Add("Document", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("Document");
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oLink = oColumn.ExtendedObject;
                    oLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Invoice;

                    oColumn = oColumns.Add("DocDate", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("Date");
                    oColumn.Width = 100;
                    oColumn.Editable = false;

                    oColumn = oColumns.Add("Sum", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("Amount");
                    oColumn.Width = 100;
                    oColumn.Editable = false;

                    oColumn = oColumns.Add("CardCode", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("BPCardCode");
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oLink = oColumn.ExtendedObject;
                    oLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_BusinessPartner;

                    oColumn = oColumns.Add("CardName", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("BPName");
                    oColumn.Width = 100;
                    oColumn.Editable = false;

                    oColumn = oColumns.Add("VATno", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("BPTin");
                    oColumn.Width = 100;
                    oColumn.Editable = false;

                    oColumn = oColumns.Add("TrnsType", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("TransportType");
                    oColumn.Width = 120;
                    oColumn.Editable = false;
                    oColumn.DisplayDesc = true;
                    oColumn.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
                    oColumn.ValidValues.Add("-1", " ");
                    oColumn.ValidValues.Add("1", BDOSResources.getTranslate("Auto"));
                    oColumn.ValidValues.Add("2", BDOSResources.getTranslate("Railway"));
                    oColumn.ValidValues.Add("3", BDOSResources.getTranslate("Aviation"));
                    oColumn.ValidValues.Add("4", BDOSResources.getTranslate("other"));
                    oColumn.ValidValues.Add("5", BDOSResources.getTranslate("AutoOtherCountry"));
                    oColumn.ValidValues.Add("6", BDOSResources.getTranslate("AutoTransporter"));
                    oColumn.ValidValues.Add("7", BDOSResources.getTranslate("WithoutTransport"));

                    oColumn = oColumns.Add("Vehicle", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("Vehicle");
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oLink = oColumn.ExtendedObject;
                    oLink.LinkedObjectType = "UDO_F_BDO_VECL_D";

                    oColumn = oColumns.Add("Driver", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("Driver");
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oLink = oColumn.ExtendedObject;
                    oLink.LinkedObjectType = "UDO_F_BDO_DRVS_D";

                    oColumn = oColumns.Add("Trnsprter", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("Transporter");
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oLink = oColumn.ExtendedObject;
                    oLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_BusinessPartner;
                    //oLink.LinkedObjectType = "UDO_F_BDO_DRVS_D";

                    oColumn = oColumns.Add("WbDoc", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("WaybillDocEntry");
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oLink = oColumn.ExtendedObject;
                    oLink.LinkedObjectType = "UDO_F_BDO_WBLD_D";

                    oColumn = oColumns.Add("WbStatus", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("WaybillStatus");
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oColumn.DisplayDesc = true;
                    oColumn.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
                    oColumn.ValidValues.Add("-1", "");
                    oColumn.ValidValues.Add("1", BDOSResources.getTranslate("Saved"));
                    oColumn.ValidValues.Add("2", BDOSResources.getTranslate("Active"));
                    oColumn.ValidValues.Add("3", BDOSResources.getTranslate("finished"));
                    oColumn.ValidValues.Add("4", BDOSResources.getTranslate("deleted"));
                    oColumn.ValidValues.Add("5", BDOSResources.getTranslate("Canceled"));
                    oColumn.ValidValues.Add("6", BDOSResources.getTranslate("SentToTransporter"));

                    oColumn = oColumns.Add("WbID", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("WaybillID");
                    oColumn.Width = 100;
                    oColumn.Editable = false;

                    oColumn = oColumns.Add("WbNo", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("WaybillNumber");
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                     
                    oColumn = oColumns.Add("FromWhs", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("FromWarehouse");
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oLink = oColumn.ExtendedObject;
                    oLink.LinkedObjectType = "64";

                    oColumn = oColumns.Add("ToWhs", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("ToWarehouse");
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oLink = oColumn.ExtendedObject;
                    oLink.LinkedObjectType = "64";

                    oMatrix.Clear();
                    oMatrix.LoadFromDataSource();
                    oMatrix.AutoResizeColumns();

                    ///////////////////
                    left = 470;
                    Top = 5;

                    //რიგი 1
                    formItems = new Dictionary<string, object>();
                    itemName = "TrnsTypeSt";
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", Top + 1);
                    formItems.Add("Caption", BDOSResources.getTranslate("Transportation"));
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left = left + 100 + 10;

                    listValidValues = new List<string>();
                    listValidValues.Add(BDOSResources.getTranslate("Auto"));
                    listValidValues.Add(BDOSResources.getTranslate("Railway"));
                    listValidValues.Add(BDOSResources.getTranslate("Aviation"));
                    listValidValues.Add(BDOSResources.getTranslate("other"));
                    listValidValues.Add(BDOSResources.getTranslate("AutoOtherCountry"));
                    listValidValues.Add(BDOSResources.getTranslate("AutoTransporter"));

                    formItems = new Dictionary<string, object>();
                    itemName = "TrnsType";
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                    formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("ValidValues", listValidValues);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left = left + 100 + 5;

                    formItems = new Dictionary<string, object>();
                    itemName = "WbFTrnTp";
                    formItems.Add("Size", 20);
                    formItems.Add("Caption", BDOSResources.getTranslate("Refill"));
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 60);
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
                    left = 470;

                    formItems = new Dictionary<string, object>();
                    itemName = "VehicleSt";
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", Top + 1);
                    formItems.Add("Caption", BDOSResources.getTranslate("Vehicle"));
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left = left + 100 + 10;

                    //გადამზიდავი კომპანია
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

                    //მანქანები
                    multiSelection = false;
                    objectType = "UDO_F_BDO_VECL_D";
                    string uniqueID_VehicleCodeCFL = "VehicleCode_CFL";
                    FormsB1.addChooseFromList( oForm, multiSelection, objectType, uniqueID_VehicleCodeCFL);

                    //მძღოლები
                    multiSelection = false;
                    objectType = "UDO_F_BDO_DRVS_D";
                    string uniqueID_DriverCodeCFL = "DriverCode_CFL";
                    FormsB1.addChooseFromList( oForm, multiSelection, objectType, uniqueID_DriverCodeCFL);

                    formItems = new Dictionary<string, object>();
                    itemName = "Vehicle"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "DBDataSources");
                    formItems.Add("TableName", "@BDO_VECL");
                    formItems.Add("Alias", "Code");
                    formItems.Add("Bound", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", Top + 1);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("ChooseFromListUID", uniqueID_VehicleCodeCFL);
                    formItems.Add("ChooseFromListAlias", "Code");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left = left + 100;

                    formItems = new Dictionary<string, object>();
                    itemName = "VehicleCV"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 20);
                    formItems.Add("Top", Top - 2);
                    formItems.Add("Height", 20);
                    formItems.Add("UID", itemName);
                    formItems.Add("Image", "CHOOSE_ICON");
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("ChooseFromListUID", uniqueID_VehicleCodeCFL);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    //რიგი 3
                    Top = Top + 20;
                    left = 470;

                    formItems = new Dictionary<string, object>();
                    itemName = "DriverSt";
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", Top + 1);
                    formItems.Add("Height", 19);
                    formItems.Add("Caption", BDOSResources.getTranslate("Driver"));
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left = left + 100 + 10;

                    formItems = new Dictionary<string, object>();
                    itemName = "Driver"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "DBDataSources");
                    formItems.Add("TableName", "@BDO_DRVS");
                    formItems.Add("Alias", "Code");
                    formItems.Add("Bound", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", Top + 1);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("ChooseFromListUID", uniqueID_DriverCodeCFL);
                    formItems.Add("ChooseFromListAlias", "Code");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left = left + 100 + 10;





                    SAPbouiCOM.ComboBox oComboTrnsType = (SAPbouiCOM.ComboBox)oForm.Items.Item("TrnsType").Specific;
                    oComboTrnsType.Select("0", SAPbouiCOM.BoSearchKey.psk_ByValue);

                    SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)oForm.Items.Item("WBDocTp").Specific;
                    oCombo.Select("0", SAPbouiCOM.BoSearchKey.psk_ByValue);
                }

                oForm.Visible = true;
                oForm.Select();
            }

            GC.Collect();
        }

        public static void addMenus()
        {
            SAPbouiCOM.Menus moduleMenus;
            SAPbouiCOM.MenuItem menuItem;
            SAPbouiCOM.MenuItem fatherMenuItem;
            SAPbouiCOM.MenuCreationParams oCreationPackage;

            try
            {
                // Find the id of the menu into wich you want to add your menu item
                // ModuleMenuId = "43520"
                menuItem = Program.uiApp.Menus.Item("43520");

                // Get the menu collection of SAP Business One
                moduleMenus = menuItem.SubMenus;

                fatherMenuItem = moduleMenus.Item(3);

                // Add a pop-up menu item
                oCreationPackage = (SAPbouiCOM.MenuCreationParams)Program.uiApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                oCreationPackage.Checked = false;
                oCreationPackage.Enabled = true;
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "BDO_WBS";
                oCreationPackage.String = BDOSResources.getTranslate("WaybillsSent");
                oCreationPackage.Position = fatherMenuItem.SubMenus.Count - 1;

                menuItem = fatherMenuItem.SubMenus.AddEx(oCreationPackage);
            }
            catch
            {

            }
        }

        public static void checkUncheckWaybills(SAPbouiCOM.Form oForm, string CheckOperation, out string errorText)
        {
            errorText = null;
            try
            {
                oForm.Freeze(true);
                SAPbouiCOM.CheckBox oCheckBox;
                SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("WBMatrix").Specific));

                for (int j = 1; j <= oMatrix.RowCount; j++)
                {
                    oCheckBox = oMatrix.Columns.Item("WbChkBx").Cells.Item(j).Specific;

                    oCheckBox.Checked = (CheckOperation == "WbCheck");
                }
                oForm.Freeze(false);
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
            finally
            {
                oForm.Freeze(false);
                GC.Collect();
            }
        }

        public static void fillWaybills(  SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            try
            {
                oForm.Freeze(true);
                SAPbouiCOM.DataTable oDataTable;
                if (oForm.DataSources.DataTables.Count == 1)
                {
                    oDataTable = oForm.DataSources.DataTables.Item("WBMatrix");
                }
                else
                {
                    oDataTable = oForm.DataSources.DataTables.Add("WBMatrix");
                }
                
                string queryStr = "";

                SAPbouiCOM.ComboBox oEditTextWBDocTp = (SAPbouiCOM.ComboBox)oForm.Items.Item("WBDocTp").Specific;
                String WBDocTp = oEditTextWBDocTp.Value;

                SAPbouiCOM.ComboBox oEditTextWBStatus = (SAPbouiCOM.ComboBox)oForm.Items.Item("WBStatus").Specific;
                String WBStatus = oEditTextWBStatus.Value;

                SAPbouiCOM.EditText oEditTextClientID = (SAPbouiCOM.EditText)oForm.Items.Item("ClientID").Specific;
                String ClientID = oEditTextClientID.Value;

                SAPbouiCOM.EditText oEditTextStartDate = (SAPbouiCOM.EditText)oForm.Items.Item("StartDate").Specific;
                String StartDate = oEditTextStartDate.Value;

                SAPbouiCOM.EditText oEditTextEndDate = (SAPbouiCOM.EditText)oForm.Items.Item("EndDate").Specific;
                String EndDate = oEditTextEndDate.Value;

                if (WBDocTp == "0")
                {
                    queryStr = "SELECT " +
                                "'000000' AS \"LineNum\", " +
                                "'false' AS \"WbChkBx\", " +
                                "'Invoice' AS \"DocType\", " +
                                "\"MNTB\".\"DocEntry\" AS \"Document\", " +
                                "\"OINV\".\"DocDate\" AS \"DocDate\", " +
                                "\"OINV\".\"CardCode\" AS \"CardCode\", " +
                                "\"OCRD\".\"LicTradNum\" AS \"VATno\", " +
                                "\"OCRD\".\"CardName\" AS \"CardName\", " +
                                "\"BDO_WBLD\".\"U_tporter\" as \"Trnsprter\", " +
                                "\"BDO_WBLD\".\"U_status\" AS \"WBStatus\", " +
                                "\"BDO_WBLD\".\"U_vehicle\" AS \"Vehicle\", " +
                                "\"BDO_WBLD\".\"U_trnsType\" AS \"TrnsType\", " +
                                "\"BDO_WBLD\".\"U_drvCode\" AS \"Driver\", " +
                                "\"BDO_WBLD\".\"U_number\" AS \"WbNo\", " +
                                "\"BDO_WBLD\".\"U_wblID\" AS \"WbID\", " +
                                "\"BDO_WBLD\".\"DocEntry\" AS \"WbDoc\", " +
                                "'' AS \"VAT Number\", " +
                                "SUM(\"MNTB\".\"GTotal\") AS \"Sum\"," +
                                "'' AS \"FromWhs\", " +
                                "'' AS \"ToWhs\" " +

                                "FROM " +

                                "(SELECT " +
                                "\"INV1\".\"DocEntry\", " +
                                "\"INV1\".\"LineNum\", " +
                                "\"INV1\".\"ItemCode\", " +
                                "\"INV1\".\"Dscription\", " +

                                "\"INV1\".\"Quantity\" * \"INV1\".\"NumPerMsr\" AS \"Quantity\", " +
                                "\"INV1\".\"GTotal\", " +
                                "\"INV1\".\"VatPrcnt\", " +
                                "\"INV1\".\"LineVat\" " +
                                "FROM \"INV1\" " +
                                "INNER JOIN \"OITM\" ON \"INV1\".\"ItemCode\" = \"OITM\".\"ItemCode\" AND (\"OITM\".\"InvntItem\" = 'Y' OR \"OITM\".\"ItemType\" = 'F')" +
                                "union " +

                                "SELECT " +
                                "\"RIN1\".\"BaseEntry\", " +
                                "\"RIN1\".\"BaseLine\", " +
                                "\"RIN1\".\"ItemCode\", " +
                                "\"RIN1\".\"Dscription\", " +
                                "\"RIN1\".\"Quantity\" * (-1) * (case when \"RIN1\".\"NoInvtryMv\"='Y' THEN 0 ELSE 1 END)*\"RIN1\".\"NumPerMsr\", " +
                                "\"RIN1\".\"GTotal\" * (-1)," +
                                "\"RIN1\".\"VatPrcnt\", " +
                                "\"RIN1\".\"LineVat\" " +

                                "FROM \"RIN1\" " +

                                "INNER JOIN \"ORIN\" " +
                                "ON \"ORIN\".\"DocEntry\" = \"RIN1\".\"DocEntry\" " +

                                "INNER JOIN \"OITM\" ON \"RIN1\".\"ItemCode\" = \"OITM\".\"ItemCode\" AND (\"OITM\".\"InvntItem\" = 'Y' OR \"OITM\".\"ItemType\" = 'F') " + 

                                "WHERE \"TargetType\" < 0 AND \"ORIN\".\"U_BDO_CNTp\" <> 1 ) AS \"MNTB\" " +



                                "LEFT JOIN  \"OINV\" AS \"OINV\" " +
                                "ON \"MNTB\".\"DocEntry\" = \"OINV\".\"DocEntry\" " +

                                "LEFT JOIN  \"OCRD\" AS \"OCRD\" " +
                                "ON \"OINV\".\"CardCode\" = \"OCRD\".\"CardCode\" " +

                                "LEFT JOIN \"@BDO_WBLD\" AS \"BDO_WBLD\" " +

                                "ON \"OINV\".\"DocEntry\" = \"BDO_WBLD\".\"U_baseDoc\" AND \"BDO_WBLD\".\"U_baseDocT\" = '13' ";
                }

                if (WBDocTp == "1")
                {
                    queryStr = "SELECT " +
                                "'000000' AS \"LineNum\", " +
                                "'false' AS \"WbChkBx\", " +
                                "'Transfer' AS \"DocType\", " +
                                "\"MNTB\".\"DocEntry\" AS \"Document\", " +
                                "\"OWTR\".\"DocDate\" AS \"DocDate\", " +
                                "'' AS \"CardCode\", " +
                                "'' AS \"VATno\", " +
                                "'' AS \"CardName\", " +                          
                                "\"BDO_WBLD\".\"U_tporter\" as \"Trnsprter\", " +
                                "\"BDO_WBLD\".\"U_status\" AS \"WBStatus\", " +
                                "\"BDO_WBLD\".\"U_vehicle\" AS \"Vehicle\", " +
                                "\"BDO_WBLD\".\"U_trnsType\" AS \"TrnsType\", " +
                                "\"BDO_WBLD\".\"U_drvCode\" AS \"Driver\", " +
                                "\"BDO_WBLD\".\"U_number\" AS \"WbNo\", " +
                                "\"BDO_WBLD\".\"U_wblID\" AS \"WbID\", " +
                                "\"BDO_WBLD\".\"DocEntry\" AS \"WbDoc\", " +
                                "'' AS \"VAT Number\"," +
                                "SUM(\"MNTB\".\"GTotal\") AS \"Sum\"," +
                                "\"OWTR\".\"Filler\" AS \"FromWhs\", " +
                                "\"OWTR\".\"ToWhsCode\" AS \"ToWhs\" " +

                                "FROM " +

                                "(SELECT " +
                                "\"WTR1\".\"DocEntry\", " +
                                "\"WTR1\".\"LineNum\", " +
                                "\"WTR1\".\"ItemCode\", " +
                                "\"WTR1\".\"Dscription\", " +

                                "\"WTR1\".\"Quantity\" * \"WTR1\".\"NumPerMsr\" AS \"Quantity\", " +
                                "0 AS \"GTotal\", " +
                                "' ' AS \"VatPrcnt\", " +
                                "0 AS \"LineVat\" " +
                                "FROM \"WTR1\" ) AS \"MNTB\" " +

                                "LEFT JOIN  \"OWTR\" AS \"OWTR\" " +
                                "ON \"MNTB\".\"DocEntry\" = \"OWTR\".\"DocEntry\" " +

                                "LEFT JOIN \"@BDO_WBLD\" AS \"BDO_WBLD\" " +

                                "ON \"OWTR\".\"DocEntry\" = \"BDO_WBLD\".\"U_baseDoc\" AND \"BDO_WBLD\".\"U_baseDocT\" = '67' ";
                }

                if (WBDocTp == "2")
                {
                    queryStr = "SELECT " +
                                "'000000' AS \"LineNum\", " +
                                "'false' AS \"WbChkBx\", " +
                                "'Return' AS \"DocType\", " +
                                "\"MNTB\".\"DocEntry\" AS \"Document\", " +
                                "\"ORIN\".\"DocDate\" AS \"DocDate\", " +
                                "\"ORIN\".\"CardCode\" AS \"CardCode\", " +
                                "\"OCRD\".\"LicTradNum\" AS \"VATno\", " +
                                "\"OCRD\".\"CardName\" AS \"CardName\", " +
                                "\"BDO_WBLD\".\"U_tporter\" as \"Trnsprter\", " +
                                "\"BDO_WBLD\".\"U_status\" AS \"WBStatus\", " +
                                "\"BDO_WBLD\".\"U_vehicle\" AS \"Vehicle\", " +
                                "\"BDO_WBLD\".\"U_trnsType\" AS \"TrnsType\", " +
                                "\"BDO_WBLD\".\"U_drvCode\" AS \"Driver\", " +
                                "\"BDO_WBLD\".\"U_number\" AS \"WbNo\", " +
                                "\"BDO_WBLD\".\"U_wblID\" AS \"WbID\", " +
                                "\"BDO_WBLD\".\"DocEntry\" AS \"WbDoc\", " +
                                "'' AS \"VAT Number\", " +
                                "SUM(\"MNTB\".\"GTotal\") AS \"Sum\"," +
                                "'' AS \"FromWhs\", " +
                                "'' AS \"ToWhs\" " +
                                "FROM " +

                                "(SELECT " +
                                "\"RIN1\".\"DocEntry\", " +
                                "\"RIN1\".\"BaseLine\", " +
                                "\"RIN1\".\"ItemCode\", " +
                                "\"RIN1\".\"Dscription\", " +
                                "\"RIN1\".\"Quantity\" * \"RIN1\".\"NumPerMsr\" AS \"Quantity\", " +
                                "\"RIN1\".\"GTotal\" , " +
                                "\"RIN1\".\"VatPrcnt\", " +
                                "\"RIN1\".\"LineVat\" " +

                                "FROM \"RIN1\" " +

                                "INNER JOIN \"ORIN\" " +
                                "ON \"ORIN\".\"DocEntry\" = \"RIN1\".\"DocEntry\" " +
                                "WHERE \"TargetType\" < 0 AND \"ORIN\".\"U_BDO_CNTp\" = 1 ) AS \"MNTB\" " +

                                "INNER JOIN \"ORIN\" " +
                                "ON \"ORIN\".\"DocEntry\" = \"MNTB\".\"DocEntry\" " +

                                "LEFT JOIN  \"OCRD\" AS \"OCRD\" " +
                                "ON \"ORIN\".\"CardCode\" = \"OCRD\".\"CardCode\" " +

                                "LEFT JOIN \"@BDO_WBLD\" AS \"BDO_WBLD\" " +

                                "ON \"ORIN\".\"DocEntry\" = \"BDO_WBLD\".\"U_baseDoc\" AND \"BDO_WBLD\".\"U_baseDocT\" = '14'";
                }

                if (WBDocTp == "3")
                {
                    queryStr = "SELECT " +
                                "'000000' AS \"LineNum\", " +
                                "'false' AS \"WbChkBx\", " +
                                "'Issue' AS \"DocType\", " +
                                "\"MNTB\".\"DocEntry\" AS \"Document\", " +
                                "\"OIGE\".\"DocDate\" AS \"DocDate\", " +
                                "'' AS \"CardCode\", " +
                                "'' AS \"VATno\", " +
                                "'' AS \"CardName\", " +
                                "\"BDO_WBLD\".\"U_tporter\" as \"Trnsprter\", " +
                                "\"BDO_WBLD\".\"U_status\" AS \"WBStatus\", " +
                                "\"BDO_WBLD\".\"U_vehicle\" AS \"Vehicle\", " +
                                "\"BDO_WBLD\".\"U_trnsType\" AS \"TrnsType\", " +
                                "\"BDO_WBLD\".\"U_drvCode\" AS \"Driver\", " +
                                "\"BDO_WBLD\".\"U_number\" AS \"WbNo\", " +
                                "\"BDO_WBLD\".\"U_wblID\" AS \"WbID\", " +
                                "\"BDO_WBLD\".\"DocEntry\" AS \"WbDoc\", " +
                                "'' AS \"VAT Number\", " +
                                "SUM(\"MNTB\".\"GTotal\") AS \"Sum\"," +
                                "'' AS \"FromWhs\", " +
                                "'' AS \"ToWhs\" " +

                                "FROM " +

                                "(SELECT " +
                                "\"IGE1\".\"DocEntry\", " +
                                "\"IGE1\".\"LineNum\", " +
                                "\"IGE1\".\"ItemCode\", " +
                                "\"IGE1\".\"Dscription\", " +

                                "\"IGE1\".\"Quantity\" * \"IGE1\".\"NumPerMsr\" AS \"Quantity\", " +
                                "0 AS \"GTotal\", " +
                                "' ' AS \"VatPrcnt\", " +
                                "0 AS \"LineVat\" " +
                                "FROM \"IGE1\" WHERE \"BaseType\"= -1) AS \"MNTB\" " +

                                "LEFT JOIN  \"OIGE\" AS \"OIGE\" " +
                                "ON \"MNTB\".\"DocEntry\" = \"OIGE\".\"DocEntry\" " +

                                "LEFT JOIN \"@BDO_WBLD\" AS \"BDO_WBLD\" " +

                                "ON \"OIGE\".\"DocEntry\" = \"BDO_WBLD\".\"U_baseDoc\" AND \"BDO_WBLD\".\"U_baseDocT\" = '60' ";
                }

                if (WBDocTp == "4")
                {
                    queryStr = "SELECT " +
                                "'000000' AS \"LineNum\", " +
                                "'false' AS \"WbChkBx\", " +
                                "'Delivery' AS \"DocType\", " +
                                "\"MNTB\".\"DocEntry\" AS \"Document\", " +
                                "\"ODLN\".\"DocDate\" AS \"DocDate\", " +
                                "\"ODLN\".\"CardCode\" AS \"CardCode\", " +
                                "\"OCRD\".\"LicTradNum\" AS \"VATno\", " +
                                "\"OCRD\".\"CardName\" AS \"CardName\", " +
                                "\"BDO_WBLD\".\"U_tporter\" as \"Trnsprter\", " +
                                "\"BDO_WBLD\".\"U_status\" AS \"WBStatus\", " +
                                "\"BDO_WBLD\".\"U_vehicle\" AS \"Vehicle\", " +
                                "\"BDO_WBLD\".\"U_trnsType\" AS \"TrnsType\", " +
                                "\"BDO_WBLD\".\"U_drvCode\" AS \"Driver\", " +
                                "\"BDO_WBLD\".\"U_number\" AS \"WbNo\", " +
                                "\"BDO_WBLD\".\"U_wblID\" AS \"WbID\", " +
                                "\"BDO_WBLD\".\"DocEntry\" AS \"WbDoc\", " +
                                "'' AS \"VAT Number\", " +
                                "SUM(\"MNTB\".\"GTotal\") AS \"Sum\"," +
                                "'' AS \"FromWhs\", " +
                                "'' AS \"ToWhs\" " +

                                "FROM " +

                                "(SELECT " +
                                "\"DLN1\".\"DocEntry\", " +
                                "\"DLN1\".\"LineNum\", " +
                                "\"DLN1\".\"ItemCode\", " +
                                "\"DLN1\".\"Dscription\", " +

                                "\"DLN1\".\"Quantity\" * \"DLN1\".\"NumPerMsr\" AS \"Quantity\", " +
                                "\"DLN1\".\"GTotal\", " +
                                "\"DLN1\".\"VatPrcnt\", " +
                                "\"DLN1\".\"LineVat\" " +
                                "FROM \"DLN1\" " +
                                "INNER JOIN \"OITM\" ON \"DLN1\".\"ItemCode\" = \"OITM\".\"ItemCode\" AND (\"OITM\".\"InvntItem\" = 'Y' OR \"OITM\".\"ItemType\" = 'F')" +
                                "union " +

                                "SELECT " +
                                "\"RIN1\".\"BaseEntry\", " +
                                "\"RIN1\".\"BaseLine\", " +
                                "\"RIN1\".\"ItemCode\", " +
                                "\"RIN1\".\"Dscription\", " +
                                "\"RIN1\".\"Quantity\" * (-1) * (case when \"RIN1\".\"NoInvtryMv\"='Y' THEN 0 ELSE 1 END)*\"RIN1\".\"NumPerMsr\", " +
                                "\"RIN1\".\"GTotal\" * (-1)," +
                                "\"RIN1\".\"VatPrcnt\", " +
                                "\"RIN1\".\"LineVat\" " +

                                "FROM \"RIN1\" " +

                                "INNER JOIN \"ORIN\" " +
                                "ON \"ORIN\".\"DocEntry\" = \"RIN1\".\"DocEntry\" " +

                                "INNER JOIN \"OITM\" ON \"RIN1\".\"ItemCode\" = \"OITM\".\"ItemCode\" AND (\"OITM\".\"InvntItem\" = 'Y' OR \"OITM\".\"ItemType\" = 'F') " +

                                "WHERE \"TargetType\" < 0 AND \"ORIN\".\"U_BDO_CNTp\" <> 1 ) AS \"MNTB\" " +



                                "LEFT JOIN  \"ODLN\" AS \"ODLN\" " +
                                "ON \"MNTB\".\"DocEntry\" = \"ODLN\".\"DocEntry\" " +

                                "LEFT JOIN  \"OCRD\" AS \"OCRD\" " +
                                "ON \"ODLN\".\"CardCode\" = \"OCRD\".\"CardCode\" " +

                                "LEFT JOIN \"@BDO_WBLD\" AS \"BDO_WBLD\" " +

                                "ON \"ODLN\".\"DocEntry\" = \"BDO_WBLD\".\"U_baseDoc\" AND \"BDO_WBLD\".\"U_baseDocT\" = '15'  AND \"BDO_WBLD\".\"Canceled\" = 'N'";
                }

                if (WBDocTp == "5")
                {
                    queryStr = "SELECT " +
                                "'000000' AS \"LineNum\", " +
                                "'false' AS \"WbChkBx\", " +
                                "'Transfer' AS \"DocType\", " +
                                "\"MNTB\".\"DocEntry\" AS \"Document\", " +
                                "\"@BDOSFASTRD\".\"U_DocDate\" AS \"DocDate\", " +
                                "'' AS \"CardCode\", " +
                                "'' AS \"VATno\", " +
                                "'' AS \"CardName\", " +
                                "\"BDO_WBLD\".\"U_tporter\" as \"Trnsprter\", " +
                                "\"BDO_WBLD\".\"U_status\" AS \"WBStatus\", " +
                                "\"BDO_WBLD\".\"U_vehicle\" AS \"Vehicle\", " +
                                "\"BDO_WBLD\".\"U_trnsType\" AS \"TrnsType\", " +
                                "\"BDO_WBLD\".\"U_drvCode\" AS \"Driver\", " +
                                "\"BDO_WBLD\".\"U_number\" AS \"WbNo\", " +
                                "\"BDO_WBLD\".\"U_wblID\" AS \"WbID\", " +
                                "\"BDO_WBLD\".\"DocEntry\" AS \"WbDoc\", " +
                                "\"OCRD\".\"LicTradNum\" AS \"VAT Number\", " + 
                                "SUM(\"MNTB\".\"GTotal\") AS \"Sum\"," +
                                "'' AS \"FromWhs\", " +
                                "'' AS \"ToWhs\" " +

                                "FROM " +

                                "(SELECT " +
                                "\"@BDOSFASTR1\".\"DocEntry\", " +
                                "\"@BDOSFASTR1\".\"LineId\", " +
                                "\"@BDOSFASTR1\".\"U_ItemCode\", " +
                                "\"@BDOSFASTR1\".\"U_ItemName\", " +

                                "CASE WHEN \"@BDOSFASTR1\".\"U_Quantity\" is null or \"@BDOSFASTR1\".\"U_Quantity\" = '0' THEN '1' ELSE \"@BDOSFASTR1\".\"U_Quantity\" END AS \"Quantity\", " +
                                "0 AS \"GTotal\", " +
                                "' ' AS \"VatPrcnt\", " +
                                "0 AS \"LineVat\" " +
                                "FROM \"@BDOSFASTR1\" ) AS \"MNTB\" " +

                                "LEFT JOIN  \"@BDOSFASTRD\" AS \"@BDOSFASTRD\" " +
                                "ON \"MNTB\".\"DocEntry\" = \"@BDOSFASTRD\".\"DocEntry\" " +

                                "left join \"OCRD\"" +
                                "ON \"@BDOSFASTRD\".\"U_CardCode\" = \"OCRD\".\"CardCode\" "+

                                "LEFT JOIN \"@BDO_WBLD\" AS \"BDO_WBLD\" " +

                                "ON \"@BDOSFASTRD\".\"DocEntry\" = \"BDO_WBLD\".\"U_baseDoc\" AND \"BDO_WBLD\".\"U_baseDocT\" = 'UDO_F_BDOSFASTRD_D' ";
                }


                //ფილტრი თარიღის მიხედვით (გადასაცემია პარამეტრი)

                if (WBDocTp == "5")
                {
                    queryStr = queryStr + " WHERE \"U_DocDate\">='" + StartDate + "' AND \"U_DocDate\"<='" + EndDate + "' ";
                }
                else
                {
                    queryStr = queryStr + " WHERE \"DocDate\">='" + StartDate + "' AND \"DocDate\"<='" + EndDate + "' ";
                }
                //ფილტრი სტატუსის მიხედვით
                if (WBStatus != "0" & WBStatus != "")
                {
                    if (WBStatus == "1")
                    {
                        WBStatus = "0";
                    }
                    WBStatus = (Convert.ToInt32(WBStatus) - 1).ToString();

                    queryStr = queryStr + " AND (\"BDO_WBLD\".\"U_status\" =  " + WBStatus + (WBStatus == "-1" ? " OR \"BDO_WBLD\".\"U_status\" IS NULL) " : ") ");
                }

                //ფილტრი გსნ მიხედვით
                if (ClientID != "" & WBDocTp != "1" & WBDocTp != "3")
                {
                    queryStr = queryStr + " AND \"OCRD\".\"LicTradNum\"= '" + ClientID + "'";
                }

                if (WBDocTp == "0")
                {
                    queryStr = queryStr + " GROUP BY \"BDO_WBLD\".\"U_tporter\",\"MNTB\".\"DocEntry\", \"OINV\".\"DocDate\",\"OINV\".\"CardCode\",\"OCRD\".\"CardName\",\"OCRD\".\"LicTradNum\",\"BDO_WBLD\".\"U_status\",\"BDO_WBLD\".\"U_vehicle\",\"BDO_WBLD\".\"U_drvCode\",\"BDO_WBLD\".\"U_number\",\"BDO_WBLD\".\"U_wblID\",\"BDO_WBLD\".\"DocEntry\",\"BDO_WBLD\".\"U_trnsType\" ";
                }

                if (WBDocTp == "1")
                {
                    queryStr = queryStr + " GROUP BY \"BDO_WBLD\".\"U_tporter\",\"MNTB\".\"DocEntry\", \"OWTR\".\"DocDate\",\"CardCode\",\"CardName\",\"LicTradNum\",\"Filler\",\"ToWhsCode\",\"BDO_WBLD\".\"U_status\",\"BDO_WBLD\".\"U_vehicle\",\"BDO_WBLD\".\"U_drvCode\",\"BDO_WBLD\".\"U_number\",\"BDO_WBLD\".\"U_wblID\",\"BDO_WBLD\".\"DocEntry\",\"BDO_WBLD\".\"U_trnsType\" ";
                }

                if (WBDocTp == "2")
                {
                    queryStr = queryStr + " GROUP BY \"BDO_WBLD\".\"U_tporter\",\"MNTB\".\"DocEntry\", \"ORIN\".\"DocDate\",\"ORIN\".\"CardCode\",\"OCRD\".\"CardName\",\"OCRD\".\"LicTradNum\",\"BDO_WBLD\".\"U_status\",\"BDO_WBLD\".\"U_vehicle\",\"BDO_WBLD\".\"U_drvCode\",\"BDO_WBLD\".\"U_number\",\"BDO_WBLD\".\"U_wblID\",\"BDO_WBLD\".\"DocEntry\",\"BDO_WBLD\".\"U_trnsType\" ";
                }

                if (WBDocTp == "3")
                {
                    queryStr = queryStr + " GROUP BY \"BDO_WBLD\".\"U_tporter\",\"MNTB\".\"DocEntry\", \"OIGE\".\"DocDate\",\"CardCode\",\"CardName\",\"LicTradNum\",\"BDO_WBLD\".\"U_status\",\"BDO_WBLD\".\"U_vehicle\",\"BDO_WBLD\".\"U_drvCode\",\"BDO_WBLD\".\"U_number\",\"BDO_WBLD\".\"U_wblID\",\"BDO_WBLD\".\"DocEntry\",\"BDO_WBLD\".\"U_trnsType\" ";
                }
                if (WBDocTp == "4")
                {
                    queryStr = queryStr + " GROUP BY \"BDO_WBLD\".\"U_tporter\",\"MNTB\".\"DocEntry\", \"ODLN\".\"DocDate\",\"ODLN\".\"CardCode\",\"OCRD\".\"CardName\",\"OCRD\".\"LicTradNum\",\"BDO_WBLD\".\"U_status\",\"BDO_WBLD\".\"U_vehicle\",\"BDO_WBLD\".\"U_drvCode\",\"BDO_WBLD\".\"U_number\",\"BDO_WBLD\".\"U_wblID\",\"BDO_WBLD\".\"DocEntry\",\"BDO_WBLD\".\"U_trnsType\" ";
                }
                if (WBDocTp == "5")
                {
                    queryStr = queryStr + " GROUP BY \"BDO_WBLD\".\"U_tporter\",\"MNTB\".\"DocEntry\", \"@BDOSFASTRD\".\"U_DocDate\",\"U_CardCode\",\"OCRD\".\"CardName\", \"OCRD\".\"LicTradNum\", \"BDO_WBLD\".\"U_status\",\"BDO_WBLD\".\"U_vehicle\",\"BDO_WBLD\".\"U_drvCode\",\"BDO_WBLD\".\"U_number\",\"BDO_WBLD\".\"U_wblID\",\"BDO_WBLD\".\"DocEntry\",\"BDO_WBLD\".\"U_trnsType\" ";
                }

                //სორტირება თარიღის მიხედვით

                queryStr = queryStr + " ORDER BY \"DocDate\"";
                
                queryStr = queryStr.Replace("Invoice", BDOSResources.getTranslate("Invoice"));
                queryStr = queryStr.Replace("Delivery", BDOSResources.getTranslate("Delivery"));
                queryStr = queryStr.Replace("Return", BDOSResources.getTranslate("Return"));
                queryStr = queryStr.Replace("Transfer", BDOSResources.getTranslate("Transfer"));
                queryStr = queryStr.Replace("Issue", BDOSResources.getTranslate("Issue"));

                oDataTable.ExecuteQuery(queryStr);

                SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("WBMatrix").Specific));
                SAPbouiCOM.Columns oColumns = oMatrix.Columns;
                SAPbouiCOM.Column oColumn;

                //ცხრილის დოკუმენტის ტიპის შევსება
                oColumn = oColumns.Item("LineNum");
                oColumn.DataBind.Bind("WBMatrix", "LineNum");

                oColumn = oColumns.Item("WbChkBx");
                oColumn.DataBind.Bind("WBMatrix", "WbChkBx");

                oColumn = oColumns.Item("DocType");
                oColumn.DataBind.Bind("WBMatrix", "DocType");

                oColumn = oColumns.Item("Document");
                oColumn.DataBind.Bind("WBMatrix", "Document");

                if (WBDocTp == "0")
                {
                    SAPbouiCOM.LinkedButton oLink;
                    oLink = oColumn.ExtendedObject;
                    oLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Invoice;
                }

                if (WBDocTp == "1")
                {
                    SAPbouiCOM.LinkedButton oLink;
                    oLink = oColumn.ExtendedObject;
                    oLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_StockTransfers;
                }

                if (WBDocTp == "2")
                {
                    SAPbouiCOM.LinkedButton oLink;
                    oLink = oColumn.ExtendedObject;
                    oLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_InvoiceCreditMemo;
                }

                if (WBDocTp == "3")
                {
                    SAPbouiCOM.LinkedButton oLink;
                    oLink = oColumn.ExtendedObject;
                    oLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_GoodsIssue;
                }
                if (WBDocTp == "4")
                {
                    SAPbouiCOM.LinkedButton oLink;
                    oLink = oColumn.ExtendedObject;
                    oLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_DeliveryNotes;
                }
                if (WBDocTp == "5")
                {
                    SAPbouiCOM.LinkedButton oLink;
                    oLink = oColumn.ExtendedObject;
                    oLink.LinkedObjectType = "UDO_F_BDOSFASTRD_D";
                }

                oColumn = oColumns.Item("DocDate");
                oColumn.DataBind.Bind("WBMatrix", "DocDate");

                oColumn = oColumns.Item("Sum");
                oColumn.DataBind.Bind("WBMatrix", "Sum");

                oColumn = oColumns.Item("CardCode");
                oColumn.DataBind.Bind("WBMatrix", "CardCode");

                oColumn = oColumns.Item("CardName");
                oColumn.DataBind.Bind("WBMatrix", "CardName");

                oColumn = oColumns.Item("VATno");
                oColumn.DataBind.Bind("WBMatrix", "VATno");

                oColumn = oColumns.Item("WbStatus");
                oColumn.DataBind.Bind("WBMatrix", "WbStatus");

                oColumn = oColumns.Item("TrnsType");
                oColumn.DataBind.Bind("WBMatrix", "TrnsType");

                oColumn = oColumns.Item("Vehicle");
                oColumn.DataBind.Bind("WBMatrix", "Vehicle");

                oColumn = oColumns.Item("Driver");
                oColumn.DataBind.Bind("WBMatrix", "Driver");

                oColumn = oColumns.Item("Trnsprter");
                oColumn.DataBind.Bind("WBMatrix", "Trnsprter");

                oColumn = oColumns.Item("WbDoc");
                oColumn.DataBind.Bind("WBMatrix", "WbDoc");

                oColumn = oColumns.Item("WbNo");
                oColumn.DataBind.Bind("WBMatrix", "WbNo");

                oColumn = oColumns.Item("WbID");
                oColumn.DataBind.Bind("WBMatrix", "WbID");

                oColumn = oColumns.Item("FromWhs");
                oColumn.DataBind.Bind("WBMatrix", "FromWhs");
                oColumn.Visible = WBDocTp == "1";

                oColumn = oColumns.Item("ToWhs");
                oColumn.DataBind.Bind("WBMatrix", "ToWhs");
                oColumn.Visible = WBDocTp == "1";
                
                oMatrix.Clear();
                oMatrix.LoadFromDataSource();
                oMatrix.AutoResizeColumns();

                for (int row = 1; row <= oMatrix.RowCount; row++)
                {
                    oMatrix.Columns.Item("LineNum").Cells.Item(row).Specific.Value = row.ToString();
                    //oMatrix.Columns.Item("WbStatus").Cells.Item(row).Specific.Value = BDO_Waybills.statusAsString(oMatrix.GetCellSpecific("WbStatus", row).Value);
                    //oMatrix.Columns.Item("TrnsType").Cells.Item(row).Specific.Value = BDO_Waybills.trnsTypeAsString(oMatrix.GetCellSpecific("TrnsType", row).Value);
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
            finally
            {
                oForm.Freeze(false);
                GC.Collect();
            }
        }

        public static void TrnsTypeChange( SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            string TrnsType = oForm.Items.Item("TrnsType").Specific.Value;

            if (TrnsType == "6" || TrnsType == "-1" || TrnsType == "") //ტრანსპორტირების გარეშე
            {
                oForm.Items.Item("Vehicle").Visible = false;
                oForm.Items.Item("VehicleSt").Visible = false;
                oForm.Items.Item("VehicleCV").Visible = false;
                oForm.Items.Item("Driver").Visible = false;
                oForm.Items.Item("DriverSt").Visible = false;

            }
            else if (TrnsType == "3")
            {
                oForm.Items.Item("VehicleCV").Visible = false;
                oForm.Items.Item("Driver").Visible = false;
                oForm.Items.Item("DriverSt").Visible = false;
            }
            else if (TrnsType == "5")
            {
                oForm.Items.Item("VehicleCV").Visible = true;
                oForm.Items.Item("Driver").Visible = false;
                oForm.Items.Item("DriverSt").Visible = false;
            }
            else
            {
                oForm.Items.Item("Vehicle").Visible = true;
                oForm.Items.Item("VehicleSt").Visible = true;
                oForm.Items.Item("VehicleCV").Visible = true;
                oForm.Items.Item("Driver").Visible = true;
                oForm.Items.Item("DriverSt").Visible = true;
                //oForm.Items.Item("VehicleCV").Left = oForm.Items.Item("Vehicle").Left + oForm.Items.Item("Vehicle").Width + 5;
                //oForm.Items.Item("WbFTrnTp").Left = oForm.Items.Item("Driver").Left + oForm.Items.Item("Driver").Width + 10;
            }

            SAPbouiCOM.StaticText oItemStText = (SAPbouiCOM.StaticText)oForm.Items.Item("VehicleSt").Specific;
            SAPbouiCOM.EditText oItemEditText = (SAPbouiCOM.EditText)oForm.Items.Item("Vehicle").Specific;
            SAPbouiCOM.EditText oItemEditTextDriver = (SAPbouiCOM.EditText)oForm.Items.Item("Driver").Specific;
            SAPbouiCOM.Button oItemButt = (SAPbouiCOM.Button)oForm.Items.Item("VehicleCV").Specific;

            oItemEditText.Value = "";
            oItemEditTextDriver.Value = "";

            if (TrnsType == "5")
            {
                oItemStText.Caption = BDOSResources.getTranslate("Transporter");
                oItemEditText.ChooseFromListUID = "BusinessPartner_CFL";
                oItemEditText.ChooseFromListAlias = "CardCode";
                oItemButt.ChooseFromListUID = "BusinessPartner_CFL";
            }
            else
            {
                oItemStText.Caption = BDOSResources.getTranslate("Vehicle");
                oItemEditText.ChooseFromListUID = "VehicleCode_CFL";
                oItemEditText.ChooseFromListAlias = "Code";
                oItemButt.ChooseFromListUID = "VehicleCode_CFL";
            }
        }

        public static void ItemsValidValues( SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            SAPbouiCOM.ComboBox oTrnsTypeComboBox = (SAPbouiCOM.ComboBox)oForm.Items.Item("TrnsType").Specific;
            SAPbouiCOM.ComboBox oEditTextWBDocTp = (SAPbouiCOM.ComboBox)oForm.Items.Item("WBDocTp").Specific;
            String WBDocTp = oEditTextWBDocTp.Value;

            if (WBDocTp == "0"|| WBDocTp == "4")
            {
                if (oTrnsTypeComboBox.ValidValues.Count < 7)
                {
                    oTrnsTypeComboBox.ValidValues.Add("6", BDOSResources.getTranslate("WithoutTransport"));
                }
            }
            else
            {
                if (oTrnsTypeComboBox.ValidValues.Count == 7)
                {
                    oTrnsTypeComboBox.Select("0");
                    oTrnsTypeComboBox.ValidValues.Remove("6");
                }
            }
        }

        public static void fillTransportType(  SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            try
            {
                oForm.Freeze(true);
                SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("WBMatrix").Specific));
                SAPbouiCOM.ComboBox oEditTextTrnsType = (SAPbouiCOM.ComboBox)oForm.Items.Item("TrnsType").Specific;

                string TrnsType = oEditTextTrnsType.Value;
                string vehicleCode = "";
                string driverCode = "";

                //if (TrnsType != "6")
                //{
                vehicleCode = oForm.Items.Item("Vehicle").Specific.Value;
                driverCode = oForm.Items.Item("Driver").Specific.Value;

                //}
                //else
                //{
                //    vehicleCode = "";
                //    driverCode = "";
                //}


                if ((TrnsType == "0" || TrnsType == "4") & driverCode == "")
                {
                    Program.uiApp.MessageBox(BDOSResources.getTranslate("NecessaryToChooseVehicle"));
                    return;
                }



                for (int row = 1; row <= oMatrix.RowCount; row++)
                {
                    SAPbouiCOM.CheckBox Edtfield = oMatrix.Columns.Item("WbChkBx").Cells.Item(row).Specific;
                    bool checkedLine = (Edtfield.Checked);
                    if (checkedLine)
                    {
                        if (TrnsType == "5")
                        {
                            oMatrix.Columns.Item("Trnsprter").Cells.Item(row).Specific.Value = vehicleCode;
                            oMatrix.Columns.Item("Vehicle").Cells.Item(row).Specific.Value = "";
                            oMatrix.Columns.Item("Driver").Cells.Item(row).Specific.Value = "";
                        }
                        else
                        {
                            oMatrix.Columns.Item("Trnsprter").Cells.Item(row).Specific.Value = "";
                            oMatrix.Columns.Item("Vehicle").Cells.Item(row).Specific.Value = vehicleCode;
                            oMatrix.Columns.Item("Driver").Cells.Item(row).Specific.Value = driverCode;
                        }
                        SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)oMatrix.Columns.Item("TrnsType").Cells.Item(row).Specific;
                        oCombo.Select((Convert.ToInt32(TrnsType) + 1).ToString());

                    }
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
            finally
            {
                oForm.Freeze(false);
                GC.Collect();
            }
        }

        public static void rsOperation( SAPbouiCOM.Form oForm,  int oOperation, out string errorText)
        {
            errorText = null;

            oForm.Freeze(true);
            string objectType = "";
            bool OpSuccess = true;

            SAPbouiCOM.ComboBox oEditTextWBDocTp = (SAPbouiCOM.ComboBox)oForm.Items.Item("WBDocTp").Specific;
            String WBDocTp = oEditTextWBDocTp.Value;

            if (WBDocTp == "0")
            {
                objectType = "13";
            }
            else if (WBDocTp == "1")
            {
                objectType = "67";
            }
            else if (WBDocTp == "3")
            {
                objectType = "60";
            }
            else if (WBDocTp == "4")
            {
                objectType = "15";
            }
            else if (WBDocTp == "UDO_F_BDOSFASTRD_D")
            {
                objectType = "15";
            }
            else
            {
                objectType = "14";
            }

            if (oOperation == 0) //შექმნა
            {
                SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("WBMatrix").Specific));

                for (int row = 1; row <= oMatrix.RowCount; row++)
                {
                    SAPbouiCOM.CheckBox Edtfield = oMatrix.Columns.Item("WbChkBx").Cells.Item(row).Specific;
                    bool checkedLine = (Edtfield.Checked);
                    string WbDoc = oMatrix.GetCellSpecific("WbDoc", row).Value;

                    if (checkedLine)
                    {
                        string WbStatus = oMatrix.GetCellSpecific("WbStatus", row).Value;

                        if (WbStatus != "-1" & WbStatus != "" & WbStatus != "4" & WbStatus != "5")
                        {
                            OpSuccess = false;
                            Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("TableRow") + row.ToString() + " " + BDOSResources.getTranslate("UnableOperationForThisStatus") + " ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            continue;
                        }

                        string oBaseDocEntry = oMatrix.GetCellSpecific("Document", row).Value;

                        if (WbDoc == "")
                        {
                            int newDocEntry = 0;
                            string oVehicle = oMatrix.GetCellSpecific("Vehicle", row).Value;
                            string oDriver = oMatrix.GetCellSpecific("Driver", row).Value;
                            string TrnsType = oMatrix.GetCellSpecific("TrnsType", row).Value;
                            string Trnsprter = oMatrix.GetCellSpecific("Trnsprter", row).Value;

                            if (oVehicle == "")
                            {
                                oVehicle = null;
                            }

                            if ((TrnsType == "1" || TrnsType == "5") & oVehicle == "")
                            {
                                OpSuccess = false;
                                Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("TableRow") + row.ToString() + ": " + "აუცილებელია სატრანსპორტო საშუალების მითითება", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                continue;
                            }

                            //დოკუმენტის შექმნა პროგრამაში
                            BDO_Waybills.createDocument( objectType, Convert.ToInt32(oBaseDocEntry), oVehicle, oDriver, TrnsType, Trnsprter, out newDocEntry, out errorText);
                            oMatrix.Columns.Item("WbDoc").Cells.Item(row).Specific.Value = newDocEntry;
                            WbDoc = newDocEntry.ToString();
                        }

                        if (errorText != null)
                        {
                            OpSuccess = false;
                            Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("TableRow") + row.ToString() + ": " + errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            continue;
                        }

                        //დოკუმენტის შექმნა
                        BDO_Waybills.saveWaybill( Convert.ToInt32(WbDoc), Convert.ToInt32(oBaseDocEntry), BDOSResources.getTranslate("RSCreate"), out errorText);

                        if (errorText != null)
                        {
                            OpSuccess = false;
                            Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("TableRow") + row.ToString() + ": " + errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            continue;
                        }

                        //სტატუსების შევსება
                        Dictionary<string, string> wblDocInfo = BDO_Waybills.getWaybillDocumentInfo( Convert.ToInt32(oBaseDocEntry), objectType, out errorText);
                        oMatrix.Columns.Item("WbID").Cells.Item(row).Specific.Value = wblDocInfo["wblID"];
                        oMatrix.Columns.Item("WbNo").Cells.Item(row).Specific.Value = wblDocInfo["number"];
                        SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)oMatrix.Columns.Item("WbStatus").Cells.Item(row).Specific;
                        oCombo.Select(wblDocInfo["statusN"]);
                    }
                }
            }

            if (oOperation == 1) //აქტივაცია
            {
                SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("WBMatrix").Specific));

                for (int row = 1; row <= oMatrix.RowCount; row++)
                {
                    SAPbouiCOM.CheckBox Edtfield = oMatrix.Columns.Item("WbChkBx").Cells.Item(row).Specific;
                    bool checkedLine = (Edtfield.Checked);

                    if (checkedLine)
                    {
                        string WbStatus = oMatrix.GetCellSpecific("WbStatus", row).Value;

                        if (WbStatus != "-1" & WbStatus != "" & WbStatus != "1" & WbStatus != "4" & WbStatus != "5")
                        {
                            OpSuccess = false;
                            Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("TableRow") + row.ToString() + " " + BDOSResources.getTranslate("UnableOperationForThisStatus") + " ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            continue;
                        }

                        string WbDoc = oMatrix.GetCellSpecific("WbDoc", row).Value;
                        string oBaseDocEntry = oMatrix.GetCellSpecific("Document", row).Value;
                        string oVehicle = oMatrix.GetCellSpecific("Vehicle", row).Value;
                        string oDriver = oMatrix.GetCellSpecific("Driver", row).Value;
                        string TrnsType = oMatrix.GetCellSpecific("TrnsType", row).Value;
                        string Trnsprter = oMatrix.GetCellSpecific("Trnsprter", row).Value;
                        int newDocEntry = 0;

                        if (oVehicle == "")
                        {
                            oVehicle = null;
                        }

                        if (WbDoc == "")
                        {
                            BDO_Waybills.createDocument( objectType, Convert.ToInt32(oBaseDocEntry), oVehicle, oDriver, TrnsType, Trnsprter, out newDocEntry, out errorText);

                            if (errorText != null)
                            {
                                OpSuccess = false;
                                Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("TableRow") + row.ToString() + ": " + errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                continue;
                            }

                            oMatrix.Columns.Item("WbDoc").Cells.Item(row).Specific.Value = newDocEntry;
                            WbDoc = newDocEntry.ToString();
                        }

                        //დოკუმენტის შექმნა
                        BDO_Waybills.saveWaybill( Convert.ToInt32(WbDoc), Convert.ToInt32(oBaseDocEntry), BDOSResources.getTranslate("RSActivation"), out errorText);

                        if (errorText != null)
                        {
                            OpSuccess = false;
                            Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("TableRow") + row.ToString() + ": " + errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            continue;
                        }

                        //სტატუსების შევსება
                        Dictionary<string, string> wblDocInfo = BDO_Waybills.getWaybillDocumentInfo( Convert.ToInt32(oBaseDocEntry), objectType, out errorText);
                        oMatrix.Columns.Item("WbID").Cells.Item(row).Specific.Value = wblDocInfo["wblID"];
                        oMatrix.Columns.Item("WbNo").Cells.Item(row).Specific.Value = wblDocInfo["number"];
                        SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)oMatrix.Columns.Item("WbStatus").Cells.Item(row).Specific;
                        oCombo.Select(wblDocInfo["statusN"]);
                    }
                }
            }

            if (oOperation == 2) //გადამზიდავთან გადაგზავნა
            {
                SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("WBMatrix").Specific));

                for (int row = 1; row <= oMatrix.RowCount; row++)
                {
                    SAPbouiCOM.CheckBox Edtfield = oMatrix.Columns.Item("WbChkBx").Cells.Item(row).Specific;
                    bool checkedLine = (Edtfield.Checked);

                    if (checkedLine)
                    {
                        string WbStatus = oMatrix.GetCellSpecific("WbStatus", row).Value;

                        if (WbStatus != "-1" & WbStatus != "" & WbStatus != "1" & WbStatus != "4" & WbStatus != "5")
                        {
                            OpSuccess = false;
                            Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("TableRow") + row.ToString() + " " + BDOSResources.getTranslate("UnableOperationForThisStatus") + " ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            continue;
                        }

                        string WbDoc = oMatrix.GetCellSpecific("WbDoc", row).Value;
                        string oBaseDocEntry = oMatrix.GetCellSpecific("Document", row).Value;
                        string oVehicle = oMatrix.GetCellSpecific("Vehicle", row).Value;

                        //დოკუმენტის შექმნა
                        BDO_Waybills.saveWaybill( Convert.ToInt32(WbDoc), Convert.ToInt32(oBaseDocEntry), BDOSResources.getTranslate("RSSendToTransporter"), out errorText);

                        if (errorText != null)
                        {
                            OpSuccess = false;
                            Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("TableRow") + row.ToString() + ": " + errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            continue;
                        }

                        //სტატუსების შევსება
                        Dictionary<string, string> wblDocInfo = BDO_Waybills.getWaybillDocumentInfo( Convert.ToInt32(oBaseDocEntry), objectType, out errorText);
                        oMatrix.Columns.Item("WbID").Cells.Item(row).Specific.Value = wblDocInfo["wblID"];
                        oMatrix.Columns.Item("WbNo").Cells.Item(row).Specific.Value = wblDocInfo["number"];
                        SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)oMatrix.Columns.Item("WbStatus").Cells.Item(row).Specific;
                        oCombo.Select(wblDocInfo["statusN"]);
                    }
                }
            }

            if (oOperation == 3) //კორექტირება
            {
                SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("WBMatrix").Specific));

                for (int row = 1; row <= oMatrix.RowCount; row++)
                {
                    SAPbouiCOM.CheckBox Edtfield = oMatrix.Columns.Item("WbChkBx").Cells.Item(row).Specific;
                    bool checkedLine = (Edtfield.Checked);

                    if (checkedLine)
                    {
                        string WbStatus = oMatrix.GetCellSpecific("WbStatus", row).Value;

                        if (WbStatus != "2" & WbStatus != "3")
                        {
                            OpSuccess = false;
                            Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("TableRow") + row.ToString() + " " + BDOSResources.getTranslate("UnableOperationForThisStatus") + " ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            continue;
                        }

                        string WbDoc = oMatrix.GetCellSpecific("WbDoc", row).Value;
                        string oBaseDocEntry = oMatrix.GetCellSpecific("Document", row).Value;
                        string oVehicle = oMatrix.GetCellSpecific("Vehicle", row).Value;

                        //დოკუმენტის შექმნა
                        BDO_Waybills.saveWaybill( Convert.ToInt32(WbDoc), Convert.ToInt32(oBaseDocEntry), BDOSResources.getTranslate("RSCorrection"), out errorText);

                        if (errorText != null)
                        {
                            OpSuccess = false;
                            Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("TableRow") + row.ToString() + ": " + errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            continue;
                        }

                        //სტატუსების შევსება
                        Dictionary<string, string> wblDocInfo = BDO_Waybills.getWaybillDocumentInfo( Convert.ToInt32(oBaseDocEntry), objectType, out errorText);
                        oMatrix.Columns.Item("WbID").Cells.Item(row).Specific.Value = wblDocInfo["wblID"];
                        oMatrix.Columns.Item("WbNo").Cells.Item(row).Specific.Value = wblDocInfo["number"];
                        SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)oMatrix.Columns.Item("WbStatus").Cells.Item(row).Specific;
                        oCombo.Select(wblDocInfo["statusN"]);
                    }
                }
            }

            if (oOperation == 4)//დასრულება
            {
                SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("WBMatrix").Specific));

                for (int row = 1; row <= oMatrix.RowCount; row++)
                {
                    SAPbouiCOM.CheckBox Edtfield = oMatrix.Columns.Item("WbChkBx").Cells.Item(row).Specific;
                    bool checkedLine = (Edtfield.Checked);

                    if (checkedLine)
                    {
                        string WbStatus = oMatrix.GetCellSpecific("WbStatus", row).Value;

                        if (WbStatus != "2" & WbStatus != "6")
                        {
                            OpSuccess = false;
                            Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("TableRow") + row.ToString() + " " + BDOSResources.getTranslate("UnableOperationForThisStatus") + " ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            continue;
                        }

                        string WbDoc = oMatrix.GetCellSpecific("WbDoc", row).Value;
                        string oBaseDocEntry = oMatrix.GetCellSpecific("Document", row).Value;
                        string oVehicle = oMatrix.GetCellSpecific("Vehicle", row).Value;

                        //დოკუმენტის შექმნა
                        BDO_Waybills.closeWaybill( Convert.ToInt32(WbDoc), Convert.ToInt32(oBaseDocEntry), out errorText);

                        if (errorText != null)
                        {
                            OpSuccess = false;
                            Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("TableRow") + row.ToString() + ": " + errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            continue;
                        }

                        //სტატუსების შევსება
                        Dictionary<string, string> wblDocInfo = BDO_Waybills.getWaybillDocumentInfo( Convert.ToInt32(oBaseDocEntry), objectType, out errorText);
                        oMatrix.Columns.Item("WbID").Cells.Item(row).Specific.Value = wblDocInfo["wblID"];
                        oMatrix.Columns.Item("WbNo").Cells.Item(row).Specific.Value = wblDocInfo["number"];
                        SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)oMatrix.Columns.Item("WbStatus").Cells.Item(row).Specific;
                        oCombo.Select(wblDocInfo["statusN"]);

                    }
                }
            }

            if (oOperation == 5)//გაუქმება
            {
                SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("WBMatrix").Specific));

                for (int row = 1; row <= oMatrix.RowCount; row++)
                {
                    SAPbouiCOM.CheckBox Edtfield = oMatrix.Columns.Item("WbChkBx").Cells.Item(row).Specific;
                    bool checkedLine = (Edtfield.Checked);

                    if (checkedLine)
                    {
                        string WbStatus = oMatrix.GetCellSpecific("WbStatus", row).Value;

                        if (WbStatus != "2" & WbStatus != "3")
                        {
                            OpSuccess = false;
                            Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("TableRow") + row.ToString() + " " + BDOSResources.getTranslate("UnableOperationForThisStatus") + " ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            continue;
                        }

                        string WbDoc = oMatrix.GetCellSpecific("WbDoc", row).Value;
                        string oBaseDocEntry = oMatrix.GetCellSpecific("Document", row).Value;
                        string oVehicle = oMatrix.GetCellSpecific("Vehicle", row).Value;

                        //დოკუმენტის შექმნა
                        BDO_Waybills.refWaybill( Convert.ToInt32(WbDoc), Convert.ToInt32(oBaseDocEntry), out errorText);

                        if (errorText != null)
                        {
                            OpSuccess = false;
                            Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("TableRow") + row.ToString() + ": " + errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            continue;
                        }

                        //სტატუსების შევსება
                        Dictionary<string, string> wblDocInfo = BDO_Waybills.getWaybillDocumentInfo( Convert.ToInt32(oBaseDocEntry), objectType, out errorText);
                        oMatrix.Columns.Item("WbID").Cells.Item(row).Specific.Value = wblDocInfo["wblID"];
                        oMatrix.Columns.Item("WbNo").Cells.Item(row).Specific.Value = wblDocInfo["number"];
                        SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)oMatrix.Columns.Item("WbStatus").Cells.Item(row).Specific;
                        oCombo.Select(wblDocInfo["statusN"]);
                    }
                }
            }

            if (oOperation == 6)//სტატუსების განახლება
            {
                SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("WBMatrix").Specific));

                for (int row = 1; row <= oMatrix.RowCount; row++)
                {
                    SAPbouiCOM.CheckBox Edtfield = oMatrix.Columns.Item("WbChkBx").Cells.Item(row).Specific;
                    bool checkedLine = (Edtfield.Checked);

                    if (checkedLine)
                    {
                        string WbDoc = oMatrix.GetCellSpecific("WbDoc", row).Value;
                        string oBaseDocEntry = oMatrix.GetCellSpecific("Document", row).Value;
                        string oVehicle = oMatrix.GetCellSpecific("Vehicle", row).Value;

                        //დოკუმენტის შექმნა
                        BDO_Waybills.getWaybill( Convert.ToInt32(WbDoc), Convert.ToInt32(oBaseDocEntry), out errorText);

                        if (errorText != null)
                        {
                            OpSuccess = false;
                            Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("TableRow") + row.ToString() + ": " + errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            continue;
                        }

                        //სტატუსების შევსება
                        Dictionary<string, string> wblDocInfo = BDO_Waybills.getWaybillDocumentInfo( Convert.ToInt32(oBaseDocEntry), objectType, out errorText);
                        oMatrix.Columns.Item("WbID").Cells.Item(row).Specific.Value = wblDocInfo["wblID"];
                        oMatrix.Columns.Item("WbNo").Cells.Item(row).Specific.Value = wblDocInfo["number"];
                        SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)oMatrix.Columns.Item("WbStatus").Cells.Item(row).Specific;
                        oCombo.Select(wblDocInfo["statusN"]);
                    }
                }
            }

            if (OpSuccess == false)
            {
                Program.uiApp.MessageBox(BDOSResources.getTranslate("Operation") + " " + BDOSResources.getTranslate("EndsWithErrorCheckMessageLog"));
            }

            oForm.Freeze(false);
        }

        public static void chooseFromList( SAPbouiCOM.Form oForm, SAPbouiCOM.IChooseFromListEvent oCFLEvento, out string errorText)
        {
            errorText = null;
            try
            {
                string sCFL_ID = oCFLEvento.ChooseFromListUID;
                SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(sCFL_ID);

                SAPbouiCOM.DataTable oDataTable = null;
                oDataTable = oCFLEvento.SelectedObjects;

                if (oDataTable != null)
                {
                    if (sCFL_ID == "VehicleCode_CFL")
                    {
                        string vehicleCode = Convert.ToString(oDataTable.GetValue("Code", 0));

                        SAPbouiCOM.EditText oVehicle = oForm.Items.Item("Vehicle").Specific;
                        oVehicle.Value = vehicleCode;

                        SAPbobsCOM.UserTable oUserTable = null;
                        oUserTable = Program.oCompany.UserTables.Item("BDO_VECL");
                        oUserTable.GetByKey(vehicleCode);
                        string driverCode = oUserTable.UserFields.Fields.Item("U_drvCode").Value;

                        SAPbouiCOM.EditText oDriver = oForm.Items.Item("Driver").Specific;
                        oDriver.Value = driverCode;
                    }

                    if (sCFL_ID == "DriverCode_CFL")
                    {
                        string DriverCode = Convert.ToString(oDataTable.GetValue("Code", 0));

                        SAPbouiCOM.EditText oDriver = oForm.Items.Item("Driver").Specific;
                        oDriver.Value = DriverCode;
                    }

                    if (sCFL_ID == "BusinessPartner_CFL")
                    {
                        string vehicleCode = Convert.ToString(oDataTable.GetValue("CardCode", 0));

                        SAPbouiCOM.EditText oVehicle = oForm.Items.Item("Vehicle").Specific;
                        oVehicle.Value = vehicleCode;
                    }
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

        public static void printWaybill(  SAPbouiCOM.Form oForm)
        {
            string errorText = null;
            bool opSuccess = true;

            string addonName = "BDOS Localisation AddOn";
            string addonFormType = "UDO_FT_UDO_F_BDO_WBLD_D";
            string defaultReportLayoutCode = CrystalReports.getDefaultReportLayoutCode( addonName, addonFormType, out errorText);

            if (string.IsNullOrEmpty(errorText) == false)
            {
                Program.uiApp.StatusBar.SetSystemMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                opSuccess = false;
                return;
            }
            if (string.IsNullOrEmpty(defaultReportLayoutCode) == true)
            {
                Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("NotSetDefaultReportForWayBillDocument"), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                opSuccess = false;
                return;
            }

            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("WBMatrix").Specific));

            for (int row = 1; row <= oMatrix.RowCount; row++)
            {
                SAPbouiCOM.CheckBox Edtfield = oMatrix.Columns.Item("WbChkBx").Cells.Item(row).Specific;
                bool checkedLine = (Edtfield.Checked);
                string waybillDoc = oMatrix.GetCellSpecific("WbDoc", row).Value;

                if (string.IsNullOrEmpty(waybillDoc) == false)
                {
                    int docEntry = Convert.ToInt32(waybillDoc);
                    if (checkedLine)
                    {
                        CrystalReports.printCrystalReport( defaultReportLayoutCode, docEntry, out errorText);
                        if (string.IsNullOrEmpty(errorText) == false)
                        {
                            opSuccess = false;
                            Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("TableRow") + row.ToString() + ": " + errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            continue;
                        }
                    }
                }
            }

            if (opSuccess == false)
            {
                Program.uiApp.MessageBox(BDOSResources.getTranslate("Operation") + " " + BDOSResources.getTranslate("EndsWithErrorCheckMessageLog"));
            }
            else
            {
                Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("Operation") + " " + BDOSResources.getTranslate("DoneSuccessfully"), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
        }

        public static void uiApp_ItemEvent(  string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (pVal.ItemUID == "WbFTrnTp" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK)
                {
                    if (pVal.BeforeAction == false)
                    {
                        BDO_WaybillsJournalSent.fillTransportType( oForm, out errorText);
                    }
                }

                if (pVal.ItemUID == "TrnsType")
                {
                    if (pVal.BeforeAction == false & pVal.EventType == SAPbouiCOM.BoEventTypes.et_COMBO_SELECT)
                    {
                        BDO_WaybillsJournalSent.TrnsTypeChange( oForm, out errorText);
                    }
                }

                if (pVal.ItemUID == "WBDocTp")
                {
                    if (pVal.BeforeAction == false & pVal.EventType == SAPbouiCOM.BoEventTypes.et_COMBO_SELECT)
                    {
                        BDO_WaybillsJournalSent.ItemsValidValues( oForm, out errorText);
                    }
                }

                if ((pVal.ItemUID == "VehicleCV" || pVal.ItemUID == "Driver") & pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                {
                    if (pVal.BeforeAction == false)
                    {
                        SAPbouiCOM.IChooseFromListEvent oCFLEvento = null;
                        oCFLEvento = ((SAPbouiCOM.IChooseFromListEvent)(pVal));

                        BDO_WaybillsJournalSent.chooseFromList( oForm, oCFLEvento, out errorText);
                    }
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_COMBO_SELECT & pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "WbSentRS")
                    {
                        SAPbouiCOM.ButtonCombo oWbSentRS = ((SAPbouiCOM.ButtonCombo)(oForm.Items.Item("WbSentRS").Specific));
                        oWbSentRS.Caption = BDOSResources.getTranslate("Operations");
                        int oOperation = pVal.PopUpIndicator;
                        BDO_WaybillsJournalSent.rsOperation( oForm, oOperation, out errorText);
                    }
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK & pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "WbFillTb")
                    {
                        BDO_WaybillsJournalSent.fillWaybills( oForm, out errorText);
                    }

                    if (pVal.ItemUID == "WbCheck" || pVal.ItemUID == "WbUncheck")
                    {
                        BDO_WaybillsJournalSent.checkUncheckWaybills(oForm, pVal.ItemUID, out errorText);
                    }

                    if (pVal.ItemUID == "WbSentPR")
                    {
                        BDO_WaybillsJournalSent.printWaybill( oForm);
                    }
                }
            }
        }
    }
}