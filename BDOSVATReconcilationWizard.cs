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
    class BDOSVATReconcilationWizard
    {
        static DataTable ItemsDT = null;
        

        public static void createForm(  out string errorText)
        {
            errorText = null;
            Dictionary<string, object> formItems;
            string itemName;
            SAPbouiCOM.Columns oColumns;
            SAPbouiCOM.Column oColumn;

            SAPbouiCOM.DataTable oDataTable;

            bool multiSelection;

            int left_s = 5;
            
            int top = 10;
            int height = 15;

            //ფორმის აუცილებელი თვისებები
            Dictionary<string, object> formProperties = new Dictionary<string, object>();
            formProperties.Add("UniqueID", "BDOSReconWizz");
            formProperties.Add("BorderStyle", SAPbouiCOM.BoFormBorderStyle.fbs_Sizable);
            formProperties.Add("Title", BDOSResources.getTranslate("VATReconcilationWizard"));
            formProperties.Add("Left", 558);
            formProperties.Add("ClientWidth", 600);
            formProperties.Add("Top", 335);
            formProperties.Add("ClientHeight", 600);

            SAPbouiCOM.Form oForm;
            bool newForm;
            bool formExist = FormsB1.createForm( formProperties, out oForm, out newForm, out errorText);

            if (formExist == true)
            {
                if (newForm)
                {
                    multiSelection = false;
                    string objectTypeCardCode = "2"; //SAPbouiCOM.BoLinkedObject.lf_BusinessPartner, Business Partner object 
                    string uniqueID_lf_BusinessPartnerCFL = "BusinessPartner_CFL";
                    FormsB1.addChooseFromList( oForm, multiSelection, objectTypeCardCode, uniqueID_lf_BusinessPartnerCFL);

                    //პირობის დადება ბიზნესპარტნიორის არჩევის სიაზე
                    SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(uniqueID_lf_BusinessPartnerCFL);
                    SAPbouiCOM.Conditions oCons = oCFL.GetConditions();
                    SAPbouiCOM.Condition oCon = oCons.Add();
                    oCon.Alias = "CardType";
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCon.CondVal = "C"; //მყიდველი
                    oCFL.SetConditions(oCons);



                    formItems = new Dictionary<string, object>();
                    itemName = "FilterSt";
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", 130);
                    formItems.Add("Top", top);
                    formItems.Add("Caption", BDOSResources.getTranslate("Filter"));
                    formItems.Add("UID", itemName);
                    formItems.Add("TextStyle", 4);
                    formItems.Add("FontSize", 10);
                    formItems.Add("FromPane", 0);
                    formItems.Add("ToPane", 0);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    //formItems = new Dictionary<string, object>();
                    //itemName = "ValuesSt";
                    //formItems.Add("Size", 20);
                    //formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    //formItems.Add("Left", left_s1);
                    //formItems.Add("Width", 130);
                    //formItems.Add("Top", top);
                    //formItems.Add("Caption", BDOSResources.getTranslate("DefaultValues"));
                    //formItems.Add("UID", itemName);
                    //formItems.Add("TextStyle", 4);
                    //formItems.Add("FontSize", 10);
                    //formItems.Add("FromPane", 0);
                    //formItems.Add("ToPane", 0);

                    //FormsB1.createFormItem(oForm, formItems, out errorText);
                    //if (errorText != null)
                    //{
                    //    return;
                    //}

                    top = top + height + 1;

                    formItems = new Dictionary<string, object>();
                    itemName = "DocPsDtS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("DocumentPostingDate"));
                    formItems.Add("FromPane", 0);
                    formItems.Add("ToPane", 0);

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
                    formItems.Add("Left", left_s + 5 + 100);
                    formItems.Add("Width", 150);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("ValueEx", new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1).AddMonths(1).AddDays(-1).ToString("yyyyMMdd"));
                    formItems.Add("FromPane", 0);
                    formItems.Add("ToPane", 0);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    //formItems = new Dictionary<string, object>();
                    //itemName = "VatGrpS"; //10 characters
                    //formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    //formItems.Add("Left", left_s1);
                    //formItems.Add("Width", 100);
                    //formItems.Add("Top", top);
                    //formItems.Add("Height", height);
                    //formItems.Add("UID", itemName);
                    //formItems.Add("Caption", BDOSResources.getTranslate("VatGroup"));
                    //formItems.Add("FromPane", 0);
                    //formItems.Add("ToPane", 0);

                    //FormsB1.createFormItem(oForm, formItems, out errorText);
                    //if (errorText != null)
                    //{
                    //    return;
                    //}

                    //formItems = new Dictionary<string, object>();
                    //itemName = "VatGrp"; //10 characters
                    //formItems.Add("isDataSource", true);
                    //formItems.Add("DataSource", "UserDataSources");
                    //formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    //formItems.Add("TableName", "");
                    //formItems.Add("Length", 20);
                    //formItems.Add("Size", 20);
                    //formItems.Add("Alias", "VatGrp");
                    //formItems.Add("Bound", true);
                    //formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                    //formItems.Add("Left", left_s1 + 5 + 100);
                    //formItems.Add("Width", 150);
                    //formItems.Add("Top", top);
                    //formItems.Add("Height", height);
                    //formItems.Add("UID", itemName);
                    //formItems.Add("DisplayDesc", true);
                    //formItems.Add("FromPane", 0);
                    //formItems.Add("ToPane", 0);

                    //FormsB1.createFormItem(oForm, formItems, out errorText);
                    //if (errorText != null)
                    //{
                    //    return;
                    //}

                    //SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    //string query = "select * " +
                    //"FROM  \"OVTG\" " +
                    //"WHERE \"Category\"='O'";
                    //oRecordSet.DoQuery(query);

                    //while (!oRecordSet.EoF)
                    //{
                    //    oForm.Items.Item("VatGrp").Specific.ValidValues.Add(oRecordSet.Fields.Item("Code").Value, oRecordSet.Fields.Item("Code").Value);
                    //    oRecordSet.MoveNext();
                    //}

                    //oForm.Items.Item("VatGrp").Specific.Select(CommonFunctions.getOADM( "DfSVatItem"));



                    top = top + height + 1;

                    formItems = new Dictionary<string, object>();
                    itemName = "BPCodeS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", 100 - 20);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("BPCardCode"));
                    formItems.Add("FromPane", 0);
                    formItems.Add("ToPane", 0);

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
                    formItems.Add("Left", left_s + 5 + 100);
                    formItems.Add("Width", 150);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("ChooseFromListUID", uniqueID_lf_BusinessPartnerCFL);
                    formItems.Add("ChooseFromListAlias", "CardCode");
                    formItems.Add("FromPane", 0);
                    formItems.Add("ToPane", 0);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "BPCodeLB"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    formItems.Add("Left", left_s + 5 + 80);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("LinkTo", "BPCode");
                    formItems.Add("LinkedObjectType", objectTypeCardCode);
                    formItems.Add("FromPane", 0);
                    formItems.Add("ToPane", 0);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    //formItems = new Dictionary<string, object>();
                    //itemName = "DescrptS"; //10 characters
                    //formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    //formItems.Add("Left", left_s1);
                    //formItems.Add("Width", 100);
                    //formItems.Add("Top", top);
                    //formItems.Add("Height", height);
                    //formItems.Add("UID", itemName);
                    //formItems.Add("Caption", BDOSResources.getTranslate("Description"));
                    //formItems.Add("FromPane", 0);
                    //formItems.Add("ToPane", 0);

                    //FormsB1.createFormItem(oForm, formItems, out errorText);
                    //if (errorText != null)
                    //{
                    //    return;
                    //}

                    //formItems = new Dictionary<string, object>();
                    //itemName = "Descrpt";
                    //formItems.Add("isDataSource", true);
                    //formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    //formItems.Add("DataSource", "UserDataSources");
                    //formItems.Add("Length", 20);
                    //formItems.Add("Size", 20);
                    //formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    //formItems.Add("TableName", "");
                    //formItems.Add("Alias", itemName);
                    //formItems.Add("Bound", true);
                    //formItems.Add("Left", left_s1 + 5 + 100);
                    //formItems.Add("Width", 150);
                    //formItems.Add("Top", top);
                    //formItems.Add("Height", height);
                    //formItems.Add("UID", itemName);
                    //formItems.Add("FromPane", 0);
                    //formItems.Add("ToPane", 0);

                    //FormsB1.createFormItem(oForm, formItems, out errorText);
                    //if (errorText != null)
                    //{
                    //    return;
                    //}
                    //oForm.Items.Item("Descrpt").Specific.Value = "Service";

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
                    formItems.Add("FromPane", 0);
                    formItems.Add("ToPane", 0);

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
                    formItems.Add("FromPane", 0);
                    formItems.Add("ToPane", 0);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "Fill";
                    formItems.Add("Caption", BDOSResources.getTranslate("Fill"));
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left_s + 20 + 1 + 20 + 1);
                    formItems.Add("Width", 150);
                    formItems.Add("Top", top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("FromPane", 0);
                    formItems.Add("ToPane", 0);

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
                    formItems.Add("Left", left_s + 20 + 1 + 20 + 1 + 150 + 1);
                    formItems.Add("Width", 150);
                    formItems.Add("Top", top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("FromPane", 0);
                    formItems.Add("ToPane", 0);

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
                    oDataTable.Columns.Add("DocEntry", SAPbouiCOM.BoFieldsType.ft_Text, 50); //ენთრი
                    oDataTable.Columns.Add("DocNum", SAPbouiCOM.BoFieldsType.ft_Text, 50); //ნომერი
                    oDataTable.Columns.Add("DocDate", SAPbouiCOM.BoFieldsType.ft_Date, 50); //თარიღი
                    oDataTable.Columns.Add("CardCode", SAPbouiCOM.BoFieldsType.ft_Text, 50); //ენთრი
                    oDataTable.Columns.Add("CardName", SAPbouiCOM.BoFieldsType.ft_Text, 50); //ნომერი
                    oDataTable.Columns.Add("LicTradNum", SAPbouiCOM.BoFieldsType.ft_Text, 50); //ენთრი
                    oDataTable.Columns.Add("ReconSum", SAPbouiCOM.BoFieldsType.ft_Sum); //თანხა
                    oDataTable.Columns.Add("TransId", SAPbouiCOM.BoFieldsType.ft_Text, 50); //თანხა
                    oDataTable.Columns.Add("AlRcnSum", SAPbouiCOM.BoFieldsType.ft_Sum); //თანხა
                    oDataTable.Columns.Add("AlRcnVat", SAPbouiCOM.BoFieldsType.ft_Sum); //თანხა
                    oDataTable.Columns.Add("DocEntVT", SAPbouiCOM.BoFieldsType.ft_Text, 50); //ენთრი
                    
                    oDataTable.Columns.Add("DocTotal", SAPbouiCOM.BoFieldsType.ft_Sum); //თანხა
                    oDataTable.Columns.Add("DocVtTotal", SAPbouiCOM.BoFieldsType.ft_Sum); //თანხა
                    oDataTable.Columns.Add("Error", SAPbouiCOM.BoFieldsType.ft_Text, 50); //ენთრი






                    foreach (SAPbouiCOM.DataColumn column in oDataTable.Columns)
                    {
                        string columnName = column.Name;

                        if (columnName == "LineNum")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = "#";
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind("InvoiceMTR", columnName);
                            oColumn.AffectsFormMode = false;
                        }
                        else if (columnName == "CheckBox")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
                            oColumn.TitleObject.Caption = "";
                            oColumn.Editable = true;
                            oColumn.ValOff = "N";
                            oColumn.ValOn = "Y";
                            oColumn.DataBind.Bind("InvoiceMTR", columnName);
                            oColumn.AffectsFormMode = false;
                        }


                        else if (columnName == "DocEntry")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate(columnName);
                            oColumn.Editable = false;
                            oLink = oColumn.ExtendedObject;
                            oLink.LinkedObjectType = "203";
                            oColumn.DataBind.Bind("InvoiceMTR", columnName);
                            oColumn.AffectsFormMode = false;
                        }

                        else if (columnName == "DocEntVT")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate(columnName);
                            oColumn.Editable = false;
                            oLink = oColumn.ExtendedObject;
                            oLink.LinkedObjectType = "UDO_F_BDO_ARDPV_D";
                            oColumn.DataBind.Bind("InvoiceMTR", columnName);
                            oColumn.AffectsFormMode = false;
                        }
                        else if (columnName == "CardCode")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("BPCardCode");
                            oColumn.Editable = false;
                            oLink = oColumn.ExtendedObject;
                            oLink.LinkedObjectType = "2";
                            oColumn.DataBind.Bind("InvoiceMTR", columnName);
                            oColumn.AffectsFormMode = false;
                        }
                        else if (columnName == "TransId")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("TransId");
                            oColumn.Editable = false;
                            oLink = oColumn.ExtendedObject;
                            oLink.LinkedObjectType = "30";
                            oColumn.DataBind.Bind("InvoiceMTR", columnName);
                            oColumn.AffectsFormMode = false;
                        }
                        else if (columnName == "CardName")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("BPName");
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind("InvoiceMTR", columnName);
                            oColumn.AffectsFormMode = false;

                        }
                        else if (columnName == "LicTradNum")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("BPTin");
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind("InvoiceMTR", columnName);
                            oColumn.AffectsFormMode = false;

                        }
                        else if (columnName == "DocTotal")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Total");
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind("InvoiceMTR", columnName);
                            oColumn.AffectsFormMode = false;

                        }
                        else if (columnName == "DocVtTotal")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("VatAmount");
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind("InvoiceMTR", columnName);
                            oColumn.AffectsFormMode = false;

                        }

                        else if (columnName == "InDetail")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_PICTURE);
                            oColumn.Width = 15;
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind("InvoiceMTR", columnName);
                        }
                        else
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate(columnName);
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind("InvoiceMTR", columnName);
                            oColumn.AffectsFormMode = false;

                        }
                    }
                    oMatrix.Clear();
                    oMatrix.LoadFromDataSource();
                    oMatrix.AutoResizeColumns();
                }


                oForm.Visible = true;
                oForm.Select();
            }
        }

        public static void uiApp_ItemEvent(  string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                {
                    SAPbouiCOM.IChooseFromListEvent oCFLEvento = null;
                    oCFLEvento = ((SAPbouiCOM.IChooseFromListEvent)(pVal));

                    ChooseFromList( oForm, pVal.BeforeAction, oCFLEvento, out errorText);

                }

                if ((pVal.ItemUID == "InCheck" || pVal.ItemUID == "InUncheck") && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.BeforeAction == false)
                {
                    CheckUncheck(oForm, pVal.ItemUID, "", out errorText);
                }

                if (pVal.ItemUID == "DocPstDt" && pVal.ItemChanged && pVal.BeforeAction == false)
                {
                    SAPbouiCOM.EditText oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("DocPstDt").Specific;
                    DateTime EndDateOp = DateTime.ParseExact(oEditText.Value, "yyyyMMdd", null);
                    EndDateOp = new DateTime(EndDateOp.Year, EndDateOp.Month, 1).AddMonths(1).AddDays(-1);
                    oEditText.Value = EndDateOp.ToString("yyyyMMdd");

                    FillMTRInvoice( oForm);
                }
                if (pVal.ItemUID == "Fill" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.BeforeAction == false)
                {
                    FillMTRInvoice( oForm);
                }



                if (pVal.ItemUID == "InvoiceMTR")
                {
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED)
                    {
                        matrixColumnSetLinkedObjectTypeInvoicesMTR(oForm, pVal, out errorText);
                    }

                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.BeforeAction == false)
                    {
                        int row = pVal.Row;

                        oForm.Freeze(true);
                        SetInvDocsMatrixRowBackColor( oForm, row, out errorText);
                        oForm.Freeze(false);
                    }

                }
                if (pVal.ItemUID == "CreatDocmt" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.BeforeAction == false)
                {
                    CreatePaymentDocuments( oForm);

                }
            }
        }

        public static void SetInvDocsMatrixRowBackColor( SAPbouiCOM.Form oForm,  int row, out string errorText)
        {
            errorText = null;

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
                    oForm.Freeze(true);
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

        public static void SetInvDocsMatrixRowCellColor(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            try
            {
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("InvoiceMTR").Specific;

                if (oMatrix.RowCount > 0)
                {
                    oForm.Freeze(false);

                    for (int i = 1; i <= oMatrix.RowCount; i++)
                    {
                        if (oMatrix.GetCellSpecific("ReconSum", i).Value != oMatrix.GetCellSpecific("AlRcnSum", i).Value)
                        {
                            oMatrix.CommonSetting.SetRowFontColor(i, FormsB1.getLongIntRGB(255, 0, 0));
                        }
                        else
                        {
                            oMatrix.CommonSetting.SetRowFontColor(i, FormsB1.getLongIntRGB(0, 0, 0));
                        }
                    }

                    oForm.Freeze(true);
                }

            }
            catch 
            {
            }
            finally
            {
                GC.Collect();
            }
        }

        private static void CheckUncheck(SAPbouiCOM.Form oForm, string CheckOperation, string type, out string errorText)
        {
            errorText = null;

            oForm.Freeze(true);

            SAPbouiCOM.CheckBox oCheckBox;
            SAPbouiCOM.Matrix oMatrix;

            oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("InvoiceMTR").Specific));

            int rowCount = oMatrix.RowCount;
            for (int j = 1; j <= rowCount; j++)
            {
                oCheckBox = oMatrix.Columns.Item("CheckBox").Cells.Item(j).Specific;
                oCheckBox.Checked = (CheckOperation == "InCheck");
            }
            oForm.Freeze(false);
        }

        public static void matrixColumnSetLinkedObjectTypeInvoicesMTR(SAPbouiCOM.Form oForm, SAPbouiCOM.ItemEvent pVal, out string errorText)
        {
            errorText = null;

            try
            {
                if (pVal.ColUID == "DocEntry")
                {
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED & pVal.BeforeAction == true)
                    {
                        SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("InvoiceMTR").Specific));

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
                else
                {

                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
            finally
            {
                GC.Collect();
            }
        }

        private static void CreatePaymentDocuments(  SAPbouiCOM.Form oForm)
        {
            int answer = Program.uiApp.MessageBox(BDOSResources.getTranslate("CreatePaymentDocuments") + "?", 1, BDOSResources.getTranslate("Yes"), BDOSResources.getTranslate("No"), "");

            if (answer == 2)
            {
                return;
            }

            DataTable AccountTable = CommonFunctions.GetOACTTable();

            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("InvoiceMTR").Specific;

            SAPbouiCOM.EditText oEditTextDocDate = (SAPbouiCOM.EditText)oForm.Items.Item("DocPstDt").Specific;
            String DocDateS = oEditTextDocDate.Value;
            DateTime DocDate = Convert.ToDateTime(DateTime.ParseExact(DocDateS, "yyyyMMdd", CultureInfo.InvariantCulture));
            DateTime DocDateEE = (new DateTime(DocDate.Year, DocDate.Month, DateTime.DaysInMonth(DocDate.Year, DocDate.Month)));

           

            DataTable reLines = new DataTable();
            reLines.Columns.Add("month");
            reLines.Columns.Add("vatAccrl");
            reLines.Columns.Add("reconSum");
            reLines.Columns.Add("vatSum");
            
            DataTable jeLines = JournalEntry.JournalEntryTable();

            for (int row = 1; row <= oMatrix.RowCount; row++)
            {

                bool checkedLine = oMatrix.GetCellSpecific("CheckBox", row).Checked;


                if (checkedLine)
                {
                    jeLines.Rows.Clear();
                    reLines.Rows.Clear();

                    string DocEntVT = oMatrix.GetCellSpecific("DocEntVT", row).Value;
                    
                    Dictionary<string, string> listAccounts = GetVatAcrualJornalEntry( DocEntVT);
                    string VatGrp = listAccounts["VatGroup"].ToString();

                    decimal VatRate = CommonFunctions.GetVatGroupRate( VatGrp, "");

                    string TransId = oMatrix.GetCellSpecific("TransId", row).Value;
                    if(TransId != "")
                    {
                        Program.uiApp.SetStatusBarMessage("დღგ უკვე გატარებულია", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                        continue;
                    }

                    
                    string InvDocEntry = oMatrix.GetCellSpecific("DocEntry", row).Value;
                    decimal DocVtTotal = Convert.ToDecimal(oMatrix.GetCellSpecific("DocVtTotal", row).Value, CultureInfo.InvariantCulture);
                    decimal ReconSum = Convert.ToDecimal(oMatrix.GetCellSpecific("ReconSum", row).Value, CultureInfo.InvariantCulture);
                    decimal ReconSumVAT = CommonFunctions.roundAmountByGeneralSettings( ReconSum * VatRate / (100 + VatRate),"Sum");

                    DocVtTotal = Math.Min(DocVtTotal, ReconSumVAT);

                    JournalEntry.AddJournalEntryRow(AccountTable, jeLines, "OnlyCredit", listAccounts["CreditAccount"], listAccounts["DebitAccount"], DocVtTotal, 0, "", "", "", "", "", "", "", "", "");

                    JournalEntry.AddJournalEntryRow(AccountTable, jeLines, "OnlyDebit", listAccounts["CreditAccount"], listAccounts["DebitAccount"], DocVtTotal, 0, "", "", "", "", "", "", "", VatGrp, "");

                    DataRow reLinesRow = reLines.Rows.Add();
                    reLinesRow["month"] = DocDateEE.ToString("yyyyMMdd");
                    reLinesRow["vatAccrl"] = DocEntVT;
                    reLinesRow["reconSum"] = ReconSum.ToString(CultureInfo.InvariantCulture);
                    reLinesRow["vatSum"] = ReconSumVAT.ToString(CultureInfo.InvariantCulture);

                    string errorText = null;

                    try
                    {
                        JournalEntry.JrnEntry( DocEntVT, "Reconcilation", "Reconcilation ", DocDate, jeLines, out errorText);
                        AddRecord( reLines,  out errorText);

                        if (errorText != null)
                        {
                            return;
                        }

                    }
                    catch (Exception ex)
                    {
                        errorText = ex.Message;
                    }

                }
            }



            FillMTRInvoice( oForm);

        }

        private static Dictionary<string, string> GetVatAcrualJornalEntry( string DocEntry)
        {
            Dictionary<string, string> listValidValuesDict = new Dictionary<string, string>();

            string query = @"select  top 2 * 
            from ""JDT1""
            inner join ""OJDT"" on ""JDT1"".""TransId"" = ""OJDT"".""TransId"" and ""OJDT"".""Ref1"" = '"+ DocEntry+ @"'
                and ""OJDT"".""Ref2""  = 'UDO_F_BDO_ARDPV_D' 
            where ""OJDT"".""TransId"" not in 
            (select ""StornoToTr"" from ""OJDT"" where ""StornoToTr"" is not null)
            AND ""OJDT"".""StornoToTr"" Is NULL";

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRecordSet.DoQuery(query);

            while (!oRecordSet.EoF)
            {
                decimal debit = Convert.ToDecimal(oRecordSet.Fields.Item("Debit").Value, CultureInfo.InvariantCulture);
                string VatGroup = oRecordSet.Fields.Item("VatGroup").Value.ToString();

                if (VatGroup!="")
                {
                    listValidValuesDict.Add("VatGroup", VatGroup);
                }


                if (debit > 0)
                {
                    listValidValuesDict.Add("DebitAccount", oRecordSet.Fields.Item("Account").Value.ToString());
                }
                else
                {
                    listValidValuesDict.Add("CreditAccount", oRecordSet.Fields.Item("Account").Value.ToString());
                }
                oRecordSet.MoveNext();
            }

            return listValidValuesDict;
        }

        private static void ChooseFromList( SAPbouiCOM.Form oForm, bool BeforeAction, SAPbouiCOM.IChooseFromListEvent oCFLEvento, out string errorText)
        {
            errorText = null;

            string sCFL_ID = oCFLEvento.ChooseFromListUID;
            SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(sCFL_ID);

            SAPbouiCOM.DataTable oDataTable = null;
            oDataTable = oCFLEvento.SelectedObjects;

            if (BeforeAction == false)
            {
                if (oDataTable != null)
                {
                    try
                    {
                        if (sCFL_ID == "BusinessPartner_CFL")
                        {
                            string CardCode = Convert.ToString(oDataTable.GetValue("CardCode", 0));

                            SAPbouiCOM.EditText oBPCode = oForm.Items.Item("BPCode").Specific;
                            oBPCode.Value = CardCode;
                            FillMTRInvoice( oForm);
                        }



                    }
                    catch
                    {
                        FillMTRInvoice( oForm);
                    }

                }
            }
        }

        public static void createUDO( out string errorText)
        {
            errorText = null;

            string tableName = "BDOSRECWIZ";
            string description = "Reconcilation wizard history";

            SAPbobsCOM.UserObjectsMD oUserObjectsMD = null;
            oUserObjectsMD = (SAPbobsCOM.UserObjectsMD)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectsMD);

            int result = UDO.addUserTable( tableName, description, SAPbobsCOM.BoUTBTableType.bott_NoObjectAutoIncrement, out errorText);

            if (result != 0)
            {
                return;
            }

            Dictionary<string, object> fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap = new Dictionary<string, object>(); // docType
            fieldskeysMap.Add("Name", "month");
            fieldskeysMap.Add("TableName", "BDOSRECWIZ");
            fieldskeysMap.Add("Description", "Month");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Date);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            
            fieldskeysMap = new Dictionary<string, object>(); // docType
            fieldskeysMap.Add("Name", "vatAccrl");
            fieldskeysMap.Add("TableName", "BDOSRECWIZ");
            fieldskeysMap.Add("Description", "Vat accrual");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // docType
            fieldskeysMap.Add("Name", "reconSum");
            fieldskeysMap.Add("TableName", "BDOSRECWIZ");
            fieldskeysMap.Add("Description", "Reconcilation sum");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // docType
            fieldskeysMap.Add("Name", "vatSum");
            fieldskeysMap.Add("TableName", "BDOSRECWIZ");
            fieldskeysMap.Add("Description", "Vat sum");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

        }

        public static void AddRecord( DataTable reLines, out string errorText)
        {
            errorText = null;
            int returnCode;

            SAPbobsCOM.UserTable oUserTable = Program.oCompany.UserTables.Item("BDOSRECWIZ");
            DataRow reLine;

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            


            try
            {
                for (int i = 0; i < reLines.Rows.Count; i++)
                {
                    reLine = reLines.Rows[i];

                    string queryOPDF = @"delete from ""@BDOSRECWIZ""
                                    where ""U_month"" = '" + reLine["month"]+ @"' and ""U_vatAccrl"" = " + reLine["vatAccrl"].ToString();
                    oRecordSet.DoQuery(queryOPDF);
                    
                    oUserTable.UserFields.Fields.Item("U_month").Value = DateTime.ParseExact(reLine["month"].ToString(), "yyyyMMdd", CultureInfo.InvariantCulture);
                    oUserTable.UserFields.Fields.Item("U_vatAccrl").Value = reLine["vatAccrl"].ToString();
                    
                    oUserTable.UserFields.Fields.Item("U_reconSum").Value = Convert.ToDouble(reLine["reconSum"],CultureInfo.InvariantCulture);
                    oUserTable.UserFields.Fields.Item("U_vatSum").Value = Convert.ToDouble(reLine["vatSum"], CultureInfo.InvariantCulture);


                    returnCode = oUserTable.Add();

                    if (returnCode != 0)
                    {
                        int errCode;
                        string errMsg;

                        Program.oCompany.GetLastError(out errCode, out errMsg);
                        errorText = "Error description : " + errMsg + "! Code : " + errCode;
                    }
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;

            }
            finally
            {
                Marshal.ReleaseComObject(oUserTable);
                oUserTable = null;
            }
        }


        public static void FillMTRInvoice( SAPbouiCOM.Form oForm)
        {

            SAPbouiCOM.EditText oEditTextDocDate = (SAPbouiCOM.EditText)oForm.Items.Item("DocPstDt").Specific;
            String DocDateS = oEditTextDocDate.Value;
            DateTime date = Convert.ToDateTime(DateTime.ParseExact(DocDateS, "yyyyMMdd", CultureInfo.InvariantCulture));
            DateTime prevDate = date.AddMonths(-1);

            string dateES = (new DateTime(date.Year, date.Month, 1)).ToString("yyyyMMdd");
            string dateEE = (new DateTime(date.Year, date.Month, DateTime.DaysInMonth(date.Year, date.Month))).ToString("yyyyMMdd");

            string prevdateES = (new DateTime(prevDate.Year, prevDate.Month, 1)).ToString("yyyyMMdd");
            string prevdateEE = (new DateTime(prevDate.Year, prevDate.Month, DateTime.DaysInMonth(prevDate.Year, prevDate.Month))).ToString("yyyyMMdd");

            string cardCodeE = oForm.Items.Item("BPCode").Specific.Value;
            cardCodeE = cardCodeE == "" ? "1=1" : @" ""ODPI"".""CardCode"" = '" + cardCodeE + "'";

            SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("InvoiceMTR");

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string errorText = "";



            string query = @"select 
                    ""ODPI"".""DocEntry"",
                    ""ODPI"".""DocNum"",
                    ""ODPI"".""DocDate"",
                    ""ODPI"".""CardCode"",
                    ""ODPI"".""CardName"",
                    ""ITR1"".""ReconSum"",
                    IFNULL(""OJDT1"".""Debit"",0)as ""AlRcnVat"",
                    IFNULL(""OJDT1"".""TransId"",0)as ""TransId"",
					IFNULL(""OCRD"".""LicTradNum"",'') as ""LicTradNum"",
                    IFNULL(""OJDT1"".""BaseSum"",0) as ""U_reconSum"",
                    ""@BDOSARDV"".""U_GrsAmnt""- IFNULL(""OJDT"".""BaseSum"",0) as ""DocTotal"",                    
                    ""@BDOSARDV"".""U_VatAmount""- IFNULL(""OJDT"".""Debit"",0) as ""DocVAtTotal"",                    
                    IFNULL(""@BDOSARDV"".""DocEntry"",0) as ""DocEntVT"",
                    IFNULL(""@BDOSARDV"".""DocNum"",0) as ""DocNumVT""

                    from (select 
					""RCT2"".""DocEntry"",
                    SUM(""ITR1"".""ReconSum"") as ""ReconSum""
					from ""RCT2""
                    inner join ""ORCT"" on ""ORCT"".""DocEntry"" = ""RCT2"".""DocNum"" and ""ORCT"".""Canceled"" = 'N'
					inner join ""ITR1"" on  ""RCT2"".""InvoiceId"" = ""ITR1"".""LineSeq"" and ""RCT2"".""DocNum"" = ""ITR1"".""SrcObjAbs"" and ""ITR1"".""SrcObjTyp"" = 24 and ""RCT2"".""InvType"" = 203
					inner join ""OITR"" on  ""ITR1"".""ReconNum"" = ""OITR"".""ReconNum"" and ""OITR"".""ReconDate"">='" + dateES + @"' and ""OITR"".""ReconDate""<='" + dateEE + @"' and ""OITR"".""Canceled""<>'C'
					group by ""RCT2"".""DocEntry"",""OITR"".""Canceled"") as ""ITR1""
                    inner join  ""ODPI"" on  ""ITR1"".""DocEntry"" = ""ODPI"".""DocEntry"" and ""DocStatus""='C'  and " + cardCodeE + @"                                      
                    inner join ""@BDOSARDV"" on ""ODPI"".""DocEntry""= ""@BDOSARDV"".""U_baseDoc"" and ""@BDOSARDV"".""U_baseDocT""=203 and ""@BDOSARDV"".""Canceled""='N' and ""@BDOSARDV"".""U_DocDate""<='" + dateES + @"'
                    inner join ""OCRD"" on ""ODPI"".""CardCode""= ""OCRD"".""CardCode""
                    left join (select
	                    
	                     ""OJDT"".""Ref1"",
                        sum(""OJDT"".""BaseSum"") as ""BaseSum"", 	                     
                        sum(""OJDT"".""Debit"") as ""Debit"" 
                    from( select
	                     ""OJDT"".""TransId"" as ""TransId"",
	                     ""OJDT"".""Ref1"",
	                     SUM(""JDT1"".""BaseSum"" + ""JDT1"".""Debit"") as ""BaseSum"",
	                     SUM(""JDT1"".""Debit"") as ""Debit"" 
	                    from ""JDT1"" 
	                    inner join ""OJDT"" on ""JDT1"".""TransId"" = ""OJDT"".""TransId"" and ""OJDT"".""TaxDate""<='" + dateEE + @"' 
	                    and ""OJDT"".""Ref2"" = 'Reconcilation' 
	                    and ""JDT1"".""VatGroup""<>'' 
	                    and ""OJDT"".""StornoToTr"" is null 
	                    group by ""OJDT"".""Ref1"",
	                     ""OJDT"".""TransId"" 
	                    union all select
	                     ""OJDT"".""StornoToTr"" as ""TransId"",
	                     ""OJDT"".""Ref1"",
                             SUM(""JDT1"".""BaseSum"" + Case when ""JDT1"".""Credit"">0 then -""JDT1"".""Credit"" else  ""JDT1"".""Debit"" End) as ""BaseSum"",
	                     SUM(Case when ""JDT1"".""Credit"">0 then -""JDT1"".""Credit"" else  ""JDT1"".""Debit"" End) as ""Credit"" 
	                    from ""JDT1"" 
	                    inner join ""OJDT"" on ""JDT1"".""TransId"" = ""OJDT"".""TransId"" and ""OJDT"".""TaxDate""<='" + dateEE + @"' 
	                    and ""OJDT"".""Ref2"" = 'Reconcilation' 
	                    and ""JDT1"".""VatGroup""<>'' 
	                    and not ""OJDT"".""StornoToTr"" is null 
	                    group by ""OJDT"".""Ref1"",
	                     ""OJDT"".""StornoToTr"" ) as ""OJDT"" 
                    group by 
	                     ""OJDT"".""Ref1"" Having sum(""OJDT"".""Debit"")>0 ) as ""OJDT"" on ""@BDOSARDV"".""DocEntry""=""OJDT"".""Ref1""

                    left join (
                               select
	                     ""OJDT"".""TransId"" as ""TransId"",
	                     ""OJDT"".""Ref1"",
                        sum(""OJDT"".""BaseSum"") as ""BaseSum"", 
	                     sum(""OJDT"".""Debit"") as ""Debit"" 
                    from( select
	                     ""OJDT"".""TransId"" as ""TransId"",
	                     ""OJDT"".""Ref1"",
                        SUM(""JDT1"".""BaseSum"" + Case when ""JDT1"".""Credit"">0 then -""JDT1"".""Credit"" else  ""JDT1"".""Debit"" End) as ""BaseSum"",
	                     SUM(""JDT1"".""Debit"") as ""Debit"" 
	                    from ""JDT1"" 
	                    inner join ""OJDT"" on ""JDT1"".""TransId"" = ""OJDT"".""TransId"" and ""OJDT"".""TaxDate"">='" + dateES + @"' and ""OJDT"".""TaxDate""<='" + dateEE + @"' 
	                    and ""OJDT"".""Ref2"" = 'Reconcilation' 
	                    and ""JDT1"".""VatGroup""<>'' 
	                    and ""OJDT"".""StornoToTr"" is null 
	                    group by ""OJDT"".""Ref1"",
	                     ""OJDT"".""TransId"" 
	                    union all select
	                     ""OJDT"".""StornoToTr"" as ""TransId"",
	                     ""OJDT"".""Ref1"",
                        SUM(""JDT1"".""BaseSum"" + Case when ""JDT1"".""Credit"">0 then -""JDT1"".""Credit"" else  ""JDT1"".""Debit"" End) as ""BaseSum"",
	                    SUM(Case when ""JDT1"".""Credit"">0 then -""JDT1"".""Credit"" else  ""JDT1"".""Debit"" End ) as ""Credit"" 
	                    from ""JDT1"" 
	                    inner join ""OJDT"" on ""JDT1"".""TransId"" = ""OJDT"".""TransId"" and ""OJDT"".""TaxDate"">='" + dateES + @"' and ""OJDT"".""TaxDate""<='" + dateEE + @"' 
	                    and ""OJDT"".""Ref2"" = 'Reconcilation' 
	                    and ""JDT1"".""VatGroup""<>'' 
	                    and not ""OJDT"".""StornoToTr"" is null 
	                    group by ""OJDT"".""Ref1"",
	                     ""OJDT"".""StornoToTr"" ) as ""OJDT"" 
                    group by ""OJDT"".""TransId"",
	                     ""OJDT"".""Ref1"" Having sum(""OJDT"".""Debit"")>0 ) as ""OJDT1"" on ""@BDOSARDV"".""DocEntry""=""OJDT1"".""Ref1""

                

                    order by  ""ODPI"".""DocDate"" asc";
            if (Program.oCompany.DbServerType != SAPbobsCOM.BoDataServerTypes.dst_HANADB)
            {
                query = query.Replace("IFNULL", "ISNULL");
            }

            oRecordSet.DoQuery(query);

            oDataTable.Rows.Clear();

            try
            {
                int rowIndex = 0;
                string DocEntry;
                string DocNum;
                string DocDate;
                
                string DocEntVT;
                string CardCode;
                string CardName;
                string LicTradNum;
                string TransId;
                decimal DocTotal = 0;
                decimal DocVatTotal = 0;
                decimal ReconSum = 0;
                decimal AlRcnSum = 0;
                decimal AlRcnVat = 0;


                ItemsDT = new DataTable();

                ItemsDT.Columns.Add("DocEntry");
                ItemsDT.Columns.Add("ItemCode");
                ItemsDT.Columns.Add("Dscptn");
                ItemsDT.Columns.Add("GrsAmnt");
                ItemsDT.Columns.Add("VatGrp");
                ItemsDT.Columns.Add("VatAmnt");


                while (!oRecordSet.EoF)
                {
                    DocEntry = oRecordSet.Fields.Item("DocEntry").Value == 0 ? "" : oRecordSet.Fields.Item("DocEntry").Value.ToString();
                    DocNum = oRecordSet.Fields.Item("DocNum").Value == 0 ? "" : oRecordSet.Fields.Item("DocNum").Value.ToString();
                    DocDate = oRecordSet.Fields.Item("DocDate").Value.ToString("yyyyMMdd");
                    DocEntVT = oRecordSet.Fields.Item("DocEntVT").Value == 0 ? "" : oRecordSet.Fields.Item("DocEntVT").Value.ToString();
                   
                    decimal allowableDeviation = Convert.ToDecimal(CommonFunctions.getOADM("U_BDOSAllDev").ToString(), CultureInfo.InvariantCulture);
                    TransId = oRecordSet.Fields.Item("TransId").Value == 0 ? "" : oRecordSet.Fields.Item("TransId").Value.ToString();
                    
                    DocTotal = Convert.ToDecimal(oRecordSet.Fields.Item("DocTotal").Value, CultureInfo.InvariantCulture);
                    DocVatTotal = Convert.ToDecimal(oRecordSet.Fields.Item("DocVatTotal").Value, CultureInfo.InvariantCulture);
                    AlRcnVat = Convert.ToDecimal(oRecordSet.Fields.Item("AlRcnVat").Value, CultureInfo.InvariantCulture);
                    AlRcnSum = Convert.ToDecimal(oRecordSet.Fields.Item("U_reconSum").Value, CultureInfo.InvariantCulture);
                    ReconSum = Convert.ToDecimal(oRecordSet.Fields.Item("ReconSum").Value, CultureInfo.InvariantCulture);
                    CardCode = oRecordSet.Fields.Item("CardCode").Value;
                    CardName = oRecordSet.Fields.Item("CardName").Value;
                    LicTradNum = oRecordSet.Fields.Item("LicTradNum").Value;

                    if (Math.Abs(DocTotal) <= allowableDeviation)
                    {
                        DocTotal = 0;
                    }
                    if (Math.Abs(AlRcnSum - ReconSum) <= allowableDeviation)
                    {
                        AlRcnSum = ReconSum;
                    }
                    
                    if (TransId=="" && ReconSum==0)
                    {
                        oRecordSet.MoveNext();
                        continue;
                    }

                    oDataTable.Rows.Add();
                    oDataTable.SetValue("LineNum", rowIndex, rowIndex + 1);
                    oDataTable.SetValue("CheckBox", rowIndex, "N");
                    oDataTable.SetValue("DocEntry", rowIndex, DocEntry);
                    oDataTable.SetValue("DocNum", rowIndex, DocNum);
                    oDataTable.SetValue("DocEntVT", rowIndex, DocEntVT);
                    
                    oDataTable.SetValue("DocDate", rowIndex, DocDate);
                    oDataTable.SetValue("CardCode", rowIndex, CardCode);
                    oDataTable.SetValue("CardName", rowIndex, CardName);
                    oDataTable.SetValue("LicTradNum", rowIndex, LicTradNum);
                    oDataTable.SetValue("TransId", rowIndex, TransId);
                    
                    oDataTable.SetValue("DocTotal", rowIndex, Convert.ToDouble(DocTotal,CultureInfo.InvariantCulture));
                    oDataTable.SetValue("DocVtTotal", rowIndex, Convert.ToDouble(DocVatTotal, CultureInfo.InvariantCulture));
                    oDataTable.SetValue("ReconSum", rowIndex, Convert.ToDouble(ReconSum, CultureInfo.InvariantCulture));
                    oDataTable.SetValue("AlRcnVat", rowIndex, Convert.ToDouble(AlRcnVat, CultureInfo.InvariantCulture));
                    oDataTable.SetValue("AlRcnSum", rowIndex, Convert.ToDouble(AlRcnSum, CultureInfo.InvariantCulture));
                    



                    oRecordSet.MoveNext();
                    rowIndex++;
                }

                SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("InvoiceMTR").Specific));
                oForm.Freeze(true);
                oMatrix.Clear();
                oMatrix.LoadFromDataSource();
                oMatrix.AutoResizeColumns();
                oForm.Update();
                oForm.Freeze(false);

                SetInvDocsMatrixRowCellColor(oForm, out errorText);

            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
            finally
            {
                oForm.Freeze(false);
                oRecordSet = null;
            }
        }

        public static void addMenus()
        {
            SAPbouiCOM.MenuItem menuItem;
            SAPbouiCOM.MenuItem fatherMenuItem;
            SAPbouiCOM.MenuCreationParams oCreationPackage;

            try
            {
                fatherMenuItem = Program.uiApp.Menus.Item("2048");

                // Add a pop-up menu item
                oCreationPackage = (SAPbouiCOM.MenuCreationParams)Program.uiApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                oCreationPackage.Checked = false;
                oCreationPackage.Enabled = true;
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "BDOSReconWizz";
                oCreationPackage.String = BDOSResources.getTranslate("VATReconcilationWizard");
                oCreationPackage.Position = fatherMenuItem.SubMenus.Count - 1;

                menuItem = fatherMenuItem.SubMenus.AddEx(oCreationPackage);
            }
            catch 
            {
               
            }
        }
    }
}

     
      