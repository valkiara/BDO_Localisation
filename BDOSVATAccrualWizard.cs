using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;
using System.Data;

namespace BDO_Localisation_AddOn
{
    class BDOSVATAccrualWizard
    {
        static DataTable ItemsDT = null;
        static string DocEntryAdd = "";

        public static void createForm(out string errorText)
        {
            errorText = null;
            Dictionary<string, object> formItems;
            string itemName;
            SAPbouiCOM.Columns oColumns;
            SAPbouiCOM.Column oColumn;

            SAPbouiCOM.DataTable oDataTable;

            bool multiSelection;


            int left_s = 5;
            int left_s1 = 330;

            int top = 10;
            int height = 15;

            //ფორმის აუცილებელი თვისებები
            Dictionary<string, object> formProperties = new Dictionary<string, object>();
            formProperties.Add("UniqueID", "BDOSVAWizzForm");
            formProperties.Add("BorderStyle", SAPbouiCOM.BoFormBorderStyle.fbs_Sizable);
            formProperties.Add("Title", BDOSResources.getTranslate("ARVATAccrualWizard"));
            formProperties.Add("Left", 558);
            formProperties.Add("ClientWidth", 600);
            formProperties.Add("Top", 335);
            formProperties.Add("ClientHeight", 600);

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

                    formItems = new Dictionary<string, object>();
                    itemName = "VatGrpS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s1);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("VatGroup"));
                    formItems.Add("FromPane", 0);
                    formItems.Add("ToPane", 0);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "VatGrp"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("TableName", "");
                    formItems.Add("Length", 20);
                    formItems.Add("Size", 20);
                    formItems.Add("Alias", "VatGrp");
                    formItems.Add("Bound", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                    formItems.Add("Left", left_s1 + 5 + 100);
                    formItems.Add("Width", 150);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("FromPane", 0);
                    formItems.Add("ToPane", 0);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    string query = "select * " +
                    "FROM  \"OVTG\" " +
                    "WHERE \"Category\"='O'";
                    oRecordSet.DoQuery(query);

                    while (!oRecordSet.EoF)
                    {
                        oForm.Items.Item("VatGrp").Specific.ValidValues.Add(oRecordSet.Fields.Item("Code").Value, oRecordSet.Fields.Item("Code").Value);
                        oRecordSet.MoveNext();
                    }

                    oForm.Items.Item("VatGrp").Specific.Select(CommonFunctions.getOADM("DfSVatItem"));



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

                    formItems = new Dictionary<string, object>();
                    itemName = "DescrptS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s1);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("Description"));
                    formItems.Add("FromPane", 0);
                    formItems.Add("ToPane", 0);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "Descrpt";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("Length", 20);
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Left", left_s1 + 5 + 100);
                    formItems.Add("Width", 150);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("FromPane", 0);
                    formItems.Add("ToPane", 0);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }
                    oForm.Items.Item("Descrpt").Specific.Value = "Service";

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
                    oDataTable.Columns.Add("cancelled", SAPbouiCOM.BoFieldsType.ft_Text, 1); //ნომერი                    
                    oDataTable.Columns.Add("DocDate", SAPbouiCOM.BoFieldsType.ft_Date, 50); //თარიღი
                    oDataTable.Columns.Add("CardCode", SAPbouiCOM.BoFieldsType.ft_Text, 50); //ენთრი
                    oDataTable.Columns.Add("CardName", SAPbouiCOM.BoFieldsType.ft_Text, 50); //ნომერი
                    oDataTable.Columns.Add("LicTradNum", SAPbouiCOM.BoFieldsType.ft_Text, 50); //ენთრი
                    oDataTable.Columns.Add("Total", SAPbouiCOM.BoFieldsType.ft_Sum); //თანხა
                    oDataTable.Columns.Add("DocEntVT", SAPbouiCOM.BoFieldsType.ft_Text, 50); //ენთრი

                    oDataTable.Columns.Add("DocTotal", SAPbouiCOM.BoFieldsType.ft_Sum); //თანხა
                    oDataTable.Columns.Add("DocVtTotal", SAPbouiCOM.BoFieldsType.ft_Sum); //თანხა
                    oDataTable.Columns.Add("InDetail", SAPbouiCOM.BoFieldsType.ft_Text, 20);
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
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Selected");
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
                        else if (columnName == "InDetail")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_PICTURE);
                            oColumn.Width = 15;
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind("InvoiceMTR", columnName);
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

        private static int createPaymentDocument(SAPbouiCOM.Form oForm, DataRow headerLine)
        {
            string errorText = null;

            SAPbobsCOM.SBObob vObj;
            vObj = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);

            SAPbobsCOM.CompanyService oCompanyService = null;
            SAPbobsCOM.GeneralService oGeneralService = null;
            SAPbobsCOM.GeneralData oGeneralData = null;
            oCompanyService = Program.oCompany.GetCompanyService();
            oGeneralService = oCompanyService.GetGeneralService("UDO_F_BDO_ARDPV_D");
            oGeneralData = ((SAPbobsCOM.GeneralData)(oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)));

            DateTime DocDate = Convert.ToDateTime(DateTime.ParseExact(headerLine["DocDate"].ToString(), "yyyyMMdd", CultureInfo.InvariantCulture));

            oGeneralData.SetProperty("U_DocDate", DocDate);
            oGeneralData.SetProperty("U_cardCode", headerLine["CardCode"].ToString());
            oGeneralData.SetProperty("U_cardCodeN", headerLine["CardName"].ToString());
            oGeneralData.SetProperty("U_baseDoc", headerLine["DocEntry"].ToString());
            oGeneralData.SetProperty("U_baseDocT", "203");
            oGeneralData.SetProperty("U_GrsAmnt", Convert.ToDouble(headerLine["Total"], CultureInfo.InvariantCulture));
            oGeneralData.SetProperty("U_VatAmount", Convert.ToDouble(headerLine["Total"], CultureInfo.InvariantCulture) * 18 / 118);
            SAPbobsCOM.GeneralDataCollection oChildren = null;

            oChildren = oGeneralData.Child("BDOSRDV1");
            double U_VatAmount = 0;
            //////////
            for (int row = 0; row < ItemsDT.Rows.Count; row++)
            {
                if (headerLine["DocEntry"].ToString() == ItemsDT.Rows[row]["DocEntry"].ToString())
                {
                    SAPbobsCOM.GeneralData oChild = oChildren.Add();
                    oChild.SetProperty("U_ItemCode", ItemsDT.Rows[row]["ItemCode"]);
                    oChild.SetProperty("U_Dscptn", ItemsDT.Rows[row]["Dscptn"]);
                    oChild.SetProperty("U_GrsAmnt", Convert.ToDouble(ItemsDT.Rows[row]["GrsAmnt"], CultureInfo.InvariantCulture));
                    oChild.SetProperty("U_VatGrp", ItemsDT.Rows[row]["VatGrp"]);
                    oChild.SetProperty("U_VatAmount", Convert.ToDouble(ItemsDT.Rows[row]["VatAmnt"], CultureInfo.InvariantCulture));
                    U_VatAmount = U_VatAmount + Convert.ToDouble(ItemsDT.Rows[row]["VatAmnt"], CultureInfo.InvariantCulture);
                }
            }
            oGeneralData.SetProperty("U_VatAmount", U_VatAmount);
            //////////


            int docEntry = 0;
            try
            {
                CommonFunctions.StartTransaction();

                var response = oGeneralService.Add(oGeneralData);
                docEntry = response.GetProperty("DocEntry");

                if (docEntry > 0)
                {
                    DataTable JrnLinesDT = BDOSARDownPaymentVATAccrual.createAdditionalEntries(null, oGeneralData, null, "", 0);
                    BDOSARDownPaymentVATAccrual.JrnEntry(docEntry.ToString(), docEntry.ToString(), DocDate, JrnLinesDT, out errorText);
                    if (errorText != null)
                    {
                        CommonFunctions.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                    }
                    else
                    {
                        CommonFunctions.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                    }
                }

            }
            catch
            {
                CommonFunctions.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
            }



            return docEntry;

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

        public static void uiApp_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
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

                    chooseFromList(oForm, pVal.BeforeAction, oCFLEvento, out errorText);

                }

                if ((pVal.ItemUID == "InCheck" || pVal.ItemUID == "InUncheck") && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.BeforeAction == false)
                {
                    checkUncheck(oForm, pVal.ItemUID, "", out errorText);
                }

                if (pVal.ItemUID == "DocPstDt" && pVal.ItemChanged && pVal.BeforeAction == false)
                {
                    SAPbouiCOM.EditText oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("DocPstDt").Specific;
                    DateTime EndDateOp = DateTime.ParseExact(oEditText.Value, "yyyyMMdd", null);
                    EndDateOp = new DateTime(EndDateOp.Year, EndDateOp.Month, 1).AddMonths(1).AddDays(-1);
                    oEditText.Value = EndDateOp.ToString("yyyyMMdd");

                    fillMTRInvoice(oForm);
                }
                if (pVal.ItemUID == "Fill" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.BeforeAction == false)
                {
                    fillMTRInvoice(oForm);
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
                        SetInvDocsMatrixRowBackColor(oForm, row, out errorText);
                        oForm.Freeze(false);
                    }

                    if (pVal.ColUID == "InDetail" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.BeforeAction == false)
                    {
                        if (oForm.DataSources.DataTables.Item("InvoiceMTR").GetValue("InDetail", pVal.Row - 1) != "")
                        {
                            openDetails(oForm, pVal.Row);
                        }
                    }

                }
                if (pVal.ItemUID == "CreatDocmt" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.BeforeAction == false)
                {
                    createPaymentDocuments(oForm);

                }
            }
        }

        public static void uiApp_ItemEventAddForm(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                setVisibleFormItemsAddForm(oForm, out errorText);

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                {
                    SAPbouiCOM.IChooseFromListEvent oCFLEvento = null;
                    oCFLEvento = ((SAPbouiCOM.IChooseFromListEvent)(pVal));

                    chooseFromListAddForm(oForm, oCFLEvento, pVal.ItemUID, pVal.BeforeAction, out errorText);
                }

                if (pVal.ItemUID == "3" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK & pVal.BeforeAction == false)
                {
                    SAPbouiCOM.DataTable ItemsMTR = OkAddForm(oForm, out errorText);

                    if (ItemsMTR != null)
                    {
                        int row = 0;

                        while (row < ItemsDT.Rows.Count)
                        {
                            if (DocEntryAdd == ItemsDT.Rows[row]["DocEntry"].ToString())
                            {
                                ItemsDT.Rows.Remove(ItemsDT.Rows[row]);
                            }
                            else
                            {
                                row++;
                            }
                        }

                        int count = ItemsMTR.Rows.Count;
                        row = 0;

                        while (row < count)
                        {
                            DataRow newDTRow = ItemsDT.Rows.Add();
                            newDTRow["DocEntry"] = ItemsMTR.GetValue("DocEntry", row);
                            newDTRow["ItemCode"] = ItemsMTR.GetValue("ItemCode", row);
                            newDTRow["Dscptn"] = ItemsMTR.GetValue("Dscptn", row);
                            //newDTRow["Qnty"] = ItemsMTR.GetValue("Qnty", row).ToString(CultureInfo.InvariantCulture);
                            newDTRow["GrsAmnt"] = ItemsMTR.GetValue("GrsAmnt", row).ToString(CultureInfo.InvariantCulture);
                            newDTRow["VatGrp"] = ItemsMTR.GetValue("VatGrp", row);
                            newDTRow["VatAmnt"] = ItemsMTR.GetValue("VatAmnt", row).ToString(CultureInfo.InvariantCulture);

                            row++;
                        }

                        oForm.Close();
                    }
                }

                if (pVal.ItemUID == "addMTRB" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK & pVal.BeforeAction == false)
                {
                    addMatrixRowAddForm(oForm, out errorText);
                }

                if (pVal.ItemUID == "delMTRB" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK & pVal.BeforeAction == false)
                {
                    delMatrixRowAddForm(oForm, out errorText);
                }

                if (pVal.ItemUID == "ItemsMTR" & pVal.ItemChanged & pVal.BeforeAction == false)
                {
                    fillTotalAmountsAddForm(oForm, out errorText);
                }

                if (pVal.ItemUID == "ItemsMTR" && (pVal.ColUID == "GrsAmnt" || pVal.ColUID == "VatGrp") & pVal.ItemChanged & pVal.BeforeAction == false)
                {
                    oForm.Freeze(true);

                    SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("ItemsMTR").Specific;
                    string VAtGroup = oMatrix.Columns.Item("VatGrp").Cells.Item(pVal.Row).Specific.Value;
                    decimal GrossAmnt = (FormsB1.cleanStringOfNonDigits(oMatrix.Columns.Item("GrsAmnt").Cells.Item(pVal.Row).Specific.Value));
                    decimal VatAmount = 0;
                    int row = pVal.Row;
                    decimal VatRate = CommonFunctions.GetVatGroupRate(VAtGroup, "");

                    VatAmount = CommonFunctions.roundAmountByGeneralSettings(GrossAmnt * VatRate / (100 + VatRate), "Sum");
                    oMatrix.Columns.Item("VatAmnt").Cells.Item(row).Specific.String = FormsB1.ConvertDecimalToStringForEditboxStrings(VatAmount);

                    oForm.Freeze(false);
                }
            }


        }

        public static void SetInvDocsMatrixRowBackColor(SAPbouiCOM.Form oForm, int row, out string errorText)
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
                        if (oMatrix.GetCellSpecific("DocTotal", i).Value != oMatrix.GetCellSpecific("Total", i).Value)
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

        private static void checkUncheck(SAPbouiCOM.Form oForm, string CheckOperation, string type, out string errorText)
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

        private static void createPaymentDocuments(SAPbouiCOM.Form oForm)
        {
            int answer = Program.uiApp.MessageBox(BDOSResources.getTranslate("CreateVATAccrualDocuments") + "?", 1, BDOSResources.getTranslate("Yes"), BDOSResources.getTranslate("No"), "");

            if (answer == 2)
            {
                return;
            }

            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("InvoiceMTR").Specific;

            SAPbouiCOM.EditText oEditTextDocDate = (SAPbouiCOM.EditText)oForm.Items.Item("DocPstDt").Specific;
            String DocDateS = oEditTextDocDate.Value;
            DateTime DocDate = Convert.ToDateTime(DateTime.ParseExact(DocDateS, "yyyyMMdd", CultureInfo.InvariantCulture));

            DataTable AccountHeader = new DataTable();
            DataRow headerLine = AccountHeader.Rows.Add();

            AccountHeader.Columns.Add("CardCode");
            AccountHeader.Columns.Add("DocEntry");
            AccountHeader.Columns.Add("DocDate");
            AccountHeader.Columns.Add("CardName");
            AccountHeader.Columns.Add("Total");


            for (int row = 1; row <= oMatrix.RowCount; row++)
            {

                bool checkedLine = oMatrix.GetCellSpecific("CheckBox", row).Checked;
                string DocEntVT = oMatrix.GetCellSpecific("DocEntVT", row).Value;

                if (checkedLine && DocEntVT == "")
                {
                    string InvDocEntry = oMatrix.GetCellSpecific("DocEntry", row).Value;
                    string CardCode = oMatrix.GetCellSpecific("CardCode", row).Value;
                    string CardName = oMatrix.GetCellSpecific("CardName", row).Value;
                    string Total = oMatrix.GetCellSpecific("Total", row).Value;


                    headerLine["DocEntry"] = InvDocEntry;
                    headerLine["CardCode"] = CardCode;
                    headerLine["CardName"] = CardName;
                    headerLine["Total"] = Total;
                    headerLine["DocDate"] = DocDateS;

                    createPaymentDocument(oForm, headerLine);

                }
            }



            fillMTRInvoice(oForm);

        }

        private static void chooseFromList(SAPbouiCOM.Form oForm, bool BeforeAction, SAPbouiCOM.IChooseFromListEvent oCFLEvento, out string errorText)
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
                            fillMTRInvoice(oForm);
                        }



                    }
                    catch
                    {
                        fillMTRInvoice(oForm);
                    }

                }
            }
        }

        public static void fillMTRInvoice(SAPbouiCOM.Form oForm)
        {

            SAPbouiCOM.EditText oEditTextDocDate = (SAPbouiCOM.EditText)oForm.Items.Item("DocPstDt").Specific;
            String DocDateS = oEditTextDocDate.Value;
            DateTime date = Convert.ToDateTime(DateTime.ParseExact(DocDateS, "yyyyMMdd", CultureInfo.InvariantCulture));

            string dateES = (new DateTime(date.Year, date.Month, 1)).ToString("yyyyMMdd");
            string dateEE = (new DateTime(date.Year, date.Month, DateTime.DaysInMonth(date.Year, date.Month))).ToString("yyyyMMdd");

            string cardCodeE = oForm.Items.Item("BPCode").Specific.Value;
            cardCodeE = cardCodeE == "" ? "1=1" : @" ""ODPI"".""CardCode"" = '" + cardCodeE + "'";

            SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("InvoiceMTR");

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string errorText = "";


            string query = @"select 
                    ""ODPI"".""DocEntry"",
                    ""ODPI"".""CANCELED"",
                    ""ODPI"".""DocNum"",
                    ""ODPI"".""DocDate"",
                    ""ODPI"".""CardCode"",
                    ""ODPI"".""CardName"",
					IFNULL(""OCRD"".""LicTradNum"",'') as ""LicTradNum"",
                    ""SumApplied""-IFNULL(""ITR1"".""ReconSum"",0) as ""Total"",
                    ""@BDOSARDV"".""U_GrsAmnt"" as ""DocTotal"",                    
                    ""@BDOSARDV"".""U_VatAmount"" as ""DocVAtTotal"",                    
                    IFNULL(""@BDOSARDV"".""DocEntry"",0) as ""DocEntVT"",
                    IFNULL(""@BDOSARDV"".""DocNum"",0) as ""DocNumVT""

                    from ""ODPI""
                    left join ""@BDOSARDV"" on ""ODPI"".""DocEntry""= ""@BDOSARDV"".""U_baseDoc"" and ""@BDOSARDV"".""U_baseDocT""=203 and ""@BDOSARDV"".""Canceled""='N'
                    inner join ""OCRD"" on ""ODPI"".""CardCode""= ""OCRD"".""CardCode""                     
                    
                    left join (select
	                 ""RCT2"".""DocEntry"",
	                 MAX(""RCT2"".""SumApplied"") as ""SumApplied"",
	                 SUM(""ITR1"".""ReconSum"") as ""ReconSum"" 
	                from ""RCT2""
                    inner join ""ORCT"" on ""ORCT"".""DocEntry"" = ""RCT2"".""DocNum"" and ""ORCT"".""Canceled"" = 'N'
	                left join (select
                    ""ITR1"".""LineSeq"",
                    ""ITR1"".""SrcObjAbs"",
                    ""ITR1"".""SrcObjTyp"",

                     ""ITR1"".""ReconSum"" 
		                from ""ITR1"" 
		                inner join ""OITR"" on ""ITR1"".""ReconNum"" = ""OITR"".""ReconNum"" 
		                and ""OITR"".""ReconDate""<='" + dateEE + @"' 
		                and ""OITR"".""Canceled"" = 'N' 
		                and ""OITR"".""ReconType""<>'3') as ""ITR1"" on ""RCT2"".""InvoiceId"" = ""ITR1"".""LineSeq"" 
	                and ""RCT2"".""DocNum"" = ""ITR1"".""SrcObjAbs"" 
	                and ""ITR1"".""SrcObjTyp"" = 24 
	                and ""RCT2"".""InvType"" = 203 
	                group by ""RCT2"".""DocEntry"") as ""ITR1""  on  ""ITR1"".""DocEntry"" = ""ODPI"".""DocEntry""                   
                    

                    where ""DocStatus""='C' and (""DocTotal"">""DpmAppl"" or IFNULL(""@BDOSARDV"".""DocEntry"",0) <>0) and ""ODPI"".""DocDate"" >= '" + dateES + @"' and ""ODPI"".""DocDate"" <= '" + dateEE + @"' and " + cardCodeE + @"
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
                decimal TotalPayment = 0;
                decimal DocTotal = 0;
                decimal DocVatTotal = 0;
                string CANCELED = "";

                ItemsDT = new DataTable();

                ItemsDT.Columns.Add("DocEntry");
                ItemsDT.Columns.Add("ItemCode");
                ItemsDT.Columns.Add("Dscptn");
                //ItemsDT.Columns.Add("Qnty");
                ItemsDT.Columns.Add("GrsAmnt");
                ItemsDT.Columns.Add("VatGrp");
                ItemsDT.Columns.Add("VatAmnt");


                while (!oRecordSet.EoF)
                {
                    DocEntry = oRecordSet.Fields.Item("DocEntry").Value == 0 ? "" : oRecordSet.Fields.Item("DocEntry").Value.ToString();
                    CANCELED = oRecordSet.Fields.Item("CANCELED").Value;
                    DocNum = oRecordSet.Fields.Item("DocNum").Value == 0 ? "" : oRecordSet.Fields.Item("DocNum").Value.ToString();
                    DocDate = oRecordSet.Fields.Item("DocDate").Value.ToString("yyyyMMdd");
                    DocEntVT = oRecordSet.Fields.Item("DocEntVT").Value == 0 ? "" : oRecordSet.Fields.Item("DocEntVT").Value.ToString();

                    TotalPayment = Convert.ToDecimal(oRecordSet.Fields.Item("Total").Value, CultureInfo.InvariantCulture);
                    DocTotal = Convert.ToDecimal(oRecordSet.Fields.Item("DocTotal").Value, CultureInfo.InvariantCulture);
                    DocVatTotal = Convert.ToDecimal(oRecordSet.Fields.Item("DocVatTotal").Value, CultureInfo.InvariantCulture);
                    CardCode = oRecordSet.Fields.Item("CardCode").Value;
                    CardName = oRecordSet.Fields.Item("CardName").Value;
                    LicTradNum = oRecordSet.Fields.Item("LicTradNum").Value;

                    if (CANCELED == "Y" && DocEntVT == "")
                    {
                        oRecordSet.MoveNext();
                        continue;
                    }

                    oDataTable.Rows.Add();
                    oDataTable.SetValue("LineNum", rowIndex, rowIndex + 1);
                    oDataTable.SetValue("CheckBox", rowIndex, "N");
                    oDataTable.SetValue("DocEntry", rowIndex, DocEntry);
                    oDataTable.SetValue("DocNum", rowIndex, DocNum);
                    oDataTable.SetValue("cancelled", rowIndex, CANCELED);
                    oDataTable.SetValue("DocEntVT", rowIndex, DocEntVT);
                    oDataTable.SetValue("DocDate", rowIndex, DocDate);
                    oDataTable.SetValue("CardCode", rowIndex, CardCode);
                    oDataTable.SetValue("CardName", rowIndex, CardName);
                    oDataTable.SetValue("LicTradNum", rowIndex, LicTradNum);
                    if (DocEntVT == "")
                    { oDataTable.SetValue("InDetail", rowIndex, "SPI_INFO"); }
                    else
                    {
                        oDataTable.SetValue("InDetail", rowIndex, "");
                    }
                    oDataTable.SetValue("Total", rowIndex, Convert.ToDouble(TotalPayment));
                    oDataTable.SetValue("DocTotal", rowIndex, Convert.ToDouble(DocTotal));
                    oDataTable.SetValue("DocVtTotal", rowIndex, Convert.ToDouble(DocVatTotal));

                    string VatGrp = oForm.Items.Item("VatGrp").Specific.Value;
                    decimal VatRate = CommonFunctions.GetVatGroupRate(VatGrp, "");

                    DataRow ItemsDTRow = ItemsDT.Rows.Add();
                    ItemsDTRow["DocEntry"] = DocEntry;
                    ItemsDTRow["ItemCode"] = "";
                    ItemsDTRow["Dscptn"] = oForm.Items.Item("Descrpt").Specific.Value;
                    //ItemsDTRow["Qnty"] = 1;
                    ItemsDTRow["GrsAmnt"] = TotalPayment.ToString(CultureInfo.InvariantCulture);
                    ItemsDTRow["VatGrp"] = VatGrp;
                    ItemsDTRow["VatAmnt"] = CommonFunctions.roundAmountByGeneralSettings(Convert.ToDecimal(TotalPayment * VatRate / (100 + VatRate), CultureInfo.InvariantCulture), "Sum").ToString(CultureInfo.InvariantCulture);


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
                oCreationPackage.UniqueID = "BDOSVAWizzForm";
                oCreationPackage.String = BDOSResources.getTranslate("ARVATAccrualWizard");
                oCreationPackage.Position = fatherMenuItem.SubMenus.Count - 1;

                menuItem = fatherMenuItem.SubMenus.AddEx(oCreationPackage);
            }
            catch
            {

            }
        }

        /// <summary>
        /// Additional items form
        /// </summary>
        /// <param name="Program.oCompany"></param>
        /// <param name="oForm"></param>
        /// <param name="oCFLEvento"></param>
        /// <param name="itemUID"></param>
        /// <param name="beforeAction"></param>
        /// <param name="errorText"></param>

        public static void chooseFromListAddForm(SAPbouiCOM.Form oForm, SAPbouiCOM.IChooseFromListEvent oCFLEvento, string itemUID, bool beforeAction, out string errorText)
        {
            errorText = null;

            try
            {
                string sCFL_ID = oCFLEvento.ChooseFromListUID;
                SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(sCFL_ID);

                if (beforeAction == true)
                {
                    if (sCFL_ID == "BaseDoc_CFL")
                    {
                        oCFL = oForm.ChooseFromLists.Item("BaseDoc_CFL");

                        SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();
                        SAPbouiCOM.Condition oCon = oCons.Add();
                        oCon.Alias = "CardCode";
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = oForm.Items.Item("cardCodeE").Specific.Value; //მყიდველი

                        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                        oCon = oCons.Add();
                        oCon.Alias = "DocDate"; //Lock Manual Transaction (Control Account)
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_LESS_EQUAL;
                        oCon.CondVal = oForm.Items.Item("DocDate").Specific.Value;

                        oCFL.SetConditions(oCons);
                    }

                }
                else
                {
                    SAPbouiCOM.DataTable oDataTable = null;
                    oDataTable = oCFLEvento.SelectedObjects;

                    if (oDataTable != null)
                    {
                        if (sCFL_ID == "Item_CFL")
                        {
                            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("ItemsMTR").Specific));
                            string ItemCode = Convert.ToString(oDataTable.GetValue("ItemCode", 0));
                            string ItemName = Convert.ToString(oDataTable.GetValue("ItemName", 0));

                            try
                            {
                                oMatrix.GetCellSpecific("Dscptn", oCFLEvento.Row).Value = ItemName;
                            }
                            catch { }

                            try
                            {
                                oMatrix.GetCellSpecific("ItemCode", oCFLEvento.Row).Value = ItemCode;
                            }
                            catch { }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                string exsd = ex.Message;
            }
        }

        private static SAPbouiCOM.DataTable OkAddForm(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("ItemsMTR").Specific));
            decimal SumValue = FormsB1.cleanStringOfNonDigits(oMatrix.Columns.Item("GrsAmnt").ColumnSetting.SumValue);
            decimal Total = FormsB1.cleanStringOfNonDigits(oForm.Items.Item("Total").Specific.Value);
            if (SumValue != Total)
            {
                return null;
            }
            else
            {
                oMatrix.FlushToDataSource();
                return oForm.DataSources.DataTables.Item(0);
            }
        }

        private static void addMatrixRowAddForm(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            oForm.Freeze(true);

            try
            {

                SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("ItemsMTR").Specific));
                int index = 0;
                if (oMatrix.RowCount == 0)
                {
                    index = 1;
                }
                else
                {
                    index = oMatrix.RowCount + 1;
                }


                oMatrix.AddRow(1, -1);
                oMatrix.AutoResizeColumns();
                //oMatrix.Columns.Item("LineID").Cells.Item(oMatrix.RowCount).Specific.Value = index;
                oMatrix.GetCellSpecific("LineID", oMatrix.RowCount).Value = index;



            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
            finally
            {
                GC.Collect();
                oForm.Freeze(false);
            }
        }

        public static void delMatrixRowAddForm(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            oForm.Freeze(true);

            try
            {
                SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("ItemsMTR").Specific));
                oMatrix.FlushToDataSource();
                int firstRow = 0;
                int row = 0;
                int deletedRowCount = 0;

                while (row != -1)
                {
                    row = oMatrix.GetNextSelectedRow(firstRow, SAPbouiCOM.BoOrderType.ot_RowOrder);
                    if (row > -1)
                    {
                        deletedRowCount++;
                        oMatrix.DeleteRow(row);

                        firstRow = row;
                    }
                }

                oMatrix.FlushToDataSource();
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE && deletedRowCount > 0)
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
            finally
            {
                GC.Collect();
                oForm.Freeze(false);
            }
        }

        private static void fillTotalAmountsAddForm(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = "";

            SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("ItemsMTR").Specific;
            decimal U_GrsAmnt = 0;
            decimal U_VatAmount = 0;
            for (int i = 1; i <= oMatrix.RowCount; i++)
            {
                U_GrsAmnt = U_GrsAmnt + Convert.ToDecimal(oMatrix.GetCellSpecific("GrsAmnt", i).Value, CultureInfo.InvariantCulture);
                U_VatAmount = U_VatAmount + Convert.ToDecimal(oMatrix.GetCellSpecific("VatAmnt", i).Value, CultureInfo.InvariantCulture);
            }
            string U_GrsAmnts = U_GrsAmnt.ToString(CultureInfo.InvariantCulture);
            string U_VatAmounts = U_VatAmount.ToString(CultureInfo.InvariantCulture);

            //oForm.Items.Item("GrsAmnt").Specific.Value = U_GrsAmnts;
            //oForm.Items.Item("VatAmount").Specific.Value = U_VatAmounts;


        }

        public static void openDetails(SAPbouiCOM.Form oFormDoc, int rowMatrix)
        {
            int formHeight = Program.uiApp.Desktop.Width;
            int formWidth = Program.uiApp.Desktop.Width;
            string errorText = "";
            SAPbouiCOM.Form oForm = null;

            //ფორმის აუცილებელი თვისებები
            Dictionary<string, object> formProperties = new Dictionary<string, object>();
            formProperties.Add("UniqueID", "BDOSVATADD");
            formProperties.Add("BorderStyle", SAPbouiCOM.BoFormBorderStyle.fbs_Sizable);
            formProperties.Add("Title", BDOSResources.getTranslate("ARVATAccrualWizard") + " (" + BDOSResources.getTranslate("InDetail") + ")");
            //formProperties.Add("Left", (Program.uiApp.Desktop.Width - formWidth) / 2);
            formProperties.Add("Width", formWidth);
            //formProperties.Add("Top", (Program.uiApp.Desktop.Height - formHeight) / 3);
            formProperties.Add("Height", formHeight);

            bool newForm;
            bool formExist = FormsB1.createForm(formProperties, out oForm, out newForm, out errorText);

            if (errorText != null)
            {
                return;
            }

            if (formExist == true)
            {
                if (newForm == true)
                {
                    errorText = null;
                    Dictionary<string, object> formItems;
                    string itemName = "";
                    NumberFormatInfo Nfi = new NumberFormatInfo() { NumberDecimalSeparator = "." };
                    DocEntryAdd = "";

                    int left = 6;
                    int top = 6;
                    int height_e = 15;
                    int height = oForm.ClientHeight - top - 8 * height_e - 1 - 30;
                    int width = oForm.ClientWidth;

                    int left_s = 6;


                    bool multiSelection = false;

                    string objectTypeItem = "4"; //SAPbouiCOM.BoLinkedObject.lf_Invoice
                    string uniqueID_lf_ItemCFL = "Item_CFL";
                    FormsB1.addChooseFromList(oForm, multiSelection, objectTypeItem, uniqueID_lf_ItemCFL);

                    string objectTypeCardCode = "2"; //SAPbouiCOM.BoLinkedObject.lf_BusinessPartner, Business Partner object 
                    string uniqueID_lf_BusinessPartnerCFL = "BusinessPartner_CFL";
                    FormsB1.addChooseFromList(oForm, multiSelection, objectTypeCardCode, uniqueID_lf_BusinessPartnerCFL);

                    //პირობის დადება ბიზნესპარტნიორის არჩევის სიაზე
                    SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(uniqueID_lf_BusinessPartnerCFL);
                    SAPbouiCOM.Conditions oCons = oCFL.GetConditions();
                    SAPbouiCOM.Condition oCon = oCons.Add();
                    oCon.Alias = "CardType";
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCon.CondVal = "C"; //მყიდველი
                    oCFL.SetConditions(oCons);

                    formItems = new Dictionary<string, object>();
                    itemName = "BPCodeS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height_e);
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
                    formItems.Add("Left", left_s + 5 + 120);
                    formItems.Add("Width", 150);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height_e);
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
                    formItems.Add("Left", left_s + 5 + 100);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height_e);
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


                    oForm.Items.Item("BPCode").Specific.Value = oFormDoc.DataSources.DataTables.Item("InvoiceMTR").GetValue("CardCode", rowMatrix - 1); ;


                    formItems = new Dictionary<string, object>();
                    itemName = "Total"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SUM);
                    formItems.Add("TableName", "");
                    formItems.Add("Length", 20);
                    formItems.Add("Size", 20);
                    formItems.Add("Alias", "Total");
                    formItems.Add("Bound", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("Left", left_s + 5 + 120 + 5 + 150);
                    formItems.Add("Width", 150);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height_e);
                    formItems.Add("UID", itemName);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("FromPane", 0);
                    formItems.Add("ToPane", 0);


                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    oForm.Items.Item("Total").Specific.Value = oFormDoc.DataSources.DataTables.Item("InvoiceMTR").GetValue("Total", rowMatrix - 1).ToString(CultureInfo.InvariantCulture);


                    top = top + height_e + 5;

                    formItems = new Dictionary<string, object>();
                    itemName = "addMTRB"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height_e);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("Add"));

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "delMTRB"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left_s + 100 + 1);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height_e);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("Delete"));

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    top = top + height_e + 5;

                    formItems = new Dictionary<string, object>();
                    itemName = "ItemsMTR"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
                    formItems.Add("Left", left);
                    formItems.Add("Width", width);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("DisplayDesc", true);
                    formItems.Add("AffectsFormMode", false);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }



                    SAPbouiCOM.DataTable oDataTableItems;

                    oDataTableItems = oForm.DataSources.DataTables.Add("ItemsMTR");

                    oDataTableItems.Columns.Add("LineID", SAPbouiCOM.BoFieldsType.ft_Integer, 50);
                    oDataTableItems.Columns.Add("DocEntry", SAPbouiCOM.BoFieldsType.ft_Text, 50);
                    oDataTableItems.Columns.Add("ItemCode", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 50);
                    oDataTableItems.Columns.Add("Dscptn", SAPbouiCOM.BoFieldsType.ft_Text, 50);
                    //oDataTableItems.Columns.Add("Qnty", SAPbouiCOM.BoFieldsType.ft_Quantity, 50);
                    oDataTableItems.Columns.Add("GrsAmnt", SAPbouiCOM.BoFieldsType.ft_Sum, 50);
                    oDataTableItems.Columns.Add("VatGrp", SAPbouiCOM.BoFieldsType.ft_Text, 50);
                    oDataTableItems.Columns.Add("VatAmnt", SAPbouiCOM.BoFieldsType.ft_Sum, 50);
                    oDataTableItems.Columns.Add("Error", SAPbouiCOM.BoFieldsType.ft_Text, 50);


                    string DocEntryDoc = oFormDoc.DataSources.DataTables.Item("InvoiceMTR").GetValue("DocEntry", rowMatrix - 1);
                    DocEntryAdd = DocEntryDoc;

                    int count = ItemsDT.Rows.Count;
                    int rowIndex = 0;
                    for (int row = 0; row < count; row++)
                    {
                        string DocEntry = ItemsDT.Rows[row]["DocEntry"].ToString();


                        if (DocEntry == DocEntryDoc)
                        {
                            oDataTableItems.Rows.Add();
                            oDataTableItems.SetValue("LineID", rowIndex, rowIndex + 1);
                            oDataTableItems.SetValue("DocEntry", rowIndex, DocEntry);
                            oDataTableItems.SetValue("ItemCode", rowIndex, ItemsDT.Rows[row]["ItemCode"].ToString());
                            oDataTableItems.SetValue("Dscptn", rowIndex, ItemsDT.Rows[row]["Dscptn"].ToString());
                            //oDataTableItems.SetValue("Qnty", rowIndex, ItemsDT.Rows[row]["Qnty"].ToString());
                            oDataTableItems.SetValue("GrsAmnt", rowIndex, ItemsDT.Rows[row]["GrsAmnt"]);
                            oDataTableItems.SetValue("VatGrp", rowIndex, ItemsDT.Rows[row]["VatGrp"].ToString());
                            oDataTableItems.SetValue("VatAmnt", rowIndex, ItemsDT.Rows[row]["VatAmnt"]);

                            rowIndex++;
                        }

                    }



                    SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("ItemsMTR").Specific));
                    oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single;
                    SAPbouiCOM.Columns oColumns = oMatrix.Columns;
                    SAPbouiCOM.Column oColumn;
                    string UID = "ItemsMTR";

                    foreach (SAPbouiCOM.DataColumn column in oDataTableItems.Columns)
                    {
                        string columnName = column.Name;
                        if (columnName == "VatGrp")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("VatGroup");
                            oColumn.Editable = true;
                            oColumn.DataBind.Bind(UID, columnName);
                            oColumn.AffectsFormMode = false;

                            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            string query = "select * " +
                            "FROM  \"OVTG\" " +
                            "WHERE \"Category\"='O'";

                            oRecordSet.DoQuery(query);
                            while (!oRecordSet.EoF)
                            {
                                oColumn.ValidValues.Add(oRecordSet.Fields.Item("Code").Value, oRecordSet.Fields.Item("Code").Value);
                                oRecordSet.MoveNext();
                            }

                        }

                        else if (columnName == "ItemCode")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate(columnName);
                            oColumn.Editable = true;
                            oColumn.DataBind.Bind(UID, columnName);
                            oColumn.AffectsFormMode = false;
                            oColumn.ChooseFromListUID = "Item_CFL";
                            oColumn.ChooseFromListAlias = "ItemCode";
                            SAPbouiCOM.LinkedButton oLink = oColumn.ExtendedObject;
                            oLink.LinkedObjectType = "4";

                        }

                        else if (columnName == "DocEntry")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate(columnName);
                            oColumn.Editable = false;
                            oColumn.Visible = false;
                            oColumn.DataBind.Bind(UID, columnName);
                            oColumn.AffectsFormMode = false;
                        }
                        else if (columnName == "LineID")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = "";
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind(UID, columnName);
                            oColumn.AffectsFormMode = false;

                        }
                        else if (columnName == "GrsAmnt")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("GrossAmount");
                            oColumn.Editable = true;
                            oColumn.DataBind.Bind(UID, columnName);
                            oColumn.AffectsFormMode = false;
                            oColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                        }

                        else if (columnName == "VatAmnt")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("VatAmount");
                            oColumn.Editable = true;
                            oColumn.DataBind.Bind(UID, columnName);
                            oColumn.AffectsFormMode = false;
                            oColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                        }
                        else if (columnName == "Dscptn")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Description");
                            oColumn.Editable = true;
                            oColumn.DataBind.Bind(UID, columnName);
                            oColumn.AffectsFormMode = false;
                            oColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                        }
                        else
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate(columnName);
                            oColumn.Editable = true;
                            oColumn.DataBind.Bind(UID, columnName);
                            oColumn.AffectsFormMode = false;
                            oColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                        }
                    }
                    oMatrix.Clear();
                    oMatrix.LoadFromDataSource();
                    oMatrix.AutoResizeColumns();

                    top = top + height + 5;

                    itemName = "3";
                    formItems = new Dictionary<string, object>();
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", 5);
                    formItems.Add("Width", 65);
                    formItems.Add("Top", oForm.ClientHeight - 20);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", "OK");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    itemName = "2";
                    formItems = new Dictionary<string, object>();
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", 75);
                    formItems.Add("Width", 65);
                    formItems.Add("Top", oForm.ClientHeight - 20);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("Close"));

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

        public static void setVisibleFormItemsAddForm(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            try
            {
                oForm.Items.Item("BPCode").Enabled = false;
                oForm.Items.Item("Total").Enabled = false;


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
    }
}
