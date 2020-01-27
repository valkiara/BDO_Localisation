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
    class BDOSVATAccrualWizard
    {
        static DataTable ItemsDT = null;
        static string DocEntryAdd = "";

        public static void createForm()
        {
            string errorText;
            int formHeight = Program.uiApp.Desktop.Height;
            int formWidth = Program.uiApp.Desktop.Width;

            //ფორმის აუცილებელი თვისებები
            Dictionary<string, object> formProperties = new Dictionary<string, object>();
            formProperties.Add("UniqueID", "BDOSVAWizzForm");
            formProperties.Add("BorderStyle", SAPbouiCOM.BoFormBorderStyle.fbs_Sizable);
            formProperties.Add("Title", BDOSResources.getTranslate("ARDownPaymentVATAccrualWizard"));
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
                    string itemName;

                    int width_s = 130;
                    int width_e = 130;
                    int left_s = 6;
                    int left_e = left_s + width_s + 20;
                    int left_s2 = 295;
                    int left_e2 = left_s2 + width_s + 20;
                    int height = 15;
                    int top = 10;

                    FormsB1.addChooseFromList(oForm, false, "2", "BusinessPartner_CFL");

                    //პირობის დადება ბიზნესპარტნიორის არჩევის სიაზე
                    SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item("BusinessPartner_CFL");
                    SAPbouiCOM.Conditions oCons = oCFL.GetConditions();
                    SAPbouiCOM.Condition oCon = oCons.Add();
                    oCon.Alias = "CardType";
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCon.CondVal = "C"; //მყიდველი
                    oCFL.SetConditions(oCons);

                    formItems = new Dictionary<string, object>();
                    itemName = "DocPsDtS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s);
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
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Left", left_e);
                    formItems.Add("Width", width_e);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("ValueEx", new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1).AddMonths(1).AddDays(-1).ToString("yyyyMMdd"));

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "VatGrpS"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left_s2);
                    formItems.Add("Width", width_s);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("VatGroup"));

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
                    formItems.Add("Alias", "VatGrp");
                    formItems.Add("Bound", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                    formItems.Add("Left", left_e2);
                    formItems.Add("Width", width_e);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("DisplayDesc", true);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    string query = "SELECT \"Code\" " +
                    "FROM  \"OVTG\" " +
                    "WHERE \"Category\" = 'O'";
                    oRecordSet.DoQuery(query);

                    while (!oRecordSet.EoF)
                    {
                        oForm.Items.Item("VatGrp").Specific.ValidValues.Add(oRecordSet.Fields.Item("Code").Value, oRecordSet.Fields.Item("Code").Value);
                        oRecordSet.MoveNext();
                    }

                    oForm.Items.Item("VatGrp").Specific.Select(CommonFunctions.getOADM("DfSVatItem"));

                    top += height + 1;

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
                    formItems.Add("LinkedObjectType", "2"); //Business Partner

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
                    itemName = "Descrpt";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("Length", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
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
                    oForm.Items.Item("Descrpt").Specific.Value = "Service";

                    top += 2 * height + 1;

                    itemName = "checkB";
                    formItems = new Dictionary<string, object>();
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

                    left_s = left_s + 20;

                    itemName = "unCheckB";
                    formItems = new Dictionary<string, object>();
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left_s);
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

                    left_s = left_s + 20;

                    itemName = "fillB";
                    formItems = new Dictionary<string, object>();
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", 65);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("Fill"));

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    itemName = "createDocB";
                    formItems = new Dictionary<string, object>();
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left_s + 65 + 2);
                    formItems.Add("Width", 65 * 2);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("CreateDocuments"));

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    top = top + height + 1;
                    left_s = 6;

                    formItems = new Dictionary<string, object>();
                    itemName = "InvoiceMTR"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
                    formItems.Add("Left", left_s);
                    formItems.Add("Height", 150);
                    formItems.Add("Top", top);
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("InvoiceMTR").Specific;
                    oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
                    SAPbouiCOM.Columns oColumns = oMatrix.Columns;
                    SAPbouiCOM.Column oColumn;

                    SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Add("InvoiceMTR");

                    oDataTable.Columns.Add("LineNum", SAPbouiCOM.BoFieldsType.ft_Integer, 50); //ინდექსი 
                    oDataTable.Columns.Add("CheckBox", SAPbouiCOM.BoFieldsType.ft_Text, 1); // 
                    oDataTable.Columns.Add("DocEntry", SAPbouiCOM.BoFieldsType.ft_Text, 50); //ენთრი
                    oDataTable.Columns.Add("DocNum", SAPbouiCOM.BoFieldsType.ft_Text, 50); //ნომერი
                    oDataTable.Columns.Add("cancelled", SAPbouiCOM.BoFieldsType.ft_Text, 1); //ნომერი                    
                    oDataTable.Columns.Add("DocDate", SAPbouiCOM.BoFieldsType.ft_Date); //თარიღი
                    oDataTable.Columns.Add("CardCode", SAPbouiCOM.BoFieldsType.ft_Text, 50); //ენთრი
                    oDataTable.Columns.Add("CardName", SAPbouiCOM.BoFieldsType.ft_Text, 50); //ნომერი
                    oDataTable.Columns.Add("LicTradNum", SAPbouiCOM.BoFieldsType.ft_Text, 50); //ენთრი
                    oDataTable.Columns.Add("Total", SAPbouiCOM.BoFieldsType.ft_Sum); //თანხა
                    oDataTable.Columns.Add("DocEntVT", SAPbouiCOM.BoFieldsType.ft_Text, 50); //ენთრი
                    oDataTable.Columns.Add("DocTotal", SAPbouiCOM.BoFieldsType.ft_Sum); //თანხა
                    oDataTable.Columns.Add("DocVtTotal", SAPbouiCOM.BoFieldsType.ft_Sum); //თანხა
                    oDataTable.Columns.Add("InDetail", SAPbouiCOM.BoFieldsType.ft_Text, 20);
                    oDataTable.Columns.Add("Error", SAPbouiCOM.BoFieldsType.ft_Text, 50); //ენთრი

                    string UID = "InvoiceMTR";
                    SAPbouiCOM.LinkedButton oLink;

                    foreach (SAPbouiCOM.DataColumn column in oDataTable.Columns)
                    {
                        string columnName = column.Name;

                        if (columnName == "LineNum")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = "#";
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind(UID, columnName);
                        }
                        else if (columnName == "CheckBox")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
                            oColumn.TitleObject.Caption = "";
                            oColumn.ValOff = "N";
                            oColumn.ValOn = "Y";
                            oColumn.DataBind.Bind(UID, columnName);
                        }
                        else if (columnName == "DocEntry")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("ARDownPaymentRequest") + " (" + BDOSResources.getTranslate(columnName) + ")";
                            oColumn.Editable = false;
                            oLink = oColumn.ExtendedObject;
                            oLink.LinkedObjectType = "203";
                            oColumn.DataBind.Bind(UID, columnName);
                        }
                        else if (columnName == "DocNum")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("ARDownPaymentRequest") + " (" + BDOSResources.getTranslate(columnName) + ")";
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind(UID, columnName);
                        }
                        else if (columnName == "DocEntVT")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("ARDownPaymentVATAccrual");
                            oColumn.Editable = false;
                            oLink = oColumn.ExtendedObject;
                            oLink.LinkedObjectType = "UDO_F_BDO_ARDPV_D";
                            oColumn.DataBind.Bind(UID, columnName);
                        }
                        else if (columnName == "CardCode")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("CardCode") + " (" + BDOSResources.getTranslate("Code") + ")";
                            oColumn.Editable = false;
                            oLink = oColumn.ExtendedObject;
                            oLink.LinkedObjectType = "2";
                            oColumn.DataBind.Bind(UID, columnName);
                        }
                        else if (columnName == "InDetail")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_PICTURE);
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind(UID, columnName);
                        }
                        else if (columnName == "CardName")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("CardCode") + " (" + BDOSResources.getTranslate("Name") + ")";
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind(UID, columnName);
                        }
                        else if (columnName == "LicTradNum")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Tin");
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind(UID, columnName);
                        }
                        else if (columnName == "DocTotal")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Total");
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind(UID, columnName);
                        }
                        else if (columnName == "DocVtTotal")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("VatAmount");
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind(UID, columnName);
                        }
                        else
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate(columnName);
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind(UID, columnName);
                        }
                    }
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

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_CLOSE && pVal.BeforeAction)
                {
                    int answer = Program.uiApp.MessageBox(BDOSResources.getTranslate("CloseARDownPaymentVATAccrualWizard") + "?", 1, BDOSResources.getTranslate("Yes"), BDOSResources.getTranslate("No"), "");

                    if (answer != 1)
                        BubbleEvent = false;
                }

                else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_RESIZE && !pVal.BeforeAction)
                {
                    resizeForm(oForm);
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
                        if (pVal.ItemUID == "checkB" || pVal.ItemUID == "unCheckB")
                            checkUncheck(oForm, pVal.ItemUID);
                        else if (pVal.ItemUID == "fillB")
                            fillMTRInvoice(oForm);
                        else if (pVal.ColUID == "InDetail" && pVal.Row > 0)
                        {
                            if (oForm.DataSources.DataTables.Item("InvoiceMTR").GetValue("InDetail", pVal.Row - 1) != "")
                                openDetails(oForm, pVal.Row);
                        }
                        else if (pVal.ItemUID == "createDocB")
                            createPaymentDocuments(oForm);
                    }
                }

                //else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED)
                //{
                //    if (pVal.ItemUID == "InvoiceMTR" && pVal.ColUID == "DocEntry")
                //        matrixColumnSetLinkedObjectTypeInvoicesMTR(oForm, pVal);
                //}

                else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
                {
                    if (!pVal.BeforeAction)
                    {
                        if (pVal.ItemUID == "InvoiceMTR" && pVal.Row > 0)
                        {
                            SetInvDocsMatrixRowBackColor(oForm, pVal.Row);
                        }
                    }
                }

                if (pVal.ItemChanged && !pVal.BeforeAction)
                {
                    if (pVal.ItemUID == "DocPstDt")
                    {
                        string docDateStr = oForm.DataSources.UserDataSources.Item("DocPstDt").ValueEx;
                        if (!string.IsNullOrEmpty(docDateStr))
                        {
                            DateTime endDateOp = DateTime.ParseExact(docDateStr, "yyyyMMdd", null);
                            endDateOp = new DateTime(endDateOp.Year, endDateOp.Month, 1).AddMonths(1).AddDays(-1);
                            oForm.DataSources.UserDataSources.Item("DocPstDt").ValueEx = endDateOp.ToString("yyyyMMdd");
                        }

                        fillMTRInvoice(oForm);
                    }
                }
            }
        }

        public static void resizeForm(SAPbouiCOM.Form oForm)
        {
            try
            {
                oForm.Freeze(true);

                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("InvoiceMTR").Specific;
                int mtrWidth = oForm.ClientWidth;
                oForm.Items.Item("InvoiceMTR").Width = mtrWidth;
                oForm.Items.Item("InvoiceMTR").Height = oForm.ClientHeight - 25;
                int columnsCount = oMatrix.Columns.Count - 2;
                oMatrix.Columns.Item("LineNum").Width = 19;
                oMatrix.Columns.Item("CheckBox").Width = 19;
                mtrWidth -= 38;
                mtrWidth /= columnsCount;

                foreach (SAPbouiCOM.Column column in oMatrix.Columns)
                {
                    if (column.UniqueID == "LineNum" || column.UniqueID == "CheckBox")
                        continue;
                    column.Width = mtrWidth;
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

        public static void uiApp_ItemEventAddForm(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_RESIZE && !pVal.BeforeAction)
                {
                    resizeFormAddForm(oForm);
                }

                else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                {
                    SAPbouiCOM.IChooseFromListEvent oCFLEvento = (SAPbouiCOM.IChooseFromListEvent)pVal;
                    chooseFromListAddForm(oForm, pVal, oCFLEvento);
                }

                else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK)
                {
                    if (!pVal.BeforeAction)
                    {
                        if (pVal.ItemUID == "3")
                        {
                            SAPbouiCOM.DataTable ItemsMTR = OkAddForm(oForm);

                            if (ItemsMTR != null)
                            {
                                int row = 0;
                                while (row < ItemsDT.Rows.Count)
                                {
                                    if (DocEntryAdd == ItemsDT.Rows[row]["DocEntry"].ToString())
                                        ItemsDT.Rows.Remove(ItemsDT.Rows[row]);
                                    else
                                        row++;
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
                        else if (pVal.ItemUID == "addMTRB")
                            addMatrixRowAddForm(oForm);
                        else if (pVal.ItemUID == "delMTRB")
                            delMatrixRowAddForm(oForm);
                    }
                }

                if (pVal.ItemChanged && !pVal.BeforeAction)
                {
                    if (pVal.ItemUID == "ItemsMTR")
                    {
                        if (pVal.ColUID == "GrsAmnt" || pVal.ColUID == "VatGrp")
                        {
                            try
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

                                fillTotalAmountsAddForm(oForm);
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
                    }
                }
            }
        }

        public static void SetInvDocsMatrixRowBackColor(SAPbouiCOM.Form oForm, int row)
        {
            try
            {
                oForm.Freeze(true);

                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("InvoiceMTR").Specific;

                if (oMatrix.RowCount > 0)
                {
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
                oForm.Freeze(false);
            }
        }

        public static void SetInvDocsMatrixRowCellColor(SAPbouiCOM.Form oForm)
        {
            try
            {
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("InvoiceMTR").Specific;

                if (oMatrix.RowCount > 0)
                {
                    for (int i = 1; i <= oMatrix.RowCount; i++)
                    {
                        if (oMatrix.GetCellSpecific("DocTotal", i).Value != oMatrix.GetCellSpecific("Total", i).Value)
                            oMatrix.CommonSetting.SetRowFontColor(i, FormsB1.getLongIntRGB(255, 0, 0));
                        else
                            oMatrix.CommonSetting.SetRowFontColor(i, FormsB1.getLongIntRGB(0, 0, 0));
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private static void checkUncheck(SAPbouiCOM.Form oForm, string checkOperation)
        {
            try
            {
                oForm.Freeze(true);

                SAPbouiCOM.CheckBox oCheckBox;
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("InvoiceMTR").Specific;

                for (int j = 1; j <= oMatrix.RowCount; j++)
                {
                    oCheckBox = oMatrix.Columns.Item("CheckBox").Cells.Item(j).Specific;

                    oCheckBox.Checked = (checkOperation == "checkB");
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

        //public static void matrixColumnSetLinkedObjectTypeInvoicesMTR(SAPbouiCOM.Form oForm, SAPbouiCOM.ItemEvent pVal)
        //{
        //    try
        //    {
        //        oForm.Freeze(true);

        //        if (pVal.BeforeAction)
        //        {
        //            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("InvoiceMTR").Specific;

        //            SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("InvoiceMTR");
        //            string docType = oDataTable.GetValue("DocType", pVal.Row - 1);

        //            SAPbouiCOM.Column oColumn;

        //            if (docType == "18")
        //            {
        //                oColumn = oMatrix.Columns.Item(pVal.ColUID);
        //                SAPbouiCOM.LinkedButton oLink = oColumn.ExtendedObject;
        //                oLink.LinkedObjectType = docType; //ARInvoice
        //            }
        //            if (docType == "204")
        //            {
        //                oColumn = oMatrix.Columns.Item(pVal.ColUID);
        //                SAPbouiCOM.LinkedButton oLink = oColumn.ExtendedObject;
        //                oLink.LinkedObjectType = docType; //ARInvoice
        //            }
        //            else if (docType == "163")
        //            {
        //                oColumn = oMatrix.Columns.Item(pVal.ColUID);
        //                SAPbouiCOM.LinkedButton oLink = oColumn.ExtendedObject;
        //                oLink.LinkedObjectType = docType; //ARCreditNote
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        throw new Exception(ex.Message);
        //    }
        //    finally
        //    {
        //        oForm.Freeze(false);
        //    }
        //}

        private static void createPaymentDocuments(SAPbouiCOM.Form oForm)
        {
            int answer = Program.uiApp.MessageBox(BDOSResources.getTranslate("CreateARDownPaymentVATAccrualDocuments") + "?", 1, BDOSResources.getTranslate("Yes"), BDOSResources.getTranslate("No"), "");

            if (answer == 2)
                return;

            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("InvoiceMTR").Specific;

            string docDateStr = oForm.DataSources.UserDataSources.Item("DocPstDt").ValueEx;
            if (string.IsNullOrEmpty(docDateStr))
            {
                string errorText = BDOSResources.getTranslate("TheFollowingFieldsAreMandatory") +
                    " : \"" + oForm.Items.Item("DocPsDtS").Specific.caption + "\"";

                Program.uiApp.SetStatusBarMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short);
                return;
            }
            DateTime DocDate = Convert.ToDateTime(DateTime.ParseExact(docDateStr, "yyyyMMdd", CultureInfo.InvariantCulture));

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
                    headerLine["DocDate"] = docDateStr;

                    createPaymentDocument(oForm, headerLine);
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

                }
                else
                {
                    SAPbouiCOM.DataTable oDataTable = oCFLEvento.SelectedObjects;

                    if (oDataTable != null)
                    {
                        if (oCFLEvento.ChooseFromListUID == "BusinessPartner_CFL")
                        {
                            string CardCode = oDataTable.GetValue("CardCode", 0);
                            LanguageUtils.IgnoreErrors<string>(() => oForm.Items.Item("BPCode").Specific.Value = CardCode);

                            fillMTRInvoice(oForm);
                        }
                    }
                    else
                        fillMTRInvoice(oForm);
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

        public static void fillMTRInvoice(SAPbouiCOM.Form oForm)
        {
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                string docDateStr = oForm.DataSources.UserDataSources.Item("DocPstDt").ValueEx;
                if (string.IsNullOrEmpty(docDateStr))
                {
                    string errorText = BDOSResources.getTranslate("TheFollowingFieldsAreMandatory") +
                        " : \"" + oForm.Items.Item("DocPsDtS").Specific.caption + "\"";

                    Program.uiApp.SetStatusBarMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short);
                    return;
                }

                DateTime date = Convert.ToDateTime(DateTime.ParseExact(docDateStr, "yyyyMMdd", CultureInfo.InvariantCulture));

                string dateES = new DateTime(date.Year, date.Month, 1).ToString("yyyyMMdd");
                string dateEE = new DateTime(date.Year, date.Month, DateTime.DaysInMonth(date.Year, date.Month)).ToString("yyyyMMdd");

                string cardCodeE = oForm.DataSources.UserDataSources.Item("BPCode").ValueEx;
                cardCodeE = cardCodeE == "" ? "1=1" : @" ""ODPI"".""CardCode"" = '" + cardCodeE + "'";

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

                SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("InvoiceMTR");
                oDataTable.Rows.Clear();

                oRecordSet.DoQuery(query);

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
                        oDataTable.SetValue("InDetail", rowIndex, "SPI_INFO");
                    else
                        oDataTable.SetValue("InDetail", rowIndex, "");

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

                oForm.Freeze(true);
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("InvoiceMTR").Specific;
                oMatrix.Clear();
                oMatrix.LoadFromDataSource();
                oForm.Update();

                SetInvDocsMatrixRowCellColor(oForm);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                oForm.Freeze(false);
                Marshal.ReleaseComObject(oRecordSet);
            }
        }

        public static void addMenus()
        {
            try
            {
                SAPbouiCOM.MenuItem fatherMenuItem = Program.uiApp.Menus.Item("2048");

                // Add a pop-up menu item
                SAPbouiCOM.MenuCreationParams oCreationPackage = (SAPbouiCOM.MenuCreationParams)Program.uiApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                oCreationPackage.Checked = false;
                oCreationPackage.Enabled = true;
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "BDOSVAWizzForm";
                oCreationPackage.String = BDOSResources.getTranslate("ARDownPaymentVATAccrualWizard");
                oCreationPackage.Position = fatherMenuItem.SubMenus.Count - 1;

                SAPbouiCOM.MenuItem menuItem = fatherMenuItem.SubMenus.AddEx(oCreationPackage);
            }
            catch
            {

            }
        }

        public static void chooseFromListAddForm(SAPbouiCOM.Form oForm, SAPbouiCOM.ItemEvent pVal, SAPbouiCOM.IChooseFromListEvent oCFLEvento)
        {
            try
            {
                oForm.Freeze(true);

                if (pVal.BeforeAction)
                {
                    if (oCFLEvento.ChooseFromListUID == "BaseDoc_CFL")
                    {
                        SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID);

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
                    SAPbouiCOM.DataTable oDataTable = oCFLEvento.SelectedObjects;

                    if (oDataTable != null)
                    {
                        if (oCFLEvento.ChooseFromListUID == "Item_CFL")
                        {
                            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("ItemsMTR").Specific;
                            string itemCode = oDataTable.GetValue("ItemCode", 0);
                            string itemName = oDataTable.GetValue("ItemName", 0);

                            LanguageUtils.IgnoreErrors<string>(() => oMatrix.GetCellSpecific("Dscptn", oCFLEvento.Row).Value = itemName);
                            LanguageUtils.IgnoreErrors<string>(() => oMatrix.GetCellSpecific("ItemCode", oCFLEvento.Row).Value = itemCode);
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
                oForm.Freeze(false);
            }
        }

        private static SAPbouiCOM.DataTable OkAddForm(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("ItemsMTR").Specific;
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

        private static void addMatrixRowAddForm(SAPbouiCOM.Form oForm)
        {
            try
            {
                oForm.Freeze(true);

                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("ItemsMTR").Specific;
                int index = 0;
                if (oMatrix.RowCount == 0)
                    index = 1;
                else
                    index = oMatrix.RowCount + 1;

                oMatrix.AddRow(1, -1);
                oMatrix.GetCellSpecific("LineID", oMatrix.RowCount).Value = index;
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

        public static void delMatrixRowAddForm(SAPbouiCOM.Form oForm)
        {
            try
            {
                oForm.Freeze(true);

                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("ItemsMTR").Specific;
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
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
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

        private static void fillTotalAmountsAddForm(SAPbouiCOM.Form oForm)
        {
            try
            {
                oForm.Freeze(true);

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

        public static void resizeFormAddForm(SAPbouiCOM.Form oForm)
        {
            try
            {
                oForm.Freeze(true);

                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("ItemsMTR").Specific;
                int mtrWidth = oForm.ClientWidth;
                oForm.Items.Item("ItemsMTR").Width = mtrWidth;
                oForm.Items.Item("ItemsMTR").Height = oForm.ClientHeight * 2 / 3;
                int columnsCount = oMatrix.Columns.Count - 2;
                oMatrix.Columns.Item("LineID").Width = 19;
                mtrWidth -= 38;
                mtrWidth /= columnsCount;

                foreach (SAPbouiCOM.Column column in oMatrix.Columns)
                {
                    if (column.UniqueID == "LineID" || column.UniqueID == "DocEntry")
                        continue;
                    column.Width = mtrWidth;
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

        public static void openDetails(SAPbouiCOM.Form oFormDoc, int rowMatrix)
        {
            string errorText;
            int formHeight = Program.uiApp.Desktop.Height;
            int formWidth = Program.uiApp.Desktop.Width;

            //ფორმის აუცილებელი თვისებები
            Dictionary<string, object> formProperties = new Dictionary<string, object>();
            formProperties.Add("UniqueID", "BDOSVATADD");
            formProperties.Add("BorderStyle", SAPbouiCOM.BoFormBorderStyle.fbs_Sizable);
            formProperties.Add("Title", BDOSResources.getTranslate("ARDownPaymentVATAccrualWizard") + " (" + BDOSResources.getTranslate("InDetail") + ")");
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
                    string itemName;
                    NumberFormatInfo Nfi = new NumberFormatInfo() { NumberDecimalSeparator = "." };
                    DocEntryAdd = "";

                    int width_s = 130;
                    int width_e = 130;
                    int left_s = 6;
                    int left_e = left_s + width_s + 20;
                    int left_s2 = 295;
                    int left_e2 = left_s2 + width_s + 20;
                    int height = 15;
                    int top = 10;

                    FormsB1.addChooseFromList(oForm, false, "4", "Item_CFL");
                    FormsB1.addChooseFromList(oForm, false, "2", "BusinessPartner_CFL");

                    //პირობის დადება ბიზნესპარტნიორის არჩევის სიაზე
                    SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item("BusinessPartner_CFL");
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
                    formItems.Add("Enabled", false);

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
                    oForm.Items.Item("BPCode").Specific.Value = oFormDoc.DataSources.DataTables.Item("InvoiceMTR").GetValue("CardCode", rowMatrix - 1); ;

                    formItems = new Dictionary<string, object>();
                    itemName = "Total"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SUM);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", "Total");
                    formItems.Add("Bound", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("Left", left_s2);
                    formItems.Add("Width", width_e);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Enabled", false);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }
                    oForm.Items.Item("Total").Specific.Value = oFormDoc.DataSources.DataTables.Item("InvoiceMTR").GetValue("Total", rowMatrix - 1).ToString(CultureInfo.InvariantCulture);

                    top += height + 5;

                    formItems = new Dictionary<string, object>();
                    itemName = "addMTRB"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left_s);
                    formItems.Add("Width", 70);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
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
                    formItems.Add("Left", left_s + 70 + 1);
                    formItems.Add("Width", 70);
                    formItems.Add("Top", top);
                    formItems.Add("Height", height);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("Delete"));

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    top += height + 1;

                    formItems = new Dictionary<string, object>();
                    itemName = "ItemsMTR"; //10 characters
                    formItems.Add("isDataSource", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
                    formItems.Add("Left", left_s);
                    formItems.Add("Top", top);
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        throw new Exception(errorText);
                    }

                    SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Add("ItemsMTR");
                    oDataTable.Columns.Add("LineID", SAPbouiCOM.BoFieldsType.ft_Integer, 50);
                    oDataTable.Columns.Add("DocEntry", SAPbouiCOM.BoFieldsType.ft_Text, 50);
                    oDataTable.Columns.Add("ItemCode", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 50);
                    oDataTable.Columns.Add("Dscptn", SAPbouiCOM.BoFieldsType.ft_Text, 50);
                    //oDataTable.Columns.Add("Qnty", SAPbouiCOM.BoFieldsType.ft_Quantity, 50);
                    oDataTable.Columns.Add("GrsAmnt", SAPbouiCOM.BoFieldsType.ft_Sum);
                    oDataTable.Columns.Add("VatGrp", SAPbouiCOM.BoFieldsType.ft_Text, 50);
                    oDataTable.Columns.Add("VatAmnt", SAPbouiCOM.BoFieldsType.ft_Sum);
                    oDataTable.Columns.Add("Error", SAPbouiCOM.BoFieldsType.ft_Text, 50);

                    string DocEntryDoc = oFormDoc.DataSources.DataTables.Item("InvoiceMTR").GetValue("DocEntry", rowMatrix - 1);
                    DocEntryAdd = DocEntryDoc;

                    int count = ItemsDT.Rows.Count;
                    int rowIndex = 0;
                    for (int row = 0; row < count; row++)
                    {
                        string DocEntry = ItemsDT.Rows[row]["DocEntry"].ToString();

                        if (DocEntry == DocEntryDoc)
                        {
                            oDataTable.Rows.Add();
                            oDataTable.SetValue("LineID", rowIndex, rowIndex + 1);
                            oDataTable.SetValue("DocEntry", rowIndex, DocEntry);
                            oDataTable.SetValue("ItemCode", rowIndex, ItemsDT.Rows[row]["ItemCode"].ToString());
                            oDataTable.SetValue("Dscptn", rowIndex, ItemsDT.Rows[row]["Dscptn"].ToString());
                            //oDataTable.SetValue("Qnty", rowIndex, ItemsDT.Rows[row]["Qnty"].ToString());
                            oDataTable.SetValue("GrsAmnt", rowIndex, ItemsDT.Rows[row]["GrsAmnt"]);
                            oDataTable.SetValue("VatGrp", rowIndex, ItemsDT.Rows[row]["VatGrp"].ToString());
                            oDataTable.SetValue("VatAmnt", rowIndex, ItemsDT.Rows[row]["VatAmnt"]);

                            rowIndex++;
                        }
                    }

                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("ItemsMTR").Specific;
                    oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single;
                    SAPbouiCOM.Columns oColumns = oMatrix.Columns;
                    SAPbouiCOM.Column oColumn;
                    string UID = "ItemsMTR";

                    foreach (SAPbouiCOM.DataColumn column in oDataTable.Columns)
                    {
                        string columnName = column.Name;
                        if (columnName == "VatGrp")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("VatGroup");
                            oColumn.DataBind.Bind(UID, columnName);

                            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            string query = "select \"Code\" " +
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
                            oColumn.DataBind.Bind(UID, columnName);
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
                        }
                        else if (columnName == "LineID")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = "#";
                            oColumn.Editable = false;
                            oColumn.DataBind.Bind(UID, columnName);
                        }
                        else if (columnName == "GrsAmnt")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("GrossAmount");
                            oColumn.DataBind.Bind(UID, columnName);
                            oColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                        }
                        else if (columnName == "VatAmnt")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("VatAmount");
                            oColumn.DataBind.Bind(UID, columnName);
                            oColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                        }
                        else if (columnName == "Dscptn")
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Description");
                            oColumn.DataBind.Bind(UID, columnName);
                            oColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                        }
                        else
                        {
                            oColumn = oColumns.Add(columnName, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oColumn.TitleObject.Caption = BDOSResources.getTranslate(columnName);
                            oColumn.DataBind.Bind(UID, columnName);
                            oColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                        }
                    }

                    oMatrix.Clear();
                    oMatrix.LoadFromDataSource();

                    top += height + 5;

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
    }
}
