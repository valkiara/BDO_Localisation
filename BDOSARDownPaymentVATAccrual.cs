using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace BDO_Localisation_AddOn
{
    static partial class BDOSARDownPaymentVATAccrual
    {
        public static bool openFormEvent = false;

        public static void createDocumentUDO( out string errorText)
        {
            errorText = null;
            string tableName = "BDOSARDV";
            string description = "A/R Down Payment VAT Accrual";

            SAPbobsCOM.UserObjectsMD oUserObjectsMD = null;
            oUserObjectsMD = (SAPbobsCOM.UserObjectsMD)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectsMD);

            int result = UDO.addUserTable( tableName, description, SAPbobsCOM.BoUTBTableType.bott_Document, out errorText);

            if (result != 0)
            {
                return;
            }

            Dictionary<string, object> fieldskeysMap;
            
            fieldskeysMap = new Dictionary<string, object>(); //ბიზნესპარტნიორი (კოდი)
            fieldskeysMap.Add("Name", "cardCode");
            fieldskeysMap.Add("TableName", "BDOSARDV");
            fieldskeysMap.Add("Description", "Customer Code");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 15);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //ბიზნესპარტნიორი (სახელი)
            fieldskeysMap.Add("Name", "cardCodeN");
            fieldskeysMap.Add("TableName", "BDOSARDV");
            fieldskeysMap.Add("Description", "Customer Name");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 100);

            UDO.addUserTableFields( fieldskeysMap, out errorText);
            
            fieldskeysMap = new Dictionary<string, object>(); //კომენტარი
            fieldskeysMap.Add("Name", "remark");
            fieldskeysMap.Add("TableName", "BDOSARDV");
            fieldskeysMap.Add("Description", "Remark");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 254);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //კომენტარი
            fieldskeysMap.Add("Name", "baseDoc");
            fieldskeysMap.Add("TableName", "BDOSARDV");
            fieldskeysMap.Add("Description", "Base document");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 20);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); //კომენტარი
            fieldskeysMap.Add("Name", "DocDate");
            fieldskeysMap.Add("TableName", "BDOSARDV");
            fieldskeysMap.Add("Description", "DocDate");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Date);
            fieldskeysMap.Add("DefaultValue", "203");
            
            UDO.addUserTableFields( fieldskeysMap, out errorText);

            
            fieldskeysMap = new Dictionary<string, object>(); //კომენტარი
            fieldskeysMap.Add("Name", "baseDocT");
            fieldskeysMap.Add("TableName", "BDOSARDV");
            fieldskeysMap.Add("Description", "Base document type");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 20);

            UDO.addUserTableFields( fieldskeysMap, out errorText);
            
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "GrsAmnt"); //საფუძველი დოკუმენტის დღგ-ის თანხა
            fieldskeysMap.Add("TableName", "BDOSARDV");
            fieldskeysMap.Add("Description", "Gross Amount");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "VatAmount"); //საფუძველი დოკუმენტის დღგ-ის თანხა
            fieldskeysMap.Add("TableName", "BDOSARDV");
            fieldskeysMap.Add("Description", "VAT Amount");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields( fieldskeysMap, out errorText);


            tableName = "BDOSRDV1";
            description = "A/R DP VAT Accr. Child1";

            result = UDO.addUserTable( tableName, description, SAPbobsCOM.BoUTBTableType.bott_DocumentLines, out errorText);

            if (result != 0)
            {
                return;
            }

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "ItemCode"); //ზედნადების ნომერი
            fieldskeysMap.Add("TableName", "BDOSRDV1");
            fieldskeysMap.Add("Description", "Item code");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 20);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "Dscptn"); //ზედნადების ნომერი
            fieldskeysMap.Add("TableName", "BDOSRDV1");
            fieldskeysMap.Add("Description", "Item description");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 100);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "VatGrp"); //ზედნადების ნომერი
            fieldskeysMap.Add("TableName", "BDOSRDV1");
            fieldskeysMap.Add("Description", "Vat Group");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 20);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "Qnty"); 
            fieldskeysMap.Add("TableName", "BDOSRDV1");
            fieldskeysMap.Add("Description", "Quantity");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Quantity);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "GrsAmnt"); //საფუძველი დოკუმენტის დღგ-ის თანხა
            fieldskeysMap.Add("TableName", "BDOSRDV1");
            fieldskeysMap.Add("Description", "Gross Amount");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "VatAmount"); //საფუძველი დოკუმენტის დღგ-ის თანხა
            fieldskeysMap.Add("TableName", "BDOSRDV1");
            fieldskeysMap.Add("Description", "Vat Amount");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields( fieldskeysMap, out errorText);            

            GC.Collect();
        }

        public static void registerUDO( out string errorText)
        {
            errorText = null;
            string code = "UDO_F_BDO_ARDPV_D"; //20 characters (must include at least one alphabetical character).
            Dictionary<string, object> formProperties;

            formProperties = new Dictionary<string, object>();
            formProperties.Add("Name", "A/R Down Payment VAT Accrual"); //100 characters
            formProperties.Add("TableName", "BDOSARDV");
            formProperties.Add("ObjectType", SAPbobsCOM.BoUDOObjType.boud_Document);
            formProperties.Add("CanCancel", SAPbobsCOM.BoYesNoEnum.tYES);
            formProperties.Add("CanClose", SAPbobsCOM.BoYesNoEnum.tYES);
            formProperties.Add("CanCreateDefaultForm", SAPbobsCOM.BoYesNoEnum.tYES);
            formProperties.Add("CanDelete", SAPbobsCOM.BoYesNoEnum.tNO);
            formProperties.Add("CanFind", SAPbobsCOM.BoYesNoEnum.tYES);
            formProperties.Add("CanYearTransfer", SAPbobsCOM.BoYesNoEnum.tYES);
            formProperties.Add("ManageSeries", SAPbobsCOM.BoYesNoEnum.tYES);

            List<Dictionary<string, object>> listFindColumns = new List<Dictionary<string, object>>();
            List<Dictionary<string, object>> listFormColumns = new List<Dictionary<string, object>>();
            List<Dictionary<string, object>> listChildTables = new List<Dictionary<string, object>>();

            Dictionary<string, object> fieldskeysMap;

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "U_cardCode");
            fieldskeysMap.Add("ColumnDescription", "Customer Code"); //30 characters
            listFindColumns.Add(fieldskeysMap);
            
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "U_cardCodeN");
            fieldskeysMap.Add("ColumnDescription", "Customer Name"); //30 characters
            listFindColumns.Add(fieldskeysMap);
            
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("ColumnAlias", "Remark");
            fieldskeysMap.Add("ColumnDescription", "Remark"); //30 characters
            listFindColumns.Add(fieldskeysMap);

            formProperties.Add("FindColumns", listFindColumns);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("FormColumnAlias", "DocEntry");
            fieldskeysMap.Add("FormColumnDescription", "DocEntry"); //30 characters
            listFormColumns.Add(fieldskeysMap);

            formProperties.Add("FormColumns", listFormColumns);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("TableName", "BDOSRDV1");
            fieldskeysMap.Add("ObjectName", "BDOSRDV1"); //30 characters
            listChildTables.Add(fieldskeysMap);

            formProperties.Add("ChildTables", listChildTables);

            UDO.registerUDO( code, formProperties, out errorText);

            GC.Collect();
        }

        public static void addMenus( out string errorText)
        {
            errorText = null;

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
                oCreationPackage.UniqueID = "UDO_F_BDO_ARDPV_D";
                oCreationPackage.String = BDOSResources.getTranslate("ARDownPaymentVATAccrual");
                oCreationPackage.Position = fatherMenuItem.SubMenus.Count - 1;

                menuItem = fatherMenuItem.SubMenus.AddEx(oCreationPackage);
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
        }

        public static void uiApp_FormDataEvent(  ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;
            
            SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(BusinessObjectInfo.FormTypeEx, Program.currentFormCount);

            if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD & BusinessObjectInfo.BeforeAction == false)
            {
                if (oForm.DataSources.DBDataSources.Item("@BDOSARDV").GetValue("U_docDate", 0) == "")
                {
                    Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("PostingDate") + " " + BDOSResources.getTranslate("YouCantLeaveEmpty"));
                    Program.uiApp.MessageBox(BDOSResources.getTranslate("OperationUnsuccesfullSeeLog"));
                    BubbleEvent = false;
                }
                
                formDataLoad( oForm, out errorText);
                setVisibleFormItems(oForm, out errorText);
            }
           
            if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE & BusinessObjectInfo.BeforeAction == false & BusinessObjectInfo.ActionSuccess == true)
            {
                if (Program.cancellationTrans == true & Program.canceledDocEntry != 0)
                {
                    cancellation( oForm, Program.canceledDocEntry, out errorText);
                    Program.canceledDocEntry = 0;
                }
            }

            if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD)
            {                
                SAPbouiCOM.DBDataSource DocDBSourceOCRD = oForm.DataSources.DBDataSources.Item(0);
                if (BusinessObjectInfo.ActionSuccess != BusinessObjectInfo.BeforeAction && BubbleEvent)
                {
                    //დოკუმენტის გატარების დროს გატარდეს ბუღლტრული გატარება
                    SAPbouiCOM.DBDataSource DocDBSourcePAYR = oForm.DataSources.DBDataSources.Item(0);

                    if (DocDBSourcePAYR.GetValue("CANCELED", 0) == "N")
                    {
                        string DocEntry = DocDBSourcePAYR.GetValue("DocEntry", 0);
                        string DocCurrency = "";
                        decimal DocRate = 0;
                        string DocNum = DocDBSourcePAYR.GetValue("DocNum", 0);
                        DateTime DocDate = DateTime.ParseExact(DocDBSourcePAYR.GetValue("U_DocDate", 0), "yyyyMMdd", CultureInfo.InvariantCulture);

                        CommonFunctions.StartTransaction();

                        Program.JrnLinesGlobal = new DataTable();
                        
                        DataTable JrnLinesDT = createAdditionalEntries( oForm, null, null, DocCurrency, DocRate);

                        JrnEntry( DocEntry, DocNum, DocDate, JrnLinesDT,  out errorText);
                        if (errorText != null)
                        {
                            Program.uiApp.MessageBox(errorText);
                            BubbleEvent = false;
                        }
                        else
                        {
                            if (BusinessObjectInfo.ActionSuccess == false)
                            {
                                Program.JrnLinesGlobal = JrnLinesDT;
                            }
                        }

                        if (Program.oCompany.InTransaction)
                        {
                            //თუ დოკუმენტი გატარდა, მერე ვაკეთებს ბუღალტრულ გატარებას
                            if (BusinessObjectInfo.ActionSuccess == true & BusinessObjectInfo.BeforeAction == false)
                            {
                                CommonFunctions.EndTransaction( SAPbobsCOM.BoWfTransOpt.wf_Commit);
                            }
                            else
                            {
                                CommonFunctions.EndTransaction( SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                            }
                        }
                        else
                        {
                            Program.uiApp.MessageBox("ტრანზაქციის დასრულებს შეცდომა");
                            BubbleEvent = false;
                        }
                    }
                }
            }

        }

        public static void uiApp_MenuEvent(  ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent, out string errorText)
        {
            errorText = null;
            BubbleEvent = true;

            //----------------------------->Preview <-----------------------------
            try
            {
                if (pVal.BeforeAction && pVal.MenuUID == "PreviewUDOJrE")
                {
                   

                    SAPbouiCOM.Form oDocForm = Program.uiApp.Forms.ActiveForm;

                    if (oDocForm.Items.Item("DocDate").Specific.Value == "")
                    {
                        Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("PostingDate") + " " + BDOSResources.getTranslate("YouCantLeaveEmpty"));
                        Program.uiApp.MessageBox(BDOSResources.getTranslate("OperationUnsuccesfullSeeLog"));
                        BubbleEvent = false;
                    }

                    else
                    {
                        JournalEntryTransaction(  oDocForm, false, true, out BubbleEvent, out errorText);

                        if (BubbleEvent)
                        {
                            SAPbouiCOM.Form oJournalForm = Program.uiApp.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_JournalPosting, "", "");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Program.uiApp.MessageBox(ex.ToString(), 1, "", "");
            }
        }

        public static void uiApp_ItemEvent(  string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD & pVal.BeforeAction == true)
                {
                    createFormItems( oForm, out errorText);
                    Program.FORM_LOAD_FOR_VISIBLE = true;
                    Program.FORM_LOAD_FOR_ACTIVATE = true;
                    formDataLoad( oForm, out errorText);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE & pVal.BeforeAction == false)
                {
                    if (Program.FORM_LOAD_FOR_VISIBLE == true)
                    {
                        setSizeForm( oForm, out errorText);
                        oForm.Title = BDOSResources.getTranslate("ARDownPaymentVAT");
                        Program.FORM_LOAD_FOR_VISIBLE = false;
                    }
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                {
                    SAPbouiCOM.IChooseFromListEvent oCFLEvento = null;
                    oCFLEvento = ((SAPbouiCOM.IChooseFromListEvent)(pVal));

                    chooseFromList( oForm, oCFLEvento, pVal.ItemUID, pVal.BeforeAction, out errorText);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_RESIZE & pVal.BeforeAction == false)
                {
                    oForm.Freeze(true);
                    reArrangeFormItems(oForm);
                    oForm.Freeze(false);
                }
                if (pVal.ItemUID == "addMTRB" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK & pVal.BeforeAction == false)
                {
                    addMatrixRow( oForm, out errorText);
                }

                if (pVal.ItemUID == "delMTRB" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK & pVal.BeforeAction == false)
                {
                    delMatrixRow( oForm, out errorText);
                }

                if (pVal.ItemUID == "ItemsMTR" & pVal.ItemChanged & pVal.BeforeAction == false)
                {
                    fillTotalAmounts( oForm, out errorText);
                }

                if (pVal.ItemUID == "BDO_TaxTxt" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK & pVal.BeforeAction == false)
                {
                    string taxDoc = oForm.DataSources.UserDataSources.Item("BDO_TaxDoc").ValueEx;
                    int newDocEntry = 0;

                    SAPbouiCOM.DBDataSource DocDBSourcePAYR = oForm.DataSources.DBDataSources.Item(0);
                    int docEntry = Convert.ToInt32(DocDBSourcePAYR.GetValue("DocEntry", 0));
                    string cancelled = DocDBSourcePAYR.GetValue("CANCELED", 0).Trim();
                    string objectType = "UDO_F_BDO_ARDPV_D";
                    

                    if (docEntry != 0 & (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_VIEW_MODE))
                    {
                        if (taxDoc == "" && cancelled == "N")
                        {
                            BDO_TaxInvoiceSent.createDocument( objectType, docEntry, "", true, 0, null, false, null, out newDocEntry, out errorText);

                            if (string.IsNullOrEmpty(errorText) && newDocEntry != 0)
                            {
                                oForm.DataSources.UserDataSources.Item("BDO_TaxDoc").ValueEx = newDocEntry.ToString();
                                formDataLoad( oForm, out errorText);
                                return;
                            }
                        }
                        else if (cancelled != "N")
                        {
                            errorText = BDOSResources.getTranslate("DocumentMustNotBeCancelledOrCancellation");
                        }
                    }
                    else
                    {
                        errorText = BDOSResources.getTranslate("ToCreateTaxInvoiceWriteDocument");
                    }
                }

                if (pVal.ItemUID == "ItemsMTR" && (pVal.ColUID == "U_GrsAmnt" || pVal.ColUID == "U_VatGrp") & pVal.ItemChanged & pVal.BeforeAction == false)
                {
                    oForm.Freeze(true);

                    SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("ItemsMTR").Specific;
                    string VAtGroup = oMatrix.Columns.Item("U_VatGrp").Cells.Item(pVal.Row).Specific.Value;
                    decimal GrossAmnt = (FormsB1.cleanStringOfNonDigits( oMatrix.Columns.Item("U_GrsAmnt").Cells.Item(pVal.Row).Specific.Value));
                    decimal VatAmount = 0;
                    int row = pVal.Row;
                    decimal VatRate = CommonFunctions.GetVatGroupRate( VAtGroup, "");

                    VatAmount = CommonFunctions.roundAmountByGeneralSettings( GrossAmnt * VatRate / (100 + VatRate), "Sum");
                    oMatrix.Columns.Item("U_VatAmnt").Cells.Item(row).Specific.String = FormsB1.ConvertDecimalToStringForEditboxStrings( VatAmount);

                    oForm.Freeze(false);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_GOT_FOCUS && oForm.Visible == true && oForm.VisibleEx == true && openFormEvent == false)
                {
                    oForm.Freeze(true);
                    formDataLoad( oForm, out errorText);
                    oForm.Freeze(false);
                    openFormEvent = true;

                    setVisibleFormItems(oForm, out errorText);

                }
            }
        }

        public static void cancellation(  SAPbouiCOM.Form oForm, int docEntry, out string errorText)
        {
            errorText = null;

            try
            {
                JournalEntry.cancellation( oForm, docEntry, "UDO_F_BDO_ARDPV_D", out errorText);
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

        public static void setVisibleFormItems(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            
            try
            {   
                oForm.Items.Item("StatusC").Enabled = false;
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


        public static void getAmount( int docEntry, out double gTotal, out double lineVat, out string errorText)
        {
            errorText = null;
            gTotal = 0;
            lineVat = 0;

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string query = @"SELECT 
            ""@BDOSARDV"".""DocEntry"" AS ""docEntry"", 
            SUM(""@BDOSARDV"".""U_GrsAmnt"") AS ""GTotal"", 
            SUM(""@BDOSARDV"".""U_VatAmount"") AS ""LineVat"" 
            FROM ""@BDOSARDV"" AS ""@BDOSARDV"" 
            WHERE ""@BDOSARDV"".""DocEntry"" = '" + docEntry + @"' 
            GROUP BY ""@BDOSARDV"".""DocEntry""";

            try
            {
                oRecordSet.DoQuery(query);

                while (!oRecordSet.EoF)
                {
                    gTotal = oRecordSet.Fields.Item("GTotal").Value;
                    lineVat = oRecordSet.Fields.Item("LineVat").Value;

                    oRecordSet.MoveNext();
                    break;
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
           
        }



        private static void fillTotalAmounts(  SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = "";

            SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("ItemsMTR").Specific;
            decimal U_GrsAmnt = 0;
            decimal U_VatAmount = 0;
            for (int i = 1; i <= oMatrix.RowCount; i++)
            {
                U_GrsAmnt = U_GrsAmnt + Convert.ToDecimal(oMatrix.GetCellSpecific("U_GrsAmnt", i).Value, CultureInfo.InvariantCulture);
                U_VatAmount = U_VatAmount + Convert.ToDecimal(oMatrix.GetCellSpecific("U_VatAmnt", i).Value, CultureInfo.InvariantCulture);
            }
            string U_GrsAmnts = U_GrsAmnt.ToString(CultureInfo.InvariantCulture);
            string U_VatAmounts = U_VatAmount.ToString(CultureInfo.InvariantCulture);

           /* oForm.Items.Item("U_GrsAmnt").Specific.Value = U_GrsAmnts;
            oForm.Items.Item("U_GrsAmnt").Specific.Value = U_VatAmounts;*/


            oForm.DataSources.DBDataSources.Item("@BDOSARDV").SetValue("U_GrsAmnt", 0, U_GrsAmnts);
            oForm.DataSources.DBDataSources.Item("@BDOSARDV").SetValue("U_VatAmount", 0, U_VatAmounts);
            
            
        }

        public static bool checkDocumentForTaxInvoice( int docEntry, DateTime docDate, DateTime docDateForMonth, out bool primary, out DataTable confirmedInvoices, out string errorText)
        {
            errorText = null;
            primary = false;
            confirmedInvoices = null;
            DataTable nonConfirmedInvoices = null;
            DateTime firstDay = new DateTime(docDateForMonth.Year, docDateForMonth.Month, 1);
            DateTime lastDay = firstDay.AddMonths(1).AddDays(-1);

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string query = @"SELECT
	             ""@BDOSARDV"".""U_DocDate"",
	             ""@BDOSARDV"".""DocEntry"",
	             ""@BDOSARDV"".""DocNum"",
	             
	             
                 ""BDO_TAXS"".""DocEntry"" AS ""invDocEntry"",
	             ""BDO_TAXS"".""DocNum"" AS ""invDocNum"",
	             ""BDO_TAXS"".""U_status"",
	             ""BDO_TAXS"".""U_invID"",
	             ""BDO_TAXS"".""U_number"",
	             ""BDO_TAXS"".""U_series"" 
            FROM ""@BDOSARDV"" AS ""@BDOSARDV"" 
            
            INNER JOIN (SELECT
            	 ""BDO_TXS1"".""U_baseDoc"" AS ""U_baseDoc"",
            	 ""BDO_TAXS"".""DocEntry"",
	             ""BDO_TAXS"".""DocNum"",            	 
            	 ""BDO_TAXS"".""U_status"" AS ""U_status"",
            	 ""BDO_TAXS"".""U_invID"" AS ""U_invID"",
            	 ""BDO_TAXS"".""U_number"" AS ""U_number"",
            	 ""BDO_TAXS"".""U_series"" AS ""U_series"" 
            	FROM ""@BDO_TXS1"" AS ""BDO_TXS1"" 
            	INNER JOIN ""@BDO_TAXS"" AS ""BDO_TAXS"" ON ""BDO_TXS1"".""DocEntry"" = ""BDO_TAXS"".""DocEntry"" 
            	WHERE ""BDO_TAXS"".""U_downPaymnt"" = 'Y' 
            	AND (""BDO_TAXS"".""Canceled"" = 'N' AND ""BDO_TAXS"".""U_status"" NOT IN ('removed',
            	 'canceled'))
            	AND ""BDO_TXS1"".""U_baseDocT"" = 'ARDownPaymentVAT' 
            	GROUP BY ""BDO_TXS1"".""U_baseDoc"",
          	     ""BDO_TAXS"".""DocEntry"",
	             ""BDO_TAXS"".""DocNum"",
            	 ""BDO_TAXS"".""U_status"",
            	 ""BDO_TAXS"".""U_invID"",
            	 ""BDO_TAXS"".""U_number"",
            	 ""BDO_TAXS"".""U_series"" ) AS ""BDO_TAXS"" ON ""@BDOSARDV"".""DocEntry"" = ""BDO_TAXS"".""U_baseDoc"" 
            WHERE 
                
                
            ""@BDOSARDV"".""U_DocDate"" <= '" + docDate.ToString("yyyyMMdd") + "' " +
            @"AND ""@BDOSARDV"".""U_DocDate"" >= '" + firstDay.ToString("yyyyMMdd") + @"' AND ""@BDOSARDV"".""U_DocDate"" <= '" + lastDay.ToString("yyyyMMdd") + "' " +
            @"AND ""@BDOSARDV"".""DocEntry"" < '" + docEntry + "' " +
            @"GROUP BY ""@BDOSARDV"".""U_DocDate"",
            	 ""@BDOSARDV"".""DocEntry"",
            	 ""@BDOSARDV"".""DocNum"",
            	 
                 ""BDO_TAXS"".""DocEntry"",
	             ""BDO_TAXS"".""DocNum"",
            	 ""BDO_TAXS"".""U_status"",
            	 ""BDO_TAXS"".""U_invID"",
            	 ""BDO_TAXS"".""U_number"",
            	 ""BDO_TAXS"".""U_series""
            ORDER BY ""@BDOSARDV"".""U_DocDate"" DESC,
             ""@BDOSARDV"".""DocEntry"" DESC";

            try
            {
                oRecordSet.DoQuery(query);
                int recordCount = oRecordSet.RecordCount;

                if (recordCount == 0)
                {
                    primary = true;
                    return true;
                }
                else
                {
                    string invStatus;

                    confirmedInvoices = new DataTable();
                    confirmedInvoices.Columns.Add("DocEntry", typeof(int));
                    confirmedInvoices.Columns.Add("DocNum", typeof(int));
                    confirmedInvoices.Columns.Add("BaseEntry", typeof(int));
                    confirmedInvoices.Columns.Add("U_invID", typeof(string));
                    confirmedInvoices.Columns.Add("U_number", typeof(string));
                    confirmedInvoices.Columns.Add("U_series", typeof(string));
                    confirmedInvoices.Columns.Add("InvDocEntry", typeof(int));
                    confirmedInvoices.Columns.Add("InvDocNum", typeof(int));

                    nonConfirmedInvoices = new DataTable();
                    nonConfirmedInvoices.Columns.Add("DocEntry", typeof(int));
                    nonConfirmedInvoices.Columns.Add("DocNum", typeof(int));
                    nonConfirmedInvoices.Columns.Add("BaseEntry", typeof(int));
                    nonConfirmedInvoices.Columns.Add("U_invID", typeof(string));
                    nonConfirmedInvoices.Columns.Add("U_number", typeof(string));
                    nonConfirmedInvoices.Columns.Add("U_series", typeof(string));
                    nonConfirmedInvoices.Columns.Add("InvDocEntry", typeof(int));
                    nonConfirmedInvoices.Columns.Add("InvDocNum", typeof(int));

                    while (!oRecordSet.EoF)
                    {
                        invStatus = oRecordSet.Fields.Item("U_status").Value.ToString();
                        DataRow taxDataRow;
                        if (invStatus == "confirmed" || invStatus == "correctionConfirmed" || invStatus == "primary" || invStatus == "corrected")
                            taxDataRow = confirmedInvoices.Rows.Add();
                        else
                            taxDataRow = nonConfirmedInvoices.Rows.Add();

                        taxDataRow["DocEntry"] = Convert.ToInt32(oRecordSet.Fields.Item("DocEntry").Value);
                        taxDataRow["DocNum"] = Convert.ToInt32(oRecordSet.Fields.Item("DocNum").Value);
                        taxDataRow["BaseEntry"] = Convert.ToInt32(oRecordSet.Fields.Item("DocEntry").Value);
                        taxDataRow["U_invID"] = oRecordSet.Fields.Item("U_invID").Value.ToString();
                        taxDataRow["U_number"] = oRecordSet.Fields.Item("U_number").Value.ToString();
                        taxDataRow["U_series"] = oRecordSet.Fields.Item("U_series").Value.ToString();
                        taxDataRow["InvDocEntry"] = Convert.ToInt32(oRecordSet.Fields.Item("InvDocEntry").Value);
                        taxDataRow["InvDocNum"] = Convert.ToInt32(oRecordSet.Fields.Item("InvDocNum").Value);

                        oRecordSet.MoveNext();
                    }

                    if (confirmedInvoices.Rows.Count > 0)
                    {
                        primary = false;
                    }
                    if (nonConfirmedInvoices.Rows.Count > 0)
                    {
                        List<int> oList = nonConfirmedInvoices.AsEnumerable().Select(r => r.Field<int>("InvDocNum")).ToList();
                        errorText = BDOSResources.getTranslate("OnARDownPaymentRequestThereIsAnotherARDownPaymentInvoiceWithTaxInvoiceSentTheStatusOfWhichShouldBeFromThisList") + " : " + "\"" + BDOSResources.getTranslate("deleted") + "\", \"" + BDOSResources.getTranslate("Canceled") + "\", \"" + BDOSResources.getTranslate("Denied") + "\", \"" + BDOSResources.getTranslate("Confirmed") + "\", \"" + BDOSResources.getTranslate("CorrectionConfirmed") + "\"! ";
                        if (oList.Count > 1)
                            errorText = errorText + '\n' + "\"" + BDOSResources.getTranslate("TaxInvoiceSent") + "\" " + BDOSResources.getTranslate("DocumentsSNumbersAre") + " : " + string.Join(",", oList);
                        else
                            errorText = errorText + '\n' + "\"" + BDOSResources.getTranslate("TaxInvoiceSent") + "\" " + BDOSResources.getTranslate("DocumentSNumberIs") + " : " + string.Join(",", oList);

                        return false;
                    }
                    return true;
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
                return false;
            }
            finally
            {
                Marshal.ReleaseComObject(oRecordSet);
                oRecordSet = null;
            }
        }


        private static void addMatrixRow( SAPbouiCOM.Form oForm, out string errorText)
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
                    index = Convert.ToInt32(oMatrix.Columns.Item("LineID").Cells.Item(oMatrix.RowCount).Specific.Value) + 1;
                }

                
                oMatrix.AddRow(1, -1);
                oMatrix.AutoResizeColumns();
                oMatrix.Columns.Item("LineID").Cells.Item(oMatrix.RowCount).Specific.Value = index;
                              
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

        public static void delMatrixRow( SAPbouiCOM.Form oForm, out string errorText)
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
                        oForm.DataSources.DBDataSources.Item("@BDOSRDV1").RemoveRecord(row - deletedRowCount);
                        firstRow = row;
                    }
                }

                oMatrix.LoadFromDataSource();
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


        public static void chooseFromList( SAPbouiCOM.Form oForm, SAPbouiCOM.IChooseFromListEvent oCFLEvento, string itemUID, bool beforeAction, out string errorText)
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

                        if (sCFL_ID == "BusinessPartner_CFL")
                        {
                            string businessPartnerCode = Convert.ToString(oDataTable.GetValue("CardCode", 0));
                            string businessPartnerName = Convert.ToString(oDataTable.GetValue("CardName", 0));
                            
                            oForm.DataSources.DBDataSources.Item("@BDOSARDV").SetValue("U_cardCode", 0, businessPartnerCode);
                            oForm.DataSources.DBDataSources.Item("@BDOSARDV").SetValue("U_cardCodeN", 0, businessPartnerName);

                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            {
                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                            }
                        }

                        if (sCFL_ID == "BaseDoc_CFL")
                        {
                            string docEntry = Convert.ToString(oDataTable.GetValue("DocEntry", 0));

                            oForm.DataSources.DBDataSources.Item("@BDOSARDV").SetValue("U_baseDoc", 0, docEntry);


                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            {
                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                            }
                        }

                        if (sCFL_ID == "Item_CFL")
                        {
                            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("ItemsMTR").Specific));
                            string ItemCode = Convert.ToString(oDataTable.GetValue("ItemCode", 0));
                            string ItemName = Convert.ToString(oDataTable.GetValue("ItemName", 0));

                           
                            oMatrix.SetCellWithoutValidation(oCFLEvento.Row, "U_ItemCode", ItemCode);
                            oMatrix.SetCellWithoutValidation(oCFLEvento.Row, "U_dscptn", ItemName);
                           
                            
                        }
                    }
                    


                }
            }
            catch (Exception ex)
            {
               string exsd = ex.Message;
            }
        }

        public static void reArrangeFormItems(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.Item oItem = null;
            int height = 15;
            int top = 6;
            top = top + height + 1;

            
            //ღილაკები
            oItem = oForm.Items.Item("1");
            oItem.Top = oForm.ClientHeight - 25;

            oItem = oForm.Items.Item("2");
            oItem.Top = oForm.ClientHeight - 25;

            oForm.Items.Item("ItemsMTR").Height = oForm.Height - 240;
           
            oForm.Items.Item("CreatorS").Top =  oForm.Items.Item("ItemsMTR").Top + oForm.Items.Item("ItemsMTR").Height + 5;
            oForm.Items.Item("CreatorE").Top = oForm.Items.Item("ItemsMTR").Top + oForm.Items.Item("ItemsMTR").Height + 5;
            
            oForm.Items.Item("RemarksS").Top = oForm.Items.Item("CreatorS").Top + oForm.Items.Item("CreatorS").Height + 5;
            oForm.Items.Item("RemarksE").Top = oForm.Items.Item("CreatorS").Top + oForm.Items.Item("CreatorS").Height + 5;

            oForm.Items.Item("AmountS").Top = oForm.Items.Item("CreatorS").Top;
            oForm.Items.Item("AmountE").Top = oForm.Items.Item("CreatorS").Top;
            oForm.Items.Item("AmountE").Enabled = false;
            oForm.Items.Item("VatAmountS").Top = oForm.Items.Item("RemarksS").Top;
            oForm.Items.Item("VatAmount").Top = oForm.Items.Item("RemarksS").Top;
            oForm.Items.Item("VatAmount").Enabled = false;
       
        }

        public static void createFormItems(  SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            Dictionary<string, object> formItems;
            string itemName = "";

            int left_s = 6;
            int left_e = 127;
            int height = 15;
            int top = 6;
            int width_s = 121;
            int width_e = 148;
            int fontSize = 10;


            left_s = 295;
            left_e = left_s + 121;

            oForm.Title = BDOSResources.getTranslate("ARDownPaymentVATAccrual");
            
            formItems = new Dictionary<string, object>();
            itemName = "DocDateS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("PostingDate"));
            formItems.Add("LinkTo", "DocDate");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "DocDate";
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSARDV");
            formItems.Add("Alias", "U_DocDate");
            formItems.Add("Bound", true);
            formItems.Add("Size", 20);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Enabled", true);
            


            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            left_s = 6;
            left_e = 127;

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "cardCodeS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("CustomerCode"));
            formItems.Add("LinkTo", "cardCodeE");
            formItems.Add("FontSize", fontSize);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            bool multiSelection = false;
            
            string objectTypeARDP = "203";
            string uniqueID_lf_BaseDocCFL = "BaseDoc_CFL";
            FormsB1.addChooseFromList( oForm, multiSelection, objectTypeARDP, uniqueID_lf_BaseDocCFL);

            string objectTypeCardCode = "2";
            string uniqueID_lf_BusinessPartnerCFL = "BusinessPartner_CFL";
            FormsB1.addChooseFromList( oForm, multiSelection, objectTypeCardCode, uniqueID_lf_BusinessPartnerCFL);


            string objectTypeItem = "4"; //SAPbouiCOM.BoLinkedObject.lf_Invoice
            string uniqueID_lf_ItemCFL = "Item_CFL";
            FormsB1.addChooseFromList( oForm, multiSelection, objectTypeItem, uniqueID_lf_ItemCFL);

            //პირობის დადება ბიზნესპარტნიორის არჩევის სიაზე
            SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(uniqueID_lf_BusinessPartnerCFL);
            SAPbouiCOM.Conditions oCons = oCFL.GetConditions();
            SAPbouiCOM.Condition oCon = oCons.Add();
            oCon.Alias = "CardType";
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCon.CondVal = "C"; //მყიდველი
            oCFL.SetConditions(oCons);

            formItems = new Dictionary<string, object>();
            itemName = "cardCodeE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSARDV");
            formItems.Add("Alias", "U_cardCode");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e / 4);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("ChooseFromListUID", uniqueID_lf_BusinessPartnerCFL);
            formItems.Add("ChooseFromListAlias", "CardCode");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "cardCodeNE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSARDV");
            formItems.Add("Alias", "U_cardCodeN");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e + width_e / 4);
            formItems.Add("Width", width_e * 3 / 4);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "cardCodeLB"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            formItems.Add("Left", left_e - 20);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("LinkTo", "cardCodeE");
            formItems.Add("LinkedObjectType", objectTypeCardCode);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            //top = top + height + 1;
            left_s = 295;
            left_e = left_s + 121;

            Dictionary<string, string> listValidValues = new Dictionary<string, string>(); //კორექტირების მიზეზები
            
            listValidValues.Add("203","DownPayment"); 
            
            formItems = new Dictionary<string, object>();
            itemName = "baseDocS"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSARDV");
            formItems.Add("Alias", "U_baseDocT");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s-30);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("baseDoc"));
            formItems.Add("LinkTo", "baseDocE");
            formItems.Add("FontSize", fontSize);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", true);
            formItems.Add("Visible", true);
            formItems.Add("ValidValues", listValidValues);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }
            

            formItems = new Dictionary<string, object>();
            itemName = "baseDocE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSARDV");
            formItems.Add("Alias", "U_baseDoc");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("ChooseFromListUID", uniqueID_lf_BaseDocCFL);
            formItems.Add("ChooseFromListAlias", "DocEntry");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

           
            formItems = new Dictionary<string, object>();
            itemName = "baseDocLB"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            formItems.Add("Left", left_e - 20);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("LinkTo", "baseDocE");
            formItems.Add("LinkedObjectType", objectTypeARDP);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 1;


            formItems = new Dictionary<string, object>();
            itemName = "StatusS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Canceled"));
            formItems.Add("LinkTo", "StatusC");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "StatusC"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSARDV");
            formItems.Add("Alias", "Canceled");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 1;

            //-------------------------------------------ანგარიშ-ფაქტურა----------------------------------->

            formItems = new Dictionary<string, object>();
            itemName = "BDO_TaxTxt"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_e * 1.5-10);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("CreateTaxInvoice"));
            formItems.Add("TextStyle", 4);
            formItems.Add("FontSize", 10);
            formItems.Add("Enabled", true);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            multiSelection = false;
            string objectType = "UDO_F_BDO_TAXS_D"; //Tax invoice sent document
            string uniqueID_TaxInvoiceSentCFL = "TaxInvoiceSent_CFL";
            FormsB1.addChooseFromList( oForm, multiSelection, objectType, uniqueID_TaxInvoiceSentCFL);

            formItems = new Dictionary<string, object>();
            itemName = "BDO_TaxDoc"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "UserDataSources");
            formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
            formItems.Add("Length", 11);
            formItems.Add("TableName", "");
            formItems.Add("Alias", itemName);
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e + width_e - 40);
            formItems.Add("Width", 40);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("AffectsFormMode", false);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", false);
            formItems.Add("ChooseFromListUID", uniqueID_TaxInvoiceSentCFL);
            formItems.Add("ChooseFromListAlias", "DocEntry");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "BDO_TaxLB"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            formItems.Add("Left", left_e + width_e - 40 - 20);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("LinkTo", "BDO_TaxDoc");
            formItems.Add("LinkedObjectType", objectType);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }
                        
            left_s = 6;
            left_e = 127;

            formItems = new Dictionary<string, object>();
            itemName = "BDOSJrnEnS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("TransactionNo"));
            formItems.Add("LinkTo", "BDOSJrnEnt");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "BDOSJrnEnt";
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "UserDataSources");
            formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
            formItems.Add("Length", 11);
            formItems.Add("TableName", "");
            formItems.Add("Alias", itemName);
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Enabled", false);
            formItems.Add("AffectsFormMode", false);


            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "BDOSJEntLB";
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            formItems.Add("Left", left_e - 20);
            formItems.Add("Top", top);
            formItems.Add("UID", itemName);
            formItems.Add("LinkTo", "BDOSJrnEnt");
            formItems.Add("LinkedObjectType", "30");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 10;
            
            formItems = new Dictionary<string, object>();
            itemName = "addMTRB"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            formItems.Add("Left", left_s);
            formItems.Add("Width", 100);
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
            formItems.Add("Left", left_s + 100 + 1);
            formItems.Add("Width", 100);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Delete"));

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 5;

            int wblMTRWidth = oForm.Width;

            formItems = new Dictionary<string, object>();
            itemName = "ItemsMTR"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
            formItems.Add("Left", left_s);
            formItems.Add("Width", wblMTRWidth+80);
            formItems.Add("Top", top);
            formItems.Add("Height", oForm.Height - 240);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("AffectsFormMode", false);
            formItems.Add("Enabled", true);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("ItemsMTR").Specific));
            oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
            SAPbouiCOM.Columns oColumns = oMatrix.Columns;
            SAPbouiCOM.Column oColumn;

            
                        

            oColumn = oColumns.Add("LineID", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "  ";
            oColumn.Width = 20 - 1;
            oColumn.Editable = false;
            oColumn.Visible = true;

            oColumn = oColumns.Add("U_ItemCode", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oColumn.TitleObject.Caption = "ItemCode";
            oColumn.Width = 20 - 1;
            oColumn.Editable = true;
            oColumn.Visible = true;

            oColumn = oColumns.Add("U_dscptn", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Dscptn";
            oColumn.Width = 20 - 1;
            oColumn.Editable = true;
            oColumn.Visible = true;



            oColumn = oColumns.Add("U_Qnty", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Qnty";
            oColumn.Width = 20 - 1;
            oColumn.Editable = true;
            oColumn.Visible = true;

            oColumn = oColumns.Add("U_GrsAmnt", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "GrsAmnt";
            oColumn.Width = 20 - 1;
            oColumn.Editable = true;
            oColumn.Visible = true;


            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string query = "select * " +
            "FROM  \"OVTG\" " +
            "WHERE \"Category\"='O'";

            oRecordSet.DoQuery(query);
            oColumn = oColumns.Add("U_VatGrp", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            oColumn.TitleObject.Caption = "VatGrp";
            while (!oRecordSet.EoF)
            {
                oColumn.ValidValues.Add(oRecordSet.Fields.Item("Code").Value, oRecordSet.Fields.Item("Code").Value);
                oRecordSet.MoveNext();
            }
            oColumn.Width = 20 - 1;
            oColumn.Editable = true;
            oColumn.Visible = true;

            oColumn = oColumns.Add("U_VatAmnt", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "VatAmount";
            oColumn.Width = 20 - 1;
            oColumn.Editable = true;
            oColumn.Visible = true;
            
            SAPbouiCOM.DBDataSource oDBDataSource;
            oDBDataSource = oForm.DataSources.DBDataSources.Add("@BDOSRDV1");

            oColumn = oColumns.Item("LineID");
            oColumn.DataBind.SetBound(true, "@BDOSRDV1", "LineID");

            oColumn = oColumns.Item("U_ItemCode");
            oColumn.DataBind.SetBound(true, "@BDOSRDV1", "U_ItemCode");
            oColumn.ChooseFromListUID = "Item_CFL";
            oColumn.ChooseFromListAlias = "ItemCode";
            SAPbouiCOM.LinkedButton oLink = oColumn.ExtendedObject;
            oLink.LinkedObjectType = "4";


            oColumn = oColumns.Item("U_dscptn");
            oColumn.DataBind.SetBound(true, "@BDOSRDV1", "U_Dscptn");
            
            oColumn = oColumns.Item("U_GrsAmnt");
            oColumn.DataBind.SetBound(true, "@BDOSRDV1", "U_GrsAmnt");

            oColumn = oColumns.Item("U_Qnty");
            oColumn.DataBind.SetBound(true, "@BDOSRDV1", "U_Qnty");            

            oColumn = oColumns.Item("U_VatGrp");
            oColumn.DataBind.SetBound(true, "@BDOSRDV1", "U_VatGrp");

            oColumn = oColumns.Item("U_VatAmnt");
            oColumn.DataBind.SetBound(true, "@BDOSRDV1", "U_VatAmount");


            oMatrix.Clear();
            oDBDataSource.Query();
            oMatrix.LoadFromDataSource();
            oMatrix.AutoResizeColumns();

            //სარდაფი
            left_s = 6;
            left_e = 127;
            top = top + oForm.Items.Item("ItemsMTR").Height + 40;

            formItems = new Dictionary<string, object>();
            itemName = "CreatorS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Creator"));
            formItems.Add("LinkTo", "CreatorE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "CreatorE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSARDV");
            formItems.Add("Alias", "Creator");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "RemarksS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Remarks"));
            formItems.Add("LinkTo", "RemarksE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "RemarksE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSARDV");
            formItems.Add("Alias", "Remark");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EXTEDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", 3 * height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("ScrollBars", SAPbouiCOM.BoScrollBars.sb_Vertical);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }


            left_s = 295;
            left_e = left_s + 121;
            top = top + oForm.Items.Item("ItemsMTR").Height + 40;

            formItems = new Dictionary<string, object>();
            itemName = "AmountS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("AmountWithVat"));
            formItems.Add("LinkTo", "AmountE");
            formItems.Add("FontSize", fontSize);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "AmountE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSARDV");
            formItems.Add("Alias", "U_GrsAmnt");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            top = top + height + 1;

            formItems = new Dictionary<string, object>();
            itemName = "VatAmountS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("VatAmount"));
            formItems.Add("LinkTo", "VatAmount");
            formItems.Add("FontSize", fontSize);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "VatAmount"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "@BDOSARDV");
            formItems.Add("Alias", "U_VatAmount");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }


            GC.Collect();
        }

        public static void JournalEntryTransaction(  SAPbouiCOM.Form oForm, bool ActionSuccess, bool BeforeAction, out bool BubbleEvent, out string errorText)
        {
            errorText = null;
            BubbleEvent = true;


            if (ActionSuccess != BeforeAction)
            {
                


                //დოკუმენტის გატარების დროს გატარდეს ბუღლტრული გატარება
                oForm.Refresh();
                SAPbouiCOM.DBDataSource DocDBSourcePAYR = oForm.DataSources.DBDataSources.Item(0);

                if (DocDBSourcePAYR.GetValue("CANCELED", 0) == "N")
                {
                    string DocEntry = DocDBSourcePAYR.GetValue("DocEntry", 0);
                    string DocCurrency = "";
                    decimal DocRate = 0;
                    string DocNum = DocDBSourcePAYR.GetValue("DocNum", 0);
                    DateTime DocDate = DateTime.ParseExact(DocDBSourcePAYR.GetValue("U_DocDate", 0), "yyyyMMdd", CultureInfo.InvariantCulture);

                    CommonFunctions.StartTransaction();

                    Program.JrnLinesGlobal = new DataTable();
                    
                    DataTable JrnLinesDT = createAdditionalEntries( oForm, null, null, DocCurrency,  DocRate);

                    JrnEntry( DocEntry, DocNum, DocDate, JrnLinesDT,  out errorText);
                    if (errorText != null)
                    {
                        Program.uiApp.MessageBox(errorText);
                        BubbleEvent = false;
                    }
                    else
                    {
                        if (ActionSuccess == false)
                        {
                            Program.JrnLinesGlobal = JrnLinesDT;
                        }
                    }

                    if (Program.oCompany.InTransaction)
                    {
                        //თუ დოკუმენტი გატარდა, მერე ვაკეთებს ბუღალტრულ გატარებას
                        if (ActionSuccess == true & BeforeAction == false)
                        {
                            CommonFunctions.EndTransaction( SAPbobsCOM.BoWfTransOpt.wf_Commit);
                        }
                        else
                        {
                            CommonFunctions.EndTransaction( SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                        }
                    }
                    else
                    {
                        Program.uiApp.MessageBox("ტრანზაქციის დასრულებს შეცდომა");
                        BubbleEvent = false;

                    }
                }


            }
        }

        public static void JrnEntry( string DocEntry, string DocNum, DateTime DocDate, DataTable JrnLinesDT, out string errorText)
        {
            errorText = null;

            try
            {
                JournalEntry.JrnEntry( DocEntry, "UDO_F_BDO_ARDPV_D", "A/R Down Payment VAT Accrual: " + DocNum, DocDate, JrnLinesDT,  out errorText);

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

        public static void formDataLoad( SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            openFormEvent = false;

            SAPbouiCOM.StaticText oStaticText = null;
            oForm.Freeze(true);
            try
            {
               
                int docEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item(0).GetValue("DocEntry", 0));

                // გატარებები
                SAPbouiCOM.DBDataSource DocDBSourceTAXP = oForm.DataSources.DBDataSources.Item(0);
                string Ref1 = docEntry.ToString();
                string Ref2 = "UDO_F_BDO_ARDPV_D";

                SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string query = "SELECT " +
                                "*  " +
                                "FROM \"OJDT\"  " +
                                "WHERE \"StornoToTr\" IS NULL " +
                                "AND \"Ref1\" = '" + Ref1 + "' " +
                                "AND \"Ref2\" = '" + Ref2 + "' ";
                oRecordSet.DoQuery(query);

                if (!oRecordSet.EoF)
                {                    
                    oForm.DataSources.UserDataSources.Item("BDOSJrnEnt").ValueEx = oRecordSet.Fields.Item("TransId").Value.ToString();
                }
                else
                {                    
                    oForm.DataSources.UserDataSources.Item("BDOSJrnEnt").ValueEx = "";
                }


                //-------------------------------------------ანგარიშ-ფაქტურა----------------------------------->
                string cardCode = oForm.DataSources.DBDataSources.Item(0).GetValue("U_cardCode", 0).Trim();
                string caption = BDOSResources.getTranslate("CreateTaxInvoice");
                int taxDocEntry = 0;
                string taxID = "";
                string taxNumber = "";
                string taxSeries = "";
                string taxStatus = "";
                string taxCreateDate = "";

                if (docEntry != 0)
                {
                    Dictionary<string, object> taxDocInfo = BDO_TaxInvoiceSent.getTaxInvoiceSentDocumentInfo( docEntry, "ARDownPaymentVAT", cardCode, out errorText);
                    if (taxDocInfo != null)
                    {
                        taxDocEntry = Convert.ToInt32(taxDocInfo["docEntry"]);
                        taxID = taxDocInfo["invID"].ToString();
                        taxNumber = taxDocInfo["number"].ToString();
                        taxSeries = taxDocInfo["series"].ToString();
                        taxStatus = taxDocInfo["status"].ToString();
                        taxCreateDate = taxDocInfo["createDate"].ToString();

                        if (taxDocEntry != 0)
                        {
                            DateTime taxCreateDateDT = DateTime.ParseExact(taxCreateDate, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
                            if (taxSeries == "")
                            {
                                caption = BDOSResources.getTranslate("TaxInvoiceDate") + " " + taxCreateDateDT;
                            }
                            else
                            {
                                caption = BDOSResources.getTranslate("TaxInvoiceSeries") + " " + taxSeries + " № " + taxNumber + " " + BDOSResources.getTranslate("Data") + " " + taxCreateDateDT;
                            }
                        }
                    }
                }
                else
                {
                    taxDocEntry = 0;
                }

                oForm.DataSources.UserDataSources.Item("BDO_TaxDoc").ValueEx = taxDocEntry == 0 ? "" : taxDocEntry.ToString();
                oForm.Items.Item("BDO_TaxDoc").Enabled = false;
                oForm.Items.Item("BDOSJrnEnt").Enabled = false;

                oStaticText = (SAPbouiCOM.StaticText)oForm.Items.Item("BDO_TaxTxt").Specific;
                oStaticText.Caption = caption;
            }
            catch (Exception ex)
            {
                oForm.DataSources.UserDataSources.Item("BDO_TaxDoc").ValueEx = "";
                oForm.Items.Item("BDOSJrnEnt").Specific.Value = "";
                

                oStaticText = (SAPbouiCOM.StaticText)oForm.Items.Item("BDO_TaxTxt").Specific;
                oStaticText.Caption = BDOSResources.getTranslate("CreateTaxInvoice");

                errorText = ex.Message;
            }
            finally
            {
                oForm.Freeze(false);
                GC.Collect();
            }
        }


        public static DataTable createAdditionalEntries( SAPbouiCOM.Form oForm, SAPbobsCOM.GeneralData oGeneralData, DataTable DTSource, string DocCurrency, decimal DocRate)
        {
            
            DataTable jeLines = JournalEntry.JournalEntryTable();
            DataTable AccountTable = CommonFunctions.GetOACTTable();
            SAPbobsCOM.GeneralDataCollection oChild = null;            
            

            SAPbouiCOM.DBDataSource DBDataSourceTable = null;
            int JEcount = 0;

            DateTime DocDate = new DateTime();
           
            if (oForm == null)
            {
                oChild = oGeneralData.Child("BDOSRDV1");
                JEcount = oChild.Count;
                DocDate = oGeneralData.GetProperty("U_DocDate");
            }
            else
            {
                SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("ItemsMTR").Specific;
                oMatrix.FlushToDataSource();
                DBDataSourceTable = oForm.DataSources.DBDataSources.Item("@BDOSRDV1");
                JEcount = DBDataSourceTable.Size;
                DocDate = DateTime.ParseExact(oForm.DataSources.DBDataSources.Item("@BDOSARDV").GetValue("U_DocDate", 0), "yyyyMMdd", CultureInfo.InvariantCulture);
            }

            
            string year = DocDate.Year.ToString();

            string CreditAccount = "";
            string DebitAccount = CommonFunctions.getPeriodsCategory( "SaleVatOff", year);

            DocCurrency = DocCurrency == CommonFunctions.getLocalCurrency() ? "" : DocCurrency;



            for (int i = 0; i < JEcount; i++)
            {
                decimal vatAmount = FormsB1.cleanStringOfNonDigits( CommonFunctions.getChildOrDbDataSourceValue(DBDataSourceTable, oChild, null, "U_VatAmount", i).ToString());
                decimal vatAmountFC = DocCurrency == "" ? 0 : vatAmount / DocRate;

                if(vatAmount==0)
                { 
                    continue;
                }

                string VatRate = CommonFunctions.getChildOrDbDataSourceValue(DBDataSourceTable, oChild, null, "U_VatGrp", i).ToString();
                SAPbobsCOM.VatGroups oVatCode;
                oVatCode = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVatGroups);
                oVatCode.GetByKey(VatRate);
                CreditAccount = oVatCode.TaxAccount;

                JournalEntry.AddJournalEntryRow(AccountTable, jeLines, "OnlyDebit", DebitAccount, CreditAccount, vatAmount, vatAmountFC, DocCurrency, "", "", "", "", "", "", "", "");
                JournalEntry.AddJournalEntryRow(AccountTable, jeLines, "OnlyCredit", DebitAccount, CreditAccount, vatAmount, vatAmountFC, DocCurrency, "", "", "", "", "", "", VatRate, "");
            }

            return jeLines;
        }

        public static void setSizeForm( SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            try
            {
                int formHeight = Program.uiApp.Desktop.Width;
                int formWidth = Program.uiApp.Desktop.Width;

                oForm.ClientHeight = formHeight;
                oForm.Height = formWidth;
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
        }
    }
}
