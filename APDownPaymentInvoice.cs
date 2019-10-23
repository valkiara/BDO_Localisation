using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace BDO_Localisation_AddOn
{
    static partial class APDownPaymentInvoice
    {
        public static void formDataLoad( SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            SAPbouiCOM.StaticText oStaticText = null;
            oForm.Freeze(true);
            try
            {
                int docEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("ODPO").GetValue("DocEntry", 0));

                //-------------------------------------------ანგარიშ-ფაქტურა----------------------------------->
                string cardCode = oForm.DataSources.DBDataSources.Item("ODPO").GetValue("CardCode", 0).Trim();
                string caption = BDOSResources.getTranslate("ChooseTaxInvoice");
                int taxDocEntry = 0;
                string taxID = "";
                string taxNumber = "";
                string taxSeries = "";
                string taxStatus = "";
                string taxCreateDate = "";

                if (docEntry != 0)
                {
                    Dictionary<string, object> taxDocInfo = BDO_TaxInvoiceReceived.getTaxInvoiceReceivedDocumentInfo( docEntry, "2", cardCode, out errorText);
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
                            DateTime taxCreateDateDT = DateTime.ParseExact(taxCreateDate, "yyyyMMdd", CultureInfo.InvariantCulture);
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
                oForm.DataSources.UserDataSources.Item("BDO_TaxSer").ValueEx = taxSeries;
                oForm.DataSources.UserDataSources.Item("BDO_TaxNum").ValueEx = taxNumber;
                oForm.DataSources.UserDataSources.Item("BDO_TaxDat").ValueEx = taxCreateDate;
                
                oStaticText = (SAPbouiCOM.StaticText)oForm.Items.Item("BDO_TaxTxt").Specific;
                oStaticText.Caption = caption;
                //<-------------------------------------------ანგარიშ-ფაქტურა-----------------------------------
            }
            catch (Exception ex)
            {
                oForm.DataSources.UserDataSources.Item("BDO_TaxDoc").ValueEx = "";
                oForm.DataSources.UserDataSources.Item("BDO_TaxSer").ValueEx = "";
                oForm.DataSources.UserDataSources.Item("BDO_TaxNum").ValueEx = "";
                oForm.DataSources.UserDataSources.Item("BDO_TaxDat").ValueEx = "";
                
                oStaticText = (SAPbouiCOM.StaticText)oForm.Items.Item("BDO_TaxTxt").Specific;
                oStaticText.Caption = BDOSResources.getTranslate("ChooseTaxInvoice");
                SAPbouiCOM.Item oItem = oForm.Items.Item("BDO_TaxCan");
                oItem.Visible = false;

                errorText = ex.Message;
            }
            finally
            {
                oForm.Freeze(false);
                GC.Collect();
            }
        }

        public static void formDataAddUpdate( SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            try
            {
                string taxDocEntry = oForm.DataSources.UserDataSources.Item("BDO_TaxDoc").ValueEx;
                string docEntryStr = oForm.DataSources.DBDataSources.Item("ODPO").GetValue("DocEntry", 0);
                if (string.IsNullOrEmpty(taxDocEntry) == false && string.IsNullOrEmpty(docEntryStr) == false)
                {
                    int docEntry = Convert.ToInt32(docEntryStr);
                    double baseDocGTotal = 0;
                    double baseDocLineVat = 0;
                    getAmount( docEntry, out baseDocGTotal, out baseDocLineVat, out errorText);
                    BDO_TaxInvoiceReceived.addBaseDoc( Convert.ToInt32(taxDocEntry), docEntry, "2", null, baseDocGTotal, baseDocLineVat, out errorText);
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

        public static void uiApp_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(BusinessObjectInfo.FormTypeEx, Program.currentFormCount);

            if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD & BusinessObjectInfo.BeforeAction == true)
            {
                oForm.Freeze(true);
                int panelLevel = oForm.PaneLevel;
                string sdocDate = oForm.Items.Item("10").Specific.Value;
                oForm.PaneLevel = 7;
                oForm.Items.Item("1000").Specific.Value = sdocDate;
                oForm.PaneLevel = panelLevel;
                oForm.Freeze(false);
            }

            if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD & BusinessObjectInfo.BeforeAction == false)
            {
                formDataLoad(oForm, out errorText);
                setVisibleFormItems(oForm, out errorText);
            }

            if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE)
            {
                if (BusinessObjectInfo.BeforeAction == true)
                {
                    formDataAddUpdate(oForm, out errorText);
                    if (string.IsNullOrEmpty(errorText) == false)
                    {
                        Program.uiApp.MessageBox(errorText);
                        BubbleEvent = false;
                    }
                }
                else if (BusinessObjectInfo.BeforeAction == false && BusinessObjectInfo.ActionSuccess == true)
                {
                    setVisibleFormItems(oForm, out errorText);
                }
            }
        }

        public static void getAmount( int docEntry, out double gTotal, out double lineVat, out string errorText)
        {
            errorText = null;
            gTotal = 0;
            lineVat = 0;

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string query = @"SELECT 
            ""DPO1"".""DocEntry"" AS ""docEntry"", 
            SUM(""DPO1"".""GTotal"") AS ""GTotal"", 
            SUM(""DPO1"".""LineVat"") AS ""LineVat"" 
            FROM ""DPO1"" AS ""DPO1"" 
            WHERE ""DPO1"".""DocEntry"" = '" + docEntry + @"' 
            GROUP BY ""DPO1"".""DocEntry""";

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
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                oRecordSet = null;
                GC.Collect();
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
                    if (sCFL_ID == "TaxInvoiceReceived_CFL")
                    {
                            string cardCode = oForm.DataSources.DBDataSources.Item("ODPO").GetValue("CardCode", 0).Trim();
                            DateTime docDate = DateTime.ParseExact(oForm.DataSources.DBDataSources.Item("ODPO").GetValue("DocDate", 0), "yyyyMMdd", CultureInfo.InvariantCulture);

                            List<string> taxInvoiceDocList = BDO_TaxInvoiceReceived.getListTaxInvoiceReceived( cardCode, null, "2", docDate, out errorText);

                            int docCount = taxInvoiceDocList.Count;
                            SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();
                            if (docCount == 0)
                            {
                                SAPbouiCOM.Condition oCon = oCons.Add();
                                oCon.Alias = "DocEntry";
                                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                                oCon.CondVal = "";
                            }
                            else
                            {
                                for (int i = 0; i < docCount; i++)
                                {
                                    SAPbouiCOM.Condition oCon = oCons.Add();
                                    oCon.Alias = "DocEntry";
                                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                                    oCon.CondVal = taxInvoiceDocList[i];
                                    oCon.Relationship = (i == docCount - 1) ? SAPbouiCOM.BoConditionRelationship.cr_NONE : SAPbouiCOM.BoConditionRelationship.cr_OR;
                                }
                            }
                            oCFL.SetConditions(oCons);
                        }
                    }
                else if (beforeAction == false)
                {
                    SAPbouiCOM.DataTable oDataTable = null;
                    oDataTable = oCFLEvento.SelectedObjects;

                    if (oDataTable != null)
                    {
                        if (sCFL_ID == "TaxInvoiceReceived_CFL")
                        {
                            string taxDocEntryStr = oDataTable.GetValue("DocEntry", 0).ToString();
                            BDO_TaxInvoiceReceived.chooseFromListForBaseDocs( oForm, taxDocEntryStr, out errorText);
                        }
                    }
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

        public static void setVisibleFormItems(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            SAPbouiCOM.Item oItem = null;
            oForm.Freeze(true);

            try
            {
                string docEntrySTR = oForm.DataSources.DBDataSources.Item("ODPO").GetValue("DocEntry", 0);

                oItem = oForm.Items.Item("BDO_TaxTxt");
                oItem.Visible = true;
                oItem = oForm.Items.Item("BDO_TaxLB");
                oItem.Visible = true;
                oItem = oForm.Items.Item("BDO_TaxDoc");
                oItem.Visible = true;

                if (oForm.DataSources.UserDataSources.Item("BDO_TaxDoc").ValueEx != "")
                {
                    oItem = oForm.Items.Item("BDO_TaxCan");
                    oItem.Visible = true;
                    oItem = oForm.Items.Item("BDO_TaxDoc");
                    oItem.Enabled = false;
                }
                else if (string.IsNullOrEmpty(docEntrySTR))
                {
                    oItem = oForm.Items.Item("BDO_TaxDoc");
                    oItem.Enabled = false;
                }
                else
                {
                    oItem = oForm.Items.Item("BDO_TaxCan");
                    oItem.Visible = false;
                    oItem = oForm.Items.Item("BDO_TaxDoc");
                    oItem.Enabled = true;

                    SAPbouiCOM.EditText oEditText = (SAPbouiCOM.EditText)oItem.Specific;
                    oEditText.ChooseFromListUID = "TaxInvoiceReceived_CFL";
                    oEditText.ChooseFromListAlias = "DocEntry";
                }

                oItem = oForm.Items.Item("liablePrTx");
                oItem.Visible = false;
                oItem = oForm.Items.Item("PrBaseS");
                oItem.Visible = false;
                oItem = oForm.Items.Item("PrBaseE");
                oItem.Visible = false;
                oItem = oForm.Items.Item("PrBsDscr");
                oItem.Visible = false;
                oItem = oForm.Items.Item("PrBaseLB");
                oItem.Visible = false;
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
            finally
            {
                oForm.Freeze(false);
                oForm.Update();
                GC.Collect();
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
                    APDownPayment.createFormItems( oForm, out errorText);
                    formDataLoad( oForm, out errorText);
                    setVisibleFormItems(oForm, out errorText);
                }               
                if (pVal.ItemUID == "BDO_TaxDoc" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                {
                    SAPbouiCOM.IChooseFromListEvent oCFLEvento = null;
                    oCFLEvento = ((SAPbouiCOM.IChooseFromListEvent)(pVal));
                        chooseFromList( oForm, oCFLEvento, pVal.ItemUID, pVal.BeforeAction, out errorText);
                }
                if (pVal.ItemUID == "BDO_TaxCan" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK & pVal.BeforeAction == false)
                {
                    int taxDocEntry = Convert.ToInt32(oForm.DataSources.UserDataSources.Item("BDO_TaxDoc").ValueEx.Trim());
                    int docEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("ODPO").GetValue("DocEntry", 0));
                    if (taxDocEntry != 0)
                    {
                        int answer = Program.uiApp.MessageBox(BDOSResources.getTranslate("DocumentLinkedToTaxInvoiceCancel"), 1, BDOSResources.getTranslate("Yes"), BDOSResources.getTranslate("No"), "");
                        if (answer == 1)
                        {
                            BDO_TaxInvoiceReceived.removeBaseDoc( taxDocEntry, docEntry, "2", out errorText);
                            formDataLoad( oForm, out errorText);
                            setVisibleFormItems(oForm, out errorText);
                        }
                    }
                    else
                    {
                        BubbleEvent = false;
                    }
                }
            }
        }

        public static List<int> getAllConnectedDoc( List<int> docEntry, out string errorText)
        {
            errorText = null;
            List<int> connectedDocList = new List<int>();

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string query = @"SELECT
            	 ""DPO1"".""DocEntry"" 
            FROM ""DPO1"" 
            WHERE ""DPO1"".""BaseEntry"" IN (SELECT
            	 ""DPO1"".""BaseEntry"" 
            	FROM ""DPO1"" 
            	WHERE ""DPO1"".""DocEntry"" IN (" + string.Join(",", docEntry) + @") 
            	AND ""DPO1"".""BaseType"" = '204') 
            AND ""DPO1"".""BaseType"" = '204' 
            AND ""DPO1"".""DocEntry"" NOT IN (" + string.Join(",", docEntry) + @")
            GROUP BY ""DPO1"".""DocEntry""";

            try
            {
                oRecordSet.DoQuery(query);
                if (oRecordSet.RecordCount > 0)
                {
                    while (!oRecordSet.EoF)
                    {
                        connectedDocList.Add(Convert.ToInt32(oRecordSet.Fields.Item("DocEntry").Value));
                        oRecordSet.MoveNext();
                    }                 
                }
                return connectedDocList;
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
                return connectedDocList;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                oRecordSet = null;
                GC.Collect();
            }
        }
    }
}
