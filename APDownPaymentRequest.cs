using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Globalization;
using System.Data;

namespace BDO_Localisation_AddOn
{
    static partial class APDownPaymentRequest
    {
        public static void createFormItems(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            Dictionary<string, object> formItems = null;

            string itemName = "";
    

            // -------------------- Use blanket agreement rates-----------------
            int pane = 7;
            int left = oForm.Items.Item("1720002167").Left;
            int height = oForm.Items.Item("1720002167").Height;
            int top = oForm.Items.Item("1720002167").Top + height + 5;

            formItems = new Dictionary<string, object>();
            itemName = "UsBlaAgRtS"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "ODPO");
            formItems.Add("Alias", "U_UseBlaAgRt");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
            formItems.Add("Length", 1);
            formItems.Add("Left", left);
            formItems.Add("Width", 100);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("UseBlAgrRt"));
            formItems.Add("ValOff", "N");
            formItems.Add("ValOn", "Y");
            formItems.Add("DisplayDesc", true);
            formItems.Add("SetAutoManaged", true);
            formItems.Add("FromPane", pane);
            formItems.Add("ToPane", pane);


            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            GC.Collect();
        }

        public static bool ProfitTaxTypeIsSharing = false;

        public static void chooseFromList( SAPbouiCOM.Form oForm, SAPbouiCOM.IChooseFromListEvent oCFLEvento, string itemUID, bool beforeAction, out string errorText)
        {
            errorText = null;

            try
            {
                string sCFL_ID = oCFLEvento.ChooseFromListUID;

                if (beforeAction == false)
                {
                    SAPbouiCOM.DataTable oDataTable = null;
                    oDataTable = oCFLEvento.SelectedObjects;

                    if (oDataTable != null)
                    {
                        if (sCFL_ID == "CFL_ProfitBase")
                        {
                            string ProfitBaseCode = Convert.ToString(oDataTable.GetValue("Code", 0));
                            string ProfitBaseName = Convert.ToString(oDataTable.GetValue("Name", 0));

                            try
                            {
                                SAPbouiCOM.EditText oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("PrBaseE").Specific;
                                oEditText.Value = ProfitBaseCode;
                            }
                            catch { }

                            try
                            {
                                SAPbouiCOM.EditText oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("PrBsDscr").Specific;
                                oEditText.Value = ProfitBaseName;
                            }
                            catch { }
                        }
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

        public static void formDataLoad( SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            oForm.Freeze(true);

            try
            {
                //setVisibleFormItems(oForm, out errorText);
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

        public static void setVisibleFormItems(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            SAPbouiCOM.Item oItem = null;
            oForm.Freeze(true);

            try
            {
                oItem = oForm.Items.Item("BDO_TaxTxt");
                oItem.Visible = false;
                oItem = oForm.Items.Item("BDO_TaxDoc");
                oItem.Visible = false;
                oItem = oForm.Items.Item("BDO_TaxLB");
                oItem.Visible = false; 
                oItem = oForm.Items.Item("BDO_TaxCan");
                oItem.Visible = false;

                oItem = oForm.Items.Item("liablePrTx");
                oItem.Visible = true;
                oItem = oForm.Items.Item("PrBaseS");
                oItem.Visible = true;
                oItem = oForm.Items.Item("PrBaseE");
                oItem.Visible = true;
                oItem = oForm.Items.Item("PrBsDscr");
                oItem.Visible = true;
                oItem = oForm.Items.Item("PrBaseLB");
                oItem.Visible = true;

                string docEntry = oForm.DataSources.DBDataSources.Item("ODPO").GetValue("DocEntry", 0).Trim();
                bool docEntryIsEmpty = string.IsNullOrEmpty(docEntry);

                if (ProfitTaxTypeIsSharing == true)
                {
                    oForm.Items.Item("16").Click(SAPbouiCOM.BoCellClickType.ct_Regular); //მისაწვდომობის შეზღუდვისთვის

                    oForm.Items.Item("liablePrTx").Enabled = (docEntryIsEmpty == true);

                    bool LiablePrTx = oForm.DataSources.DBDataSources.Item("ODPO").GetValue("U_liablePrTx", 0) == "Y";
                    oForm.Items.Item("PrBaseE").Enabled = (LiablePrTx && docEntryIsEmpty == true);

                    string uniqueID_lf_ProfitBaseCFL = "CFL_ProfitBase";
                    oForm.Items.Item("PrBaseE").Specific.ChooseFromListUID = uniqueID_lf_ProfitBaseCFL;
                    oForm.Items.Item("PrBaseE").Specific.ChooseFromListAlias = "Code";
                }
                else
                {
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

        public static void FillDefaultValuesProfitTax( SAPbouiCOM.Form oForm, string CardCode, out string errorText)
        {
            errorText = null;
            oForm.Freeze(true);

            try
            {
                bool NoRecords = true;

                if (CardCode == "")
                {
                    CardCode = oForm.DataSources.DBDataSources.Item("ODPO").GetValue("CardCode", 0).Trim();
                }

                if (CardCode != "")
                {
                    SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    string query = @"SELECT  ""CardCode"",
		                                 ""U_BDO_prBsDR"",
		                                 ""U_BDO_prDRDs"" 
                                FROM ""OCRD"",
	                                 ""OADM"" 
                                WHERE (""U_BDO_PTExem"" = 'Y' OR ""U_BDO_RIOfsh"" = 'Y' OR ""U_BDO_PhysTp"" <> '')" +
                                    @" AND ""OCRD"".""CardCode"" = '" + CardCode.Replace("'", "''") + "' ";

                    oRecordSet.DoQuery(query);
                    if (!oRecordSet.EoF)
                    {
                        SAPbouiCOM.CheckBox oCheck = oForm.Items.Item("liablePrTx").Specific;
                        oCheck.Checked = true;

                        //oForm.Items.Item("PrBaseE").Enabled = true;
                        try
                        {
                            SAPbouiCOM.EditText oEdit = oForm.Items.Item("PrBaseE").Specific;
                            oEdit.Value = oRecordSet.Fields.Item("U_BDO_prBsDR").Value;
                        }
                        catch { }

                        try
                        {
                            SAPbouiCOM.EditText oEdit = oForm.Items.Item("PrBsDscr").Specific;
                            oEdit.Value = oRecordSet.Fields.Item("U_BDO_prDRDs").Value;
                        }
                        catch { }
                        NoRecords = false;
                    }
                }

                if (CardCode == "" || NoRecords == true)
                {
                    SAPbouiCOM.CheckBox oCheck = oForm.Items.Item("liablePrTx").Specific;
                    oCheck.Checked = false;

                    try
                    {
                        SAPbouiCOM.EditText oEdit = oForm.Items.Item("PrBaseE").Specific;
                        oEdit.Value = "";
                    }
                    catch { }

                    try
                    {
                        SAPbouiCOM.EditText oEdit = oForm.Items.Item("PrBsDscr").Specific;
                        oEdit.Value = "";
                    }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
            finally
            {
                GC.Collect();
                oForm.Update();
                oForm.Freeze(false);
            }
        }

        public static void uiApp_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(BusinessObjectInfo.FormTypeEx, Program.currentFormCount);

            if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD & BusinessObjectInfo.BeforeAction == false)
            {
                formDataLoad(oForm, out errorText);
                setVisibleFormItems(oForm, out errorText);
            }
            if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD)
            {
                if (BusinessObjectInfo.BeforeAction == true)
                {
                    SAPbouiCOM.DBDataSource DocDBSourceOCRD = oForm.DataSources.DBDataSources.Item(0);

                    // მოგების გადასახადი
                    if (ProfitTaxTypeIsSharing == true)
                    {
                        if (oForm.DataSources.DBDataSources.Item("ODPO").GetValue("U_liablePrTx", 0) == "Y")
                        {
                            if (oForm.DataSources.DBDataSources.Item("OIGE").GetValue("U_prBase", 0) == "")
                            {
                                Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("TaxableObject") + " " + BDOSResources.getTranslate("YouCantLeaveEmpty"));
                                Program.uiApp.MessageBox(BDOSResources.getTranslate("OperationUnsuccesfullSeeLog"));
                                BubbleEvent = false;
                            }
                        }
                    }
                }
            }
        }

        public static void uiApp_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText;

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD & pVal.BeforeAction == true)
                {
                    ProfitTaxTypeIsSharing = ProfitTax.ProfitTaxTypeIsSharing();
                    APDownPayment.createFormItems(oForm, out errorText);
                    formDataLoad(oForm, out errorText);
                    setVisibleFormItems(oForm, out errorText);
                    createFormItems(oForm, out errorText);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_RESIZE && pVal.BeforeAction == false)
                {
                    oForm.Freeze(true);
                    resizeForm(oForm, out errorText);
                    oForm.Freeze(false);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
                {
                    if ((pVal.ItemUID == "liablePrTx") && pVal.BeforeAction == false)
                    {
                        oForm.Freeze(true);
                        LiableTaxes_OnClick(oForm, out errorText);
                        oForm.Freeze(false);
                    }

                    if (pVal.ItemUID == "1" && pVal.BeforeAction == true)
                    {
                        CommonFunctions.fillDocRate(oForm, "ODPO", "ODPO");
                    }

                    if (pVal.ItemUID == "UsBlaAgRtS" & pVal.BeforeAction == false)
                    {

                        SAPbouiCOM.EditText oBlankAgr = (SAPbouiCOM.EditText)oForm.Items.Item("1980002192").Specific;

                        if (string.IsNullOrEmpty(oBlankAgr.Value))
                        {
                            Program.uiApp.SetStatusBarMessage(errorText = BDOSResources.getTranslate("EmptyBlaAgrError"), SAPbouiCOM.BoMessageTime.bmt_Short);
                            SAPbouiCOM.CheckBox oUsBlaAgRtCB = (SAPbouiCOM.CheckBox)oForm.Items.Item("UsBlaAgRtS").Specific;
                            oUsBlaAgRtCB.Checked = false;
                            oForm.Items.Item("1980002192").Click();
                        }
                    }

                    if ((pVal.ItemUID == "PrBaseE") & pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                    {
                        if (pVal.BeforeAction == false)
                        {
                            SAPbouiCOM.ChooseFromListEvent oCFLEvento = null;
                            oCFLEvento = ((SAPbouiCOM.ChooseFromListEvent)(pVal));

                            chooseFromList(oForm, oCFLEvento, pVal.ItemUID, pVal.BeforeAction, out errorText);
                        }
                    }
                    if (pVal.ItemUID == "4" && pVal.BeforeAction == false) //& pVal.InnerEvent == false
                    {
                        if (pVal.ItemUID == "4" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST && pVal.BeforeAction == false)
                        {
                            SAPbouiCOM.ChooseFromListEvent oCFLEvento = null;
                            oCFLEvento = ((SAPbouiCOM.ChooseFromListEvent)(pVal));
                            string CardCode = oCFLEvento.SelectedObjects.GetValue("CardCode", 0);

                            FillDefaultValuesProfitTax(oForm, CardCode, out errorText);
                            setVisibleFormItems(oForm, out errorText);
                        }
                    }
                }
            }
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

        public static void reArrangeFormItems(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.Item oItem = null;

            oItem = oForm.Items.Item("70");
            int top = oItem.Top;
            int height = oItem.Height; ;

            top = top + height + 5;
            oItem = oForm.Items.Item("liablePrTx");
            oItem.Top = top;

            top = top + height + 1;
            oItem = oForm.Items.Item("PrBaseS");
            oItem.Top = top;
            oItem = oForm.Items.Item("PrBaseE");
            oItem.Top = top;
            oItem = oForm.Items.Item("PrBsDscr");
            oItem.Top = top;
            oItem = oForm.Items.Item("PrBaseLB");
            oItem.Top = top;
        }

        public static void LiableTaxes_OnClick( SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            string liablePrTx = oForm.DataSources.DBDataSources.Item("ODPO").GetValue("U_liablePrTx", 0).Trim();

            if (liablePrTx != "Y")
            {
                SAPbouiCOM.EditText oEdit = oForm.Items.Item("PrBaseE").Specific;
                oEdit.Value = "";

                oEdit = oForm.Items.Item("PrBsDscr").Specific;
                oEdit.Value = "";
            }
            setVisibleFormItems(oForm, out errorText);
        }
    }
}

