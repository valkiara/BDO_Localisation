using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using System.Data;
using System.Resources;
using System.IO;
using System.Reflection;
using System.Data.SqlClient;
using BDO_Localisation_AddOn.TBC_Integration_Services;
using BDO_Localisation_AddOn.BOG_Integration_Services;
using System.Runtime.InteropServices;

namespace BDO_Localisation_AddOn
{
    static class Program
    {
        public static SAPbouiCOM.Application uiApp;
        public static SAPbobsCOM.Company oCompany;
        public static SAPbouiCOM.Form oExchangeFormRatesAndIndexes;
        public static string BDO_SU;
        public static int USERID;
        public static bool FORM_LOAD_FOR_VISIBLE = false;
        public static bool FORM_LOAD_FOR_ACTIVATE = false;
        public static bool cancellationDoc = false;
        public static bool cancellationTrans = false;
        public static int canceledDocEntry = 0;
        public static int removeRecordRow = 0;
        public static bool removeRecordTrans = false;
        public static bool removeLineTrans = false;
        public static SAPbouiCOM.Form oIncWaybDocFormAPInv;
        public static SAPbouiCOM.Form oIncWaybDocFormCrMemo;
        public static SAPbouiCOM.Form oIncWaybDocFormAPCorInv;
        public static SAPbouiCOM.Form oIncWaybDocFormGdsRecpPO;
        public static int currentFormCount = 1;
        public static CultureInfo cultureInfo = null;
        public static ResourceManager resourceManager = null;
        public static string LocalCurrency = null;
        public static string MainCurrency = null;
        public static bool openPaymentMeans = false;
        public static decimal transferSumFC;
        public static decimal overallAmount;
        public static string newPostDateStr;
        public static DataTable paymentInvoices;
        public static bool openPaymentMeansByPostDateChange = false;
        public static DataTable JrnLinesGlobal = new DataTable();
        public static bool clickUnitedJournalEntry = false;
        public static bool Exchange_Rate_Save_Click = false;
        public static DataTable UserDefinedFieldsCurrentCompany;
        public static DataTable UserDefinedTablesCurrentCompany;
        public static bool localisationAddonLicensed = false;
        public static readonly string ExecutionDateISO = DateTime.UtcNow.ToString("o");
        public static bool selectItemsToCopyOkClick = false;
        

    static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            if (runAddOn())
            {
                BDOSAutomaticTasks.importCurrencyRate();

                uiApp.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(uiApp_ItemEvent);
                uiApp.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(uiApp_MenuEvent);
                uiApp.FormDataEvent += new SAPbouiCOM._IApplicationEvents_FormDataEventEventHandler(uiApp_FormDataEvent);
                uiApp.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(uiApp_AppEvent);
                uiApp.LayoutKeyEvent += new SAPbouiCOM._IApplicationEvents_LayoutKeyEventEventHandler(uiApp_LayoutKeyEvent);
                uiApp.RightClickEvent += new SAPbouiCOM._IApplicationEvents_RightClickEventEventHandler(uiApp_RightClickEvent);

                Application.Run();
            }
            else
            {
                Application.Exit();
                return;
            }
        }

        static bool runAddOn()
        {
            string errorText;

            bool connectResult = ConnectB1.connectShared(out errorText);
            if (connectResult)
            {
                //SAPbouiCOM.ProgressBar ProgressBarForm;
                //ProgressBarForm = uiApp.StatusBar.CreateProgressBar("", 20, true);
                //ProgressBarForm.Value = 0;

                BDOSResources.initResource(Convert.ToInt32(oCompany.language), out cultureInfo, out resourceManager, out errorText);

                uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("AddonConnectedSuccesfully"), SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                UserDefinedTablesCurrentCompany = UDO.UserDefinedTablesCurrentCompany();
                UserDefinedFieldsCurrentCompany = UDO.UserDefinedFieldsCurrentCompany();

                if (!UDO.UserDefinedFieldExists("OADM", "BDOSLocLic"))
                {
                    License.createUserFields(out errorText);
                    if (!String.IsNullOrEmpty(errorText))
                    {
                        uiApp.StatusBar.SetSystemMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("LocalisationLicensingDataCouldNotBeCreated") + BDOSResources.getTranslate("RetryStartingAddon"), SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        return false;
                    }
                }

                try
                {
                    License.UpdateAddOnLicense();
                }
                catch { }

                if (!runLocalisationAddOn()) return false;

                SAPbobsCOM.SBObob oSBOBob = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);
                LocalCurrency = CommonFunctions.getCurrencyInternationalCode(oSBOBob.GetLocalCurrency().Fields.Item("LocalCurrency").Value);

                MainCurrency = CurrencyB1.getMainCurrency(out errorText);

                return connectResult;
            }
            else
            {
                MessageBox.Show(errorText);
                return connectResult;
            }
        }

        static bool runLocalisationAddOn()
        {
            string errorText;

            Dictionary<string, string> CompanyLicenseInfo = CommonFunctions.getCompanyLicenseInfo();
            if (CompanyLicenseInfo["LicenseStatus"] == BDOSResources.getTranslate("Active"))
            {
                localisationAddonLicensed = true;

                BDO_BPCatalog.updateFields();

                /////////////////
                string version = "1.1.5.1";

                BDOSTablesLog.CreateTable(out errorText);

                if ((UDO.UserDefinedTableExists("BDOSLOGS")) == false)
                {
                    if (!String.IsNullOrEmpty(errorText))
                    {
                        uiApp.StatusBar.SetSystemMessage(errorText);
                    }

                    uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("LogTableCouldNotBeCreated") + BDOSResources.getTranslate("RetryStartingAddon"), SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return false;
                }

                CompanyDetails.createUserFields(out errorText);

                UDO.allUDOForAddOn(out errorText);
                FormsB1.allUserFieldsForAddOn(out errorText);

                BDOSInternetBankingIntegrationServicesRules.updateUDO();
                BDO_TaxInvoiceReceived.updateUDO();

                updateAddonVersion(version);

                FormsB1.addMenusForAddOn();

                CrystalReports.addStandAloneCrystalReportForAddOn(Application.StartupPath, out errorText);

                CrystalReports.addDocumentTypeCrystalReportForAddOn(Application.StartupPath, out errorText);

                uiApp.MessageBox(BDOSResources.getTranslate("Localisation") + " " + BDOSResources.getTranslate("AddonLoadingSuccesfully"));
            }

            return true;

        }

        private static string AddonVersion()
        {
            try
            {
                string query = @"select ""U_Version"" from ""@BDOSAVRS"" WHERE ""Name"" = 'Localization'";

                SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                oRecordSet.DoQuery(query);

                if (!oRecordSet.EoF)
                {
                    return oRecordSet.Fields.Item("U_Version").Value;
                }

            }
            catch
            {
            }
            return "";
        }

        private static void updateAddonVersion(string version)
        {
            try
            {
                string query = @"Select * FROM ""@BDOSAVRS"" WHERE ""Name"" = 'Localization'";

                SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                oRecordSet.DoQuery(query);
                string updateQuery = "";


                if (!oRecordSet.EoF)
                {
                    updateQuery = @"UPDATE ""@BDOSAVRS""
                SET ""U_Version"" = '" + version + @"'
                WHERE ""@BDOSAVRS"".""Name"" = 'Localization'";
                }
                else
                {
                    updateQuery = @"INSERT INTO ""@BDOSAVRS"" (""Code"",""Name"",""U_Version"")
                                VALUES('1','Localization','" + version + @"')";
                }


                oRecordSet.DoQuery(updateQuery);
            }
            catch
            { }
        }

        static void uiApp_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            string errorText;

            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    //uiApp.MessageBox("A Shut Down Event has been caught" + Environment.NewLine + "Terminating 'Add Menu Item' Add On...", 1, "Ok", "", "");
                    Application.Exit();
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                    {
                        BDOSResources.initResource(Convert.ToInt32(uiApp.Language), out cultureInfo, out resourceManager, out errorText);
                        FormsB1.addMenusForAddOn();
                    }
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                    {
                        if (runAddOn())
                        {
                            uiApp.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(uiApp_ItemEvent);
                            uiApp.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(uiApp_MenuEvent);
                            uiApp.FormDataEvent += new SAPbouiCOM._IApplicationEvents_FormDataEventEventHandler(uiApp_FormDataEvent);
                            uiApp.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(uiApp_AppEvent);
                            uiApp.LayoutKeyEvent += new SAPbouiCOM._IApplicationEvents_LayoutKeyEventEventHandler(uiApp_LayoutKeyEvent);

                            //Application.Run();
                        }
                        else
                        {
                            Application.Exit();
                            return;
                        }
                    }
                    break;
            }
        }

        static void uiApp_LayoutKeyEvent(ref SAPbouiCOM.LayoutKeyInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (!localisationAddonLicensed) return;

            SAPbouiCOM.Form oForm = uiApp.Forms.ActiveForm;
            string formUID = eventInfo.FormUID;

            //----------------------------->Waybill document<-----------------------------
            if (eventInfo.BeforeAction && formUID.Contains("UDO_F_UDO_F_BDO_WBLD_D"))
            {
                if (eventInfo.EventType == SAPbouiCOM.BoEventTypes.et_PRINT_LAYOUT_KEY)
                    eventInfo.LayoutKey = oForm.DataSources.DBDataSources.Item("@BDO_WBLD").GetValue("DocEntry", 0).Trim();
            }
        }

        static void uiApp_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText;

            if (!localisationAddonLicensed) return;

            //preview  standart    
            if (pVal.BeforeAction && pVal.MenuUID == "6005")
            {
                SAPbouiCOM.Form oDocForm = uiApp.Forms.ActiveForm;
                

                if (oDocForm.TypeEx == "141" || oDocForm.TypeEx == "60092")
                {
                    CommonFunctions.fillDocRate(oDocForm, "OPCH");
                }

                if (oDocForm.TypeEx == "133" || oDocForm.TypeEx == "60091")
                {
                    CommonFunctions.fillDocRate(oDocForm, "OINV");
                }

                if (oDocForm.TypeEx == "170")
                {
                    CommonFunctions.fillDocRate(oDocForm, "ORCT");
                }

                if (oDocForm.TypeEx == "426")
                {
                    CommonFunctions.fillDocRate(oDocForm, "OVPM");
                }

                if (oDocForm.TypeEx == "65308")
                {
                    CommonFunctions.fillDocRate(oDocForm, "ODPI");
                }

                if (oDocForm.TypeEx == "65309")
                {
                    CommonFunctions.fillDocRate(oDocForm, "ODPO");
                }

            }

            //preview addon
            if (pVal.BeforeAction && pVal.MenuUID == "PreviewUDOJrE")
            {
                SAPbouiCOM.Form oDocForm = uiApp.Forms.ActiveForm;
                if (oDocForm.TypeEx == "UDO_FT_UDO_F_BDO_TAXP_D")
                {
                    BDO_ProfitTaxAccrual.uiApp_MenuEvent(ref pVal, out BubbleEvent, out errorText);
                }

                if (oDocForm.TypeEx == "UDO_FT_UDO_F_BDO_ARDPV_D")
                {
                    BDOSARDownPaymentVATAccrual.uiApp_MenuEvent(ref pVal, out BubbleEvent);
                }
            }

            //preview addon
            if (pVal.BeforeAction && pVal.MenuUID == "PreviewUDOJrE")
            {
                SAPbouiCOM.Form oDocForm = uiApp.Forms.ActiveForm;
                if (oDocForm.TypeEx == "UDO_FT_UDO_F_BDO_TAXP_D")
                {
                    BDO_ProfitTaxAccrual.uiApp_MenuEvent(ref pVal, out BubbleEvent, out errorText);
                }

            }

            if (pVal.BeforeAction && pVal.MenuUID == "1281")
            {
                SAPbouiCOM.Form oDocForm = uiApp.Forms.ActiveForm;
                if (oDocForm.TypeEx == "150")
                {
                    Items.uiApp_MenuEvent(ref pVal, out BubbleEvent);
                }
            }

            //Delete Ln Item
            if (pVal.BeforeAction && pVal.MenuUID == "5910")
            {
                SAPbouiCOM.Form oDocForm = uiApp.Forms.ActiveForm;
                if (oDocForm.TypeEx == "80030")
                {
                    CashFlowLineItem.uiApp_MenuEvent(ref pVal, out BubbleEvent, out errorText);
                }
            }

            //----------------------------->მოგების გადასახადის დაბეგვრის ობიექტების ტიპები<-----------------------------
            try
            {
                if (pVal.BeforeAction && pVal.MenuUID == "UDO_F_BDO_PTBT_D")
                {
                    errorText = null;
                    uiApp.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_UserDefinedObject, "UDO_F_BDO_PTBT_D", "");
                }
            }
            catch (Exception ex)
            {
                uiApp.MessageBox(ex.ToString(), 1, "", "");
            }

            //----------------------------->მოგების გადასახადის დაბეგვრის ობიექტები<-----------------------------
            try
            {
                if (pVal.BeforeAction && pVal.MenuUID == "UDO_F_BDO_PTBS_D")
                {
                    errorText = null;
                    uiApp.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_UserDefinedObject, "UDO_F_BDO_PTBS_D", "");
                }
            }
            catch (Exception ex)
            {
                uiApp.MessageBox(ex.ToString(), 1, "", "");
            }

            //----------------------------->მიღებული ზედნადებების ანგარიშგება<-----------------------------
            try
            {
                if (pVal.BeforeAction && pVal.MenuUID == "BDO_WBRA")
                {
                    errorText = null;
                    BDOSWaybillsAnalysisReceived.createForm(out errorText);
                }
            }
            catch (Exception ex)
            {
                uiApp.MessageBox(ex.ToString(), 1, "", "");
            }

            //----------------------------->გაცემული ზედნადებების ანგარიშგება<-----------------------------
            try
            {
                if (pVal.BeforeAction && pVal.MenuUID == "BDO_WBSA")
                {
                    errorText = null;
                    BDOSWaybillsAnalysisSent.createForm(out errorText);
                }
            }
            catch (Exception ex)
            {
                uiApp.MessageBox(ex.ToString(), 1, "", "");
            }

            //----------------------------->მიღებული ზედნადებების ჟურნალი<-----------------------------
            try
            {
                if (pVal.BeforeAction && pVal.MenuUID == "BDO_WBR")
                {
                    errorText = null;
                    SAPbouiCOM.Form noForm = null;
                    BDO_WaybillsJournalReceived.createForm(noForm, out errorText);
                }
            }
            catch (Exception ex)
            {
                uiApp.MessageBox(ex.ToString(), 1, "", "");
            }

            //----------------------------->მიღებული ფაქტურების ჟურნალი<-----------------------------
            try
            {
                if (pVal.BeforeAction && pVal.MenuUID == "BDOSTAXJ")
                {
                    errorText = null;
                    BDOSTaxJournal.createForm(out errorText);
                }
            }
            catch (Exception ex)
            {
                uiApp.MessageBox(ex.ToString(), 1, "", "");
            }

            //----------------------------->RS - ს კოდების მითითება საზომ ერთეულებზე<-----------------------------
            try
            {
                if (pVal.BeforeAction && pVal.MenuUID == "BDO_UoMRS")
                {
                    errorText = null;
                    BDO_RSUoM.createForm(out errorText);
                }
            }
            catch (Exception ex)
            {
                uiApp.MessageBox(ex.ToString(), 1, "", "");
            }

            //----------------------------->Driver Master Data<-----------------------------
            try
            {
                if (pVal.BeforeAction && pVal.MenuUID == "UDO_F_BDO_DRVS_D")
                {
                    errorText = null;
                    uiApp.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_UserDefinedObject, "UDO_F_BDO_DRVS_D", "");
                }
            }
            catch (Exception ex)
            {
                uiApp.MessageBox(ex.ToString(), 1, "", "");
            }

            //----------------------------->Vehicle Master Data<-----------------------------
            try
            {
                if (pVal.BeforeAction && pVal.MenuUID == "UDO_F_BDO_VECL_D")
                {
                    errorText = null;
                    uiApp.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_UserDefinedObject, "UDO_F_BDO_VECL_D", "");
                }
            }
            catch (Exception ex)
            {
                uiApp.MessageBox(ex.ToString(), 1, "", "");
            }

            //----------------------------->Fuel Types Master Data<-----------------------------
            try
            {
                if (pVal.BeforeAction && pVal.MenuUID == "UDO_F_BDOSFUTP_D")
                {
                    errorText = null;
                    uiApp.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_UserDefinedObject, "UDO_F_BDOSFUTP_D", "");
                }
            }
            catch (Exception ex)
            {
                uiApp.MessageBox(ex.ToString(), 1, "", "");
            }

            //----------------------------->Fuel Criteria Master Data<-----------------------------
            try
            {
                if (pVal.BeforeAction && pVal.MenuUID == "UDO_F_BDOSFUCR_D")
                {
                    errorText = null;
                    uiApp.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_UserDefinedObject, "UDO_F_BDOSFUCR_D", "");
                }
            }
            catch (Exception ex)
            {
                uiApp.MessageBox(ex.ToString(), 1, "", "");
            }

            //----------------------------->Specification of Fuel Norm Master Data<-----------------------------
            try
            {
                if (pVal.BeforeAction && pVal.MenuUID == "UDO_F_BDOSFUNR_D")
                {
                    errorText = null;
                    uiApp.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_UserDefinedObject, "UDO_F_BDOSFUNR_D", "");
                }
            }
            catch (Exception ex)
            {
                uiApp.MessageBox(ex.ToString(), 1, "", "");
            }

            //----------------------------->Fuel Consumption Act Document<-----------------------------
            try
            {
                if (pVal.BeforeAction && pVal.MenuUID == "UDO_F_BDOSFUCN_D")
                {
                    errorText = null;
                    uiApp.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_UserDefinedObject, "UDO_F_BDOSFUCN_D", "");
                }
            }
            catch (Exception ex)
            {
                uiApp.MessageBox(ex.ToString(), 1, "", "");
            }

            //----------------------------->Waybill document<-----------------------------
            try
            {
                if (pVal.BeforeAction && pVal.MenuUID == "UDO_F_BDO_WBLD_D")
                {
                    errorText = null;
                    uiApp.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_UserDefinedObject, "UDO_F_BDO_WBLD_D", "");
                }
            }
            catch (Exception ex)
            {
                uiApp.MessageBox(ex.ToString(), 1, "", "");
            }

            //----------------------------->Tax Invoice Sent<-----------------------------
            try
            {
                if (pVal.BeforeAction && pVal.MenuUID == "UDO_F_BDO_TAXS_D")
                {
                    errorText = null;
                    uiApp.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_UserDefinedObject, "UDO_F_BDO_TAXS_D", "");
                }
            }
            catch (Exception ex)
            {
                uiApp.MessageBox(ex.ToString(), 1, "", "");
            }


            //----------------------------->A/R Down Payment VAT Accrual<-----------------------------
            try
            {
                if (pVal.BeforeAction && pVal.MenuUID == "UDO_F_BDO_ARDPV_D")
                {
                    errorText = null;
                    uiApp.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_UserDefinedObject, "UDO_F_BDO_ARDPV_D", "");
                }
            }
            catch (Exception ex)
            {
                uiApp.MessageBox(ex.ToString(), 1, "", "");
            }

            //----------------------------->Profit Tax Accural<-----------------------------
            try
            {
                if (pVal.BeforeAction && pVal.MenuUID == "UDO_F_BDO_TAXP_D")
                {
                    errorText = null;
                    uiApp.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_UserDefinedObject, "UDO_F_BDO_TAXP_D", "");
                }
            }
            catch (Exception ex)
            {
                uiApp.MessageBox(ex.ToString(), 1, "", "");
            }

            //----------------------------->Tax Invoice Received<-----------------------------
            try
            {
                if (pVal.BeforeAction && pVal.MenuUID == "UDO_F_BDO_TAXR_D")
                {
                    errorText = null;
                    uiApp.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_UserDefinedObject, "UDO_F_BDO_TAXR_D", "");
                }
            }
            catch (Exception ex)
            {
                uiApp.MessageBox(ex.ToString(), 1, "", "");
            }

            //----------------------------->Fixed Asset Transfer<-----------------------------
            try
            {
                if (pVal.BeforeAction && pVal.MenuUID == "UDO_F_BDOSFASTRD_D")
                {
                    errorText = null;
                    uiApp.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_UserDefinedObject, "UDO_F_BDOSFASTRD_D", "");
                }
            }
            catch (Exception ex)
            {
                uiApp.MessageBox(ex.ToString(), 1, "", "");
            }

            //----------------------------->Fixed Asset Transfer Add/Delete Row<-----------------------------
            try
            {
                if (!pVal.BeforeAction && (pVal.MenuUID == "BDOSDelRow" || pVal.MenuUID == "BDOSAddRow"))
                {
                    SAPbouiCOM.Form oDocForm = uiApp.Forms.ActiveForm;

                    if (oDocForm.TypeEx == "UDO_FT_UDO_F_BDOSFASTRD_D")
                    {
                        BDOSFixedAssetTransfer.uiApp_MenuEvent(ref pVal, out BubbleEvent, out errorText);
                    }
                }
            }
            catch (Exception ex)
            {
                uiApp.MessageBox(ex.ToString(), 1, "", "");
            }

            //----------------------------->Depreciation Accrual<-----------------------------
            try
            {
                if (pVal.BeforeAction && pVal.MenuUID == "UDO_F_BDOSDEPACR_D")
                {
                    errorText = null;
                    uiApp.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_UserDefinedObject, "UDO_F_BDOSDEPACR_D", "");
                }
            }
            catch (Exception ex)
            {
                uiApp.MessageBox(ex.ToString(), 1, "", "");
            }

            //----------------------------->Depreciation Accrual Add/Delete Row<-----------------------------
            try
            {
                if (!pVal.BeforeAction && (pVal.MenuUID == "BDOSDelRow" || pVal.MenuUID == "BDOSAddRow"))
                {
                    SAPbouiCOM.Form oDocForm = uiApp.Forms.ActiveForm;

                    if (oDocForm.TypeEx == "UDO_FT_UDO_F_BDOSDEPACR_D")
                    {
                        BDOSDepreciationAccrualDocument.uiApp_MenuEvent(ref pVal, out BubbleEvent);
                    }
                }
            }
            catch (Exception ex)
            {
                uiApp.MessageBox(ex.ToString(), 1, "", "");
            }

            //----------------------------->Internet Banking<-----------------------------
            try
            {
                if (!pVal.BeforeAction && pVal.MenuUID == "BDOSInternetBankingForm")
                {
                    errorText = null;
                    BDOSInternetBanking.createForm(out errorText);
                }
            }
            catch (Exception ex)
            {
                uiApp.MessageBox(ex.ToString(), 1, "", "");
            }

            //----------------------------->Delete UDF<-----------------------------
            try
            {
                if (!pVal.BeforeAction && pVal.MenuUID == "BDOSDeleteUDFForm")
                {
                    errorText = null;
                    BDOSDeleteUDF.createForm(out errorText);
                }
            }
            catch (Exception ex)
            {
                uiApp.MessageBox(ex.ToString(), 1, "", "");
            }

            //----------------------------->Outgoing Payments wizzard<-----------------------------
            try
            {
                if (pVal.BeforeAction && pVal.MenuUID == "BDOSSOPWizzForm")
                {
                    errorText = null;
                    BDOSOutgoingPaymentsWizard.createForm(out errorText);
                }
            }
            catch (Exception ex)
            {
                uiApp.MessageBox(ex.ToString(), 1, "", "");
            }

            //----------------------------->Stock Transfer Wizard-----------------------------
            try
            {
                if (pVal.BeforeAction && pVal.MenuUID == "BDOSSTTRWZ")
                {
                    errorText = null;
                    BDOSStockTransferWizard.createForm(out errorText);
                }
            }
            catch (Exception ex)
            {
                uiApp.MessageBox(ex.ToString(), 1, "", "");
            }

            //----------------------------->Depreciation Accruing wizzard<-----------------------------
            try
            {
                if (pVal.BeforeAction && pVal.MenuUID == "BDOSDepAccrForm")
                {
                    BDOSDepreciationAccrualWizard.createForm();
                }
            }
            catch (Exception ex)
            {
                uiApp.MessageBox(ex.ToString(), 1, "", "");
            }

            //----------------------------->AR Down Payment VAT Accrual Wizard<-----------------------------
            try
            {
                if (pVal.BeforeAction && pVal.MenuUID == "BDOSVAWizzForm")
                {
                    BDOSVATAccrualWizard.createForm();
                }
            }
            catch (Exception ex)
            {
                uiApp.MessageBox(ex.ToString(), 1, "", "");
            }

            //----------------------------->Fuel Write-Off Wizard<-----------------------------
            try
            {
                if (pVal.BeforeAction && pVal.MenuUID == "BDOSFuelWOForm")
                {
                    BDOSFuelWriteOffWizard.createForm();
                }
            }
            catch (Exception ex)
            {
                uiApp.MessageBox(ex.ToString(), 1, "", "");
            }

            //----------------------------->VAT Reconcilation Wizard<-----------------------------
            try
            {
                if (pVal.BeforeAction && pVal.MenuUID == "BDOSReconWizz")
                {
                    BDOSVATReconcilationWizard.createForm();
                }
            }
            catch (Exception ex)
            {
                uiApp.MessageBox(ex.ToString(), 1, "", "");
            }

            //----------------------------->Fuel Transfer Wizard<-----------------------------
            try
            {
                if (pVal.BeforeAction && pVal.MenuUID == "BDOSFuelTransferWizard")
                {
                    errorText = null;
                    BDOSFuelTransferWizard.createForm();
                }
            }
            catch (Exception ex)
            {
                uiApp.MessageBox(ex.ToString(), 1, "", "");
            }

            //----------------------------->Internet Banking Integration Services Rules<-----------------------------
            try
            {
                if (pVal.BeforeAction && pVal.MenuUID == "UDO_F_BDOSINTR_D")
                {
                    errorText = null;
                    uiApp.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_UserDefinedObject, "UDO_F_BDOSINTR_D", "");
                }
            }
            catch (Exception ex)
            {
                uiApp.MessageBox(ex.ToString(), 1, "", "");
            }

            //----------------------------->Item Categories<-----------------------------
            try
            {
                if (pVal.BeforeAction && pVal.MenuUID == "UDO_F_BDOSITMCTG_D")
                {
                    errorText = null;
                    uiApp.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_UserDefinedObject, "UDO_F_BDOSITMCTG_D", "");
                }
            }
            catch (Exception ex)
            {
                uiApp.MessageBox(ex.ToString(), 1, "", "");
            }

            //----------------------------->Credit Line Master Data<-----------------------------
            try
            {
                if (pVal.BeforeAction && pVal.MenuUID == "UDO_F_BDOSCRLN_D")
                {
                    errorText = null;
                    uiApp.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_UserDefinedObject, "UDO_F_BDOSCRLN_D", "");
                }
            }
            catch (Exception ex)
            {
                uiApp.MessageBox(ex.ToString(), 1, "", "");
            }

            //----------------------------->Interest Accrual Document<-----------------------------
            try
            {
                if (pVal.BeforeAction && pVal.MenuUID == "UDO_F_BDOSINAC_D")
                {
                    errorText = null;
                    uiApp.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_UserDefinedObject, "UDO_F_BDOSINAC_D", "");
                }
            }
            catch (Exception ex)
            {
                uiApp.MessageBox(ex.ToString(), 1, "", "");
            }

            //----------------------------->Interest Accrual Wizard<-----------------------------
            try
            {
                if (pVal.BeforeAction && pVal.MenuUID == "BDOSInterestAccrualWizard")
                {
                    errorText = null;
                    BDOSInterestAccrualWizard.createForm();
                }
            }
            catch (Exception ex)
            {
                uiApp.MessageBox(ex.ToString(), 1, "", "");
            }

            //----------------------------->Cancel<-----------------------------
            if (pVal.MenuUID == "1284")
            {
                try
                {
                    if (pVal.BeforeAction)
                    {
                        SAPbouiCOM.Form oForm = uiApp.Forms.ActiveForm;

                        //----------------------------->A/R Invoice<-----------------------------
                        if (oForm.TypeEx == "133")
                        {
                            cancellationTrans = true;
                            canceledDocEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("OINV").GetValue("DocEntry", 0));
                        }
                        //----------------------------->Inventory Transfer<-----------------------------
                        else if (oForm.TypeEx == "940")
                        {

                            StockTransfer.uiApp_MenuEvent(ref pVal, out BubbleEvent, out errorText);

                            cancellationTrans = true;
                            canceledDocEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("OWTR").GetValue("DocEntry", 0));
                        }
                        //----------------------------->A/R Credit Note<-----------------------------
                        else if (oForm.TypeEx == "179")
                        {
                            cancellationTrans = true;
                            canceledDocEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("ORIN").GetValue("DocEntry", 0));
                        }

                        //----------------------------->A/R Correction Invoice<-----------------------------
                        else if (oForm.TypeEx == "70008")
                        {
                            cancellationTrans = true;
                            canceledDocEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("OCSI").GetValue("DocEntry", 0));
                        }
                        //----------------------------->Depreciation<-----------------------------
                        else if (oForm.TypeEx == "UDO_FT_UDO_F_BDOSDEPACR_D")
                        {
                            cancellationTrans = true;
                            canceledDocEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("@BDOSDEPACR").GetValue("DocEntry", 0));
                        }
                        //----------------------------->Profit Tax Accrual<-----------------------------
                        else if (oForm.TypeEx == "UDO_FT_UDO_F_BDO_TAXP_D")
                        {
                            cancellationTrans = true;
                            canceledDocEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("@BDO_TAXP").GetValue("DocEntry", 0));
                        }
                        //----------------------------->Fixes Asset Transfer<-----------------------------
                        else if (oForm.TypeEx == "UDO_FT_UDO_F_BDOSFASTRD_D")
                        {
                            cancellationTrans = true;
                            canceledDocEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("@BDOSFASTRD").GetValue("DocEntry", 0));
                        }
                        //----------------------------->Outgoing Payments<-----------------------------
                        else if (oForm.TypeEx == "426")
                        {
                            cancellationTrans = true;
                            canceledDocEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("OVPM").GetValue("DocEntry", 0));
                        }
                        //----------------------------->Incoming Paymentss<-----------------------------
                        else if (oForm.TypeEx == "170")
                        {
                            cancellationTrans = true;
                            canceledDocEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("ORCT").GetValue("DocEntry", 0));
                        }
                        //----------------------------->A/P Invoice<-----------------------------
                        if (oForm.TypeEx == "141")
                        {
                            cancellationTrans = true;
                            canceledDocEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("OPCH").GetValue("DocEntry", 0));
                        }
                        //----------------------------->A/P Reserve Invoice<-----------------------------
                        if (oForm.TypeEx == "60092")
                        {
                            cancellationTrans = true;
                            canceledDocEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("OPCH").GetValue("DocEntry", 0));
                        }
                        //----------------------------->Retirement<-----------------------------
                        if (oForm.TypeEx == "1470000014")
                        {
                            cancellationTrans = true;
                            canceledDocEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("ORTI").GetValue("DocEntry", 0));
                        }
                        //----------------------------->A/P Credit Memo<-----------------------------
                        if (oForm.TypeEx == "181")
                        {
                            cancellationTrans = true;
                            canceledDocEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("ORPC").GetValue("DocEntry", 0));
                        }
                        //----------------------------->Tax Invoice Received<-----------------------------
                        if (oForm.TypeEx == "UDO_FT_UDO_F_BDO_TAXR_D")
                        {
                            cancellationTrans = true;
                            canceledDocEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("@BDO_TAXR").GetValue("DocEntry", 0));
                        }
                        //----------------------------->Tax Invoice Sent<-----------------------------
                        if (oForm.TypeEx == "UDO_FT_UDO_F_BDO_TAXS_D")
                        {
                            cancellationTrans = true;
                            canceledDocEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("@BDO_TAXS").GetValue("DocEntry", 0));
                        }
                        //----------------------------->AR Down Payment VAT Accrual<-----------------------------
                        else if (oForm.TypeEx == "UDO_FT_UDO_F_BDO_ARDPV_D")
                        {
                            cancellationTrans = true;
                            canceledDocEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("@BDOSARDV").GetValue("DocEntry", 0));
                        }
                        //----------------------------->Interest Accrual Document<-----------------------------
                        else if (oForm.TypeEx == "UDO_FT_UDO_F_BDOSINAC_D")
                        {
                            cancellationTrans = true;
                            canceledDocEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("@BDOSINAC").GetValue("DocEntry", 0));
                        }
                        //----------------------------->>Journal Entry<<-----------------------------
                        else if (oForm.TypeEx == "392")
                        {
                            cancellationTrans = true;
                            canceledDocEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("OJDT").GetValue("TransId", 0));

                            JournalEntry.uiApp_MenuEvent(ref pVal, out BubbleEvent);
                        }
                    }
                    else
                    {
                        if (!cancellationTrans)
                            cancellationDoc = true;
                        cancellationTrans = false;
                    }
                }
                catch (Exception ex)
                {
                    uiApp.StatusBar.SetSystemMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short);
                }
            }

            //----------------------------->Remove<-----------------------------
            if (pVal.MenuUID == "1283")
            {
                if (pVal.BeforeAction)
                {
                    SAPbouiCOM.Form oForm = uiApp.Forms.ActiveForm;

                    //----------------------------->Profit Tax Base Type<-----------------------------
                    if (oForm.TypeEx == "UDO_F_BDO_PTBT_D")
                    {
                        removeRecordTrans = true;
                    }
                    //----------------------------->Profit Tax Base<-----------------------------
                    if (oForm.TypeEx == "UDO_F_BDO_PTBS_D")
                    {
                        removeRecordTrans = true;
                    }
                    //----------------------------->Vehicle<-----------------------------
                    if (oForm.TypeEx == "UDO_FT_UDO_F_BDO_VECL_D")
                    {
                        removeRecordTrans = true;
                    }
                    //----------------------------->Drivers<-----------------------------
                    if (oForm.TypeEx == "UDO_FT_UDO_F_BDO_DRVS_D")
                    {
                        removeRecordTrans = true;
                    }
                }
            }

            //----------------------------->Remove Line<-----------------------------
            if (pVal.MenuUID == "UDO_F_BDO_TAXP_D_Remove_Line" & !pVal.BeforeAction)
            {
                removeLineTrans = true;
            }

            //----------------------------->Duplicate<-----------------------------
            if (pVal.MenuUID == "1287")
            {
                try
                {
                    if (pVal.BeforeAction)
                    {
                        SAPbouiCOM.Form oForm = uiApp.Forms.ActiveForm;
                    }
                    else
                    {
                        SAPbouiCOM.Form oForm = uiApp.Forms.ActiveForm;
                        //----------------------------->A/R Invoice<-----------------------------
                        if (oForm.TypeEx == "133")
                        {
                            ARInvoice.formDataLoad(oForm, out errorText);
                        }
                        //if (oForm.TypeEx == "70001")
                        //{
                        //    StockRevaluation.formDataLoad(oForm);
                        //}

                        //if (oForm.TypeEx == "992")
                        //{
                        //    LandedCosts.formDataLoad(oForm);
                        //}

                        //----------------------------->Asset Master Data<-----------------------------
                        if (oForm.TypeEx == "1473000075")
                        {
                            FixedAsset.formDataLoad(oForm);
                        }

                        //----------------------------->Depreciation Accrual Document<-----------------------------
                        if (oForm.TypeEx == "UDO_FT_UDO_F_BDOSDEPACR_D")
                        {
                            BDOSDepreciationAccrualDocument.formDataLoad(oForm);
                        }

                        //----------------------------->A/R Reserve Invoice<-----------------------------
                        if (oForm.TypeEx == "60091")
                        {
                            ARReserveInvoice.formDataLoad(oForm, out errorText);
                        }

                        //----------------------------->Inventory Transfer<-----------------------------
                        else if (oForm.TypeEx == "940")
                        {
                            StockTransfer.formDataLoad(oForm, out errorText);
                        }

                        //----------------------------->A/R Credit Note<-----------------------------
                        else if (oForm.TypeEx == "179")
                        {
                            ARCreditNote.formDataLoad(oForm, out errorText);
                        }

                        //----------------------------->A/R Correction Invoice<-----------------------------
                        else if (oForm.TypeEx == "70008")
                        {
                            ArCorrectionInvoice.FormDataLoad(oForm, out errorText);
                        }

                        //----------------------------->A/P Invoice<-----------------------------
                        else if (oForm.TypeEx == "141")
                        {
                            APInvoice.formDataLoad(oForm, out errorText);
                            APInvoice.setVisibleFormItems(oForm, out errorText);
                        }

                        //----------------------------->Goods Receipt PO<-----------------------------
                        else if (oForm.TypeEx == "143")
                        {
                            GoodsReceiptPO.formDataLoad(oForm, out errorText);
                            GoodsReceiptPO.setVisibleFormItems(oForm, out errorText);
                        }

                        //----------------------------->A/P Credit Memo<-----------------------------
                        else if (oForm.TypeEx == "181")
                        {
                            APCreditMemo.formDataLoad(oForm, out errorText);
                            APCreditMemo.setVisibleFormItems(oForm, out errorText);
                        }

                        //----------------------------->Outgoing Payments<-----------------------------
                        else if (oForm.TypeEx == "426")
                        {
                            OutgoingPayment.formDataLoad(oForm);
                        }

                        //----------------------------->Blanket agreement<-----------------------------
                        else if (oForm.TypeEx == "1250000100")
                        {
                            BlanketAgreement.formDataLoad(oForm, out errorText);
                        }

                        //----------------------------->Goods Issue<-----------------------------
                        else if (oForm.TypeEx == "720")
                        {
                            GoodsIssue.formDataLoad(oForm, out errorText);
                        }

                        //----------------------------->Profit Tax Accural<-----------------------------
                        else if (oForm.TypeEx == "UDO_FT_UDO_F_BDO_TAXP_D")
                        {
                            BDO_ProfitTaxAccrual.formDataLoad(oForm);
                        }

                        //----------------------------->A/P Down Payment Request<-----------------------------
                        else if (oForm.TypeEx == "65309")
                        {
                            APDownPaymentRequest.formDataLoad(oForm, out errorText);
                        }

                        //----------------------------->A/P Down Payment Invoice<-----------------------------
                        else if (oForm.TypeEx == "65301")
                        {
                            APDownPaymentInvoice.formDataLoad(oForm, out errorText);
                        }

                        //----------------------------->A/R Down Payment Request<-----------------------------
                        //else if (oForm.TypeEx == "65308")
                        //{
                        //    ARDownPaymentRequest.formDataLoad(oForm, out errorText);
                        //}
                        //----------------------------->A/R Down Payment VAT Accrual<-----------------------------
                        else if (oForm.TypeEx == "UDO_FT_UDO_F_BDO_ARDPV_D")
                        {
                            BDOSARDownPaymentVATAccrual.formDataLoad(oForm);
                        }

                        //----------------------------->Retirement<-----------------------------
                        else if (oForm.TypeEx == "1470000014")
                        {
                            Retirement.formDataLoad(oForm, out errorText);
                        }

                        //----------------------------->Warehouses<-----------------------------
                        else if (oForm.TypeEx == "62")
                        {
                            Warehouses.formDataLoad(oForm, out errorText);
                        }
                    }
                }
                catch (Exception ex)
                {
                    uiApp.StatusBar.SetSystemMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short);
                }
            }
            //----------------------------->Add<-----------------------------
            if (pVal.MenuUID == "1282")
            {
                try
                {
                    if (pVal.BeforeAction)
                    {
                        SAPbouiCOM.Form oForm = uiApp.Forms.ActiveForm;

                        //----------------------------->Waybill document<-----------------------------
                        if (oForm.TypeEx == "UDO_FT_UDO_F_BDO_WBLD_D")
                        {
                            uiApp.MessageBox(BDOSResources.getTranslate("CreateWaybillAllowedBasedOnlyOtherDocument"));
                            BubbleEvent = false;
                        }
                    }
                    else
                    {
                        SAPbouiCOM.Form oForm = uiApp.Forms.ActiveForm;

                        //----------------------------->A/R Invoice<-----------------------------
                        if (oForm.TypeEx == "133")
                        {
                            ARInvoice.formDataLoad(oForm, out errorText);
                        }

                        //if (oForm.TypeEx == "70001")
                        //{
                        //    StockRevaluation.formDataLoad(oForm);
                        //}

                        //if (oForm.TypeEx == "992")
                        //{
                        //    LandedCosts.formDataLoad(oForm);
                        //}

                        //----------------------------->Depreciation Accrual Document<-----------------------------
                        if (oForm.TypeEx == "UDO_FT_UDO_F_BDOSDEPACR_D")
                        {
                            BDOSDepreciationAccrualDocument.formDataLoad(oForm);
                        }

                        //----------------------------->Delivery<-----------------------------
                        if (oForm.TypeEx == "140")
                        {
                            Delivery.formDataLoad(oForm, out errorText);
                        }

                        //----------------------------->Asset Master Data<-----------------------------
                        if (oForm.TypeEx == "1473000075")
                        {
                            FixedAsset.formDataLoad(oForm);
                        }

                        //----------------------------->A/R Reserve Invoice<-----------------------------
                        if (oForm.TypeEx == "60091")
                        {
                            ARReserveInvoice.formDataLoad(oForm, out errorText);
                        }

                        //----------------------------->Inventory Transfer<-----------------------------
                        else if (oForm.TypeEx == "940")
                        {
                            StockTransfer.formDataLoad(oForm, out errorText);
                        }

                        //----------------------------->A/R Credit Note<-----------------------------
                        else if (oForm.TypeEx == "179")
                        {
                            ARCreditNote.formDataLoad(oForm, out errorText);
                        }

                        //----------------------------->A/R Correction Invoice<-----------------------------
                        else if (oForm.TypeEx == "70008")
                        {
                            ArCorrectionInvoice.FormDataLoad(oForm, out errorText);
                        }

                        //----------------------------->A/P Invoice<-----------------------------
                        else if (oForm.TypeEx == "141")
                        {
                            APInvoice.formDataLoad(oForm, out errorText);
                        }

                        //----------------------------->Goods Receipt PO<-----------------------------
                        else if (oForm.TypeEx == "143")
                        {
                            GoodsReceiptPO.formDataLoad(oForm, out errorText);
                        }

                        //----------------------------->A/P Credit Memo<-----------------------------
                        else if (oForm.TypeEx == "181")
                        {
                            APCreditMemo.formDataLoad(oForm, out errorText);
                        }

                        //----------------------------->Tax Invoice Sent<-----------------------------
                        else if (oForm.TypeEx == "UDO_FT_UDO_F_BDO_TAXS_D")
                        {
                            BDO_TaxInvoiceSent.formDataLoad(oForm, out errorText);
                        }
                        //----------------------------->Profit Tax Accural<-----------------------------
                        else if (oForm.TypeEx == "UDO_FT_UDO_F_BDO_TAXP_D")
                        {
                            BDO_ProfitTaxAccrual.formDataLoad(oForm);
                        }

                        //----------------------------->Outgoing Payments<-----------------------------
                        else if (oForm.TypeEx == "426")
                        {
                            OutgoingPayment.formDataLoad(oForm);
                        }

                        //----------------------------->Blanket agreement<-----------------------------
                        else if (oForm.TypeEx == "1250000100")
                        {
                            BlanketAgreement.formDataLoad(oForm, out errorText);
                        }

                        //----------------------------->Goods Issue<-----------------------------
                        else if (oForm.TypeEx == "720")
                        {
                            GoodsIssue.formDataLoad(oForm, out errorText);
                        }

                        //----------------------------->A/P Down Payment Request<-----------------------------
                        else if (oForm.TypeEx == "65309")
                        {
                            APDownPaymentRequest.formDataLoad(oForm, out errorText);
                        }

                        //----------------------------->A/P Down Payment Invoice<-----------------------------
                        else if (oForm.TypeEx == "65301")
                        {
                            APDownPaymentInvoice.formDataLoad(oForm, out errorText);
                        }

                        //----------------------------->A/R Down Payment Request<-----------------------------
                        //else if (oForm.TypeEx == "65308")
                        //{
                        //    ARDownPaymentRequest.formDataLoad(oForm, out errorText);
                        //}
                        //----------------------------->A/R Down Payment VAT Accrual<-----------------------------
                        else if (oForm.TypeEx == "UDO_FT_UDO_F_BDO_ARDPV_D")
                        {
                            BDOSARDownPaymentVATAccrual.formDataLoad(oForm);
                        }

                        //----------------------------->Retirement<-----------------------------
                        else if (oForm.TypeEx == "1470000014")
                        {
                            Retirement.formDataLoad(oForm, out errorText);
                        }

                        //-----------------------------Warehouses<-----------------------------
                        else if (oForm.TypeEx == "62")
                        {
                            Warehouses.formDataLoad(oForm, out errorText);
                        }

                        //-----------------------------Landed Cost<-----------------------------
                        else if (oForm.TypeEx == "992")
                        {
                            LandedCosts.formDataLoad(oForm);
                        }

                        //-----------------------------Stock Revaluation<-----------------------------
                        else if (oForm.TypeEx == "70001")
                        {
                            StockRevaluation.formDataLoad(oForm);
                        }
                    }
                }
                catch (Exception ex)
                {
                    uiApp.StatusBar.SetSystemMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short);
                }
            }

            //----------------------------->Find<-----------------------------
            if (pVal.MenuUID == "1281")
            {
                try
                {
                    SAPbouiCOM.Form oForm = uiApp.Forms.ActiveForm;

                    if (!pVal.BeforeAction)
                    {
                        //----------------------------->A/R Down Payment VAT Accrual<-----------------------------
                        if (oForm.TypeEx == "UDO_FT_UDO_F_BDO_ARDPV_D")
                        {
                            BDOSARDownPaymentVATAccrual.formDataLoad(oForm);
                        }

                        //----------------------------->Depreciation Accrual Document<-----------------------------
                        else if (oForm.TypeEx == "UDO_FT_UDO_F_BDOSDEPACR_D")
                        {
                            BDOSDepreciationAccrualDocument.formDataLoad(oForm);
                        }
                    }
                }
                catch (Exception ex)
                {
                    uiApp.StatusBar.SetSystemMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short);
                }
            }

            try
            {
                if (!pVal.BeforeAction && pVal.MenuUID == "BDO_WBS")
                {
                    BDO_WaybillsJournalSent.createForm(out errorText);
                }
            }
            catch (Exception ex)
            {
                uiApp.MessageBox(ex.ToString(), 1, "", "");
            }
        }

        static void uiApp_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (!localisationAddonLicensed) return;

            try
            {
                //----------------------------->Journal Entry<-----------------------------
                if (BusinessObjectInfo.Type == "30")
                {
                    JournalEntry.uiApp_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                }

                //----------------------------->Cahrt Of Accounts<-----------------------------
                else if (BusinessObjectInfo.Type == "1")
                {
                    ChartOfAccounts.uiApp_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                }

                //----------------------------->Outgoing Payments<-----------------------------
                else if (BusinessObjectInfo.Type == "46")
                {
                    OutgoingPayment.uiApp_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                }

                //----------------------------->Incoming Payments<-----------------------------
                else if (BusinessObjectInfo.Type == "24")
                {
                    IncomingPayment.uiApp_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                }

                //----------------------------->Goods Issue<-----------------------------
                else if (BusinessObjectInfo.Type == "60")
                {
                    GoodsIssue.uiApp_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                }

                //----------------------------->FA Capitalization<-----------------------------
                else if (BusinessObjectInfo.Type == "1470000049")
                {
                    Capitalization.uiApp_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                }

                //----------------------------->Retirement<-----------------------------
                else if (BusinessObjectInfo.Type == "1470000094")
                {
                    Retirement.uiApp_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                }

                //----------------------------->House Bank Account<-----------------------------
                else if (BusinessObjectInfo.Type == "231")
                {
                    HouseBankAccounts.uiApp_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                }

                //----------------------------->Landed costs<-----------------------------
                else if (BusinessObjectInfo.Type == "69")
                {
                    LandedCosts.uiApp_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                }

                //----------------------------->Business Partner Master Data<-----------------------------
                else if (BusinessObjectInfo.Type == "2")
                {
                    BusinessPartners.uiApp_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                }

                //----------------------------->Item Master Data<-----------------------------
                else if (BusinessObjectInfo.Type == "4" && BusinessObjectInfo.FormTypeEx == "150")
                {
                    Items.uiApp_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                }

                //----------------------------->Items Groups<-----------------------------
                else if (BusinessObjectInfo.Type == "52" && BusinessObjectInfo.FormTypeEx == "63")
                {
                    ItemGroup.uiApp_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                }

                //----------------------------->A/P Invoice<-----------------------------
                else if (BusinessObjectInfo.Type == "18" && BusinessObjectInfo.FormTypeEx == "141")
                {
                    APInvoice.uiApp_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                }

                //----------------------------->Goods Receipt PO<-----------------------------
                else if (BusinessObjectInfo.Type == "20" && BusinessObjectInfo.FormTypeEx == "143")
                {
                    GoodsReceiptPO.uiApp_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                }

                //----------------------------->A/P Reserve Invoice<-----------------------------
                else if (BusinessObjectInfo.Type == "18" && BusinessObjectInfo.FormTypeEx == "60092")
                {
                    APReserveInvoice.uiApp_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                }

                //----------------------------->A/P Correction Invoice<-----------------------------
                else if (BusinessObjectInfo.Type == "164" || BusinessObjectInfo.Type == "163")
                {
                    APCorrectionInvoice.uiApp_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                }

                //----------------------------->A/P Credit Memo<-----------------------------
                else if (BusinessObjectInfo.Type == "19")
                {
                    APCreditMemo.uiApp_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                }

                //----------------------------->A/R Invoice<-----------------------------
                else if (BusinessObjectInfo.Type == "13" && BusinessObjectInfo.FormTypeEx == "133")
                {
                    ARInvoice.uiApp_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                }

                else if (BusinessObjectInfo.FormTypeEx == "70001")
                {
                    StockRevaluation.uiApp_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                }

                //----------------------------->Asset Master Data<-----------------------------
                else if (BusinessObjectInfo.Type == "4" && BusinessObjectInfo.FormTypeEx == "1473000075")
                {
                    FixedAsset.uiApp_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                }

                //----------------------------->Asset Class<-----------------------------
                else if (BusinessObjectInfo.Type == "1470000032") // && BusinessObjectInfo.FormTypeEx == "1472000006"
                {
                    AssetClass.uiApp_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                }

                //----------------------------->A/R Reserve Invoice<-----------------------------
                else if (BusinessObjectInfo.Type == "13" && BusinessObjectInfo.FormTypeEx == "60091")
                {
                    ARReserveInvoice.uiApp_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                }

                //----------------------------->A/R CreditNote<-----------------------------
                else if (BusinessObjectInfo.Type == "14")
                {
                    ARCreditNote.uiApp_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                }

                //----------------------------->Inventory Transfer<-----------------------------
                else if (BusinessObjectInfo.Type == "67")
                {
                    StockTransfer.uiApp_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                }


                //----------------------------->Inventory Transfer Request<-----------------------------
                else if (BusinessObjectInfo.Type == "1250000001")
                {
                    StockTransferRequest.uiApp_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                }

                //----------------------------->Delivery<-----------------------------
                else if (BusinessObjectInfo.Type == "15")
                {
                    Delivery.uiApp_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                }

                //----------------------------->Waybill document<-----------------------------
                else if (BusinessObjectInfo.Type == "UDO_F_BDO_WBLD_D")
                {
                    BDO_Waybills.uiApp_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                }

                //----------------------------->Tax Invoice Received<-----------------------------
                else if (BusinessObjectInfo.Type == "UDO_F_BDO_TAXR_D")
                {
                    BDO_TaxInvoiceReceived.uiApp_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                }

                //----------------------------->Tax Invoice Sent<-----------------------------
                else if (BusinessObjectInfo.Type == "UDO_F_BDO_TAXS_D")
                {
                    BDO_TaxInvoiceSent.uiApp_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                }

                //----------------------------->Company Details<-----------------------------
                else if (BusinessObjectInfo.Type == "39")
                {
                    CompanyDetails.uiApp_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                }

                //----------------------------->Profit Tax Accural<-----------------------------
                else if (BusinessObjectInfo.Type == "UDO_F_BDO_TAXP_D")
                {
                    BDO_ProfitTaxAccrual.uiApp_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                }

                //----------------------------->Fixed Asset Transfer Document<-----------------------------
                else if (BusinessObjectInfo.Type == "UDO_F_BDOSFASTRD_D")
                {
                    BDOSFixedAssetTransfer.uiApp_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                }

                //----------------------------->Depreciation Accrual Document<-----------------------------
                else if (BusinessObjectInfo.Type == "UDO_F_BDOSDEPACR_D")
                {
                    BDOSDepreciationAccrualDocument.uiApp_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                }

                ////----------------------------->Fuel Consumption Document<-----------------------------
                //else if (BusinessObjectInfo.Type == "UDO_F_BDOSFUECON_D")
                //{
                //    BDOSFuelConsumption.uiApp_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                //}

                //----------------------------->A/R Down Payment Invoice<-----------------------------
                else if (BusinessObjectInfo.Type == "203" && BusinessObjectInfo.FormTypeEx == "65300")
                {
                    ARDownPaymentInvoice.uiApp_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                }

                //----------------------------->A/R Down Payment Request<-----------------------------
                else if (BusinessObjectInfo.Type == "203" && BusinessObjectInfo.FormTypeEx == "65308")
                {
                    ARDownPaymentRequest.uiApp_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                }

                //----------------------------->A/R Down Payment VAT Accrual<-----------------------------
                else if (BusinessObjectInfo.Type == "UDO_F_BDO_ARDPV_D")
                {
                    BDOSARDownPaymentVATAccrual.uiApp_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                }

                //----------------------------->A/P DownPayment Invoice<-----------------------------
                else if (BusinessObjectInfo.Type == "204" && BusinessObjectInfo.FormTypeEx == "65301")
                {
                    APDownPaymentInvoice.uiApp_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                }

                //----------------------------->A/P Down Payment Request<-----------------------------
                else if (BusinessObjectInfo.Type == "204" && BusinessObjectInfo.FormTypeEx == "65309")
                {
                    APDownPaymentRequest.uiApp_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                }

                //----------------------------->Vehicles<----------------------------
                else if (BusinessObjectInfo.Type == "UDO_F_BDO_VECL_D")
                {
                    BDO_Vehicles.uiApp_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                }

                //----------------------------->Drivers<----------------------------
                else if (BusinessObjectInfo.Type == "UDO_F_BDO_DRVS_D")
                {
                    BDO_Drivers.uiApp_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                }

                //----------------------------->Fuel Types<----------------------------
                else if (BusinessObjectInfo.Type == "UDO_F_BDOSFUTP_D")
                {
                    BDOSFuelTypes.uiApp_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                }

                //----------------------------->Specification of Fuel Norm<----------------------------
                else if (BusinessObjectInfo.Type == "UDO_F_BDOSFUNR_D")
                {
                    BDOSFuelNormSpecification.uiApp_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                }

                //----------------------------->Fuel Criteria Master Data<----------------------------
                else if (BusinessObjectInfo.Type == "UDO_F_BDOSFUCR_D")
                {
                    BDOSFuelCriteria.uiApp_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                }

                //----------------------------->Fuel Consumption Act Document<----------------------------
                else if (BusinessObjectInfo.Type == "UDO_F_BDOSFUCN_D")
                {
                    BDOSFuelConsumptionAct.uiApp_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                }

                //----------------------------->Blanket Agreement<-----------------------------
                else if (BusinessObjectInfo.Type == "1250000025")
                {
                    BlanketAgreement.uiApp_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                }

                //----------------------------->Warehouses<-----------------------------
                else if (BusinessObjectInfo.Type == "64")
                {
                    Warehouses.uiApp_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                }

                //----------------------------->Landed Costs Setup-----------------------------
                else if (BusinessObjectInfo.Type == "48")
                {
                    LandedCostsSetup.uiApp_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                }

                //----------------------------->Credit Line Master Data<----------------------------
                else if (BusinessObjectInfo.Type == "UDO_F_BDOSCRLN_D")
                {
                    BDOSCreditLine.uiApp_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                }

                //----------------------------->Interest Accrual Document<----------------------------
                else if (BusinessObjectInfo.Type == "UDO_F_BDOSINAC_D")
                {
                    BDOSInterestAccrual.uiApp_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                }

                //----------------------------->A/R Correction Invoice<-----------------------------
                else if (BusinessObjectInfo.Type == "165" && BusinessObjectInfo.FormTypeEx == "70008")
                {
                    ArCorrectionInvoice.UiApp_FormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
                }
            }
            catch (Exception ex)
            {
                uiApp.StatusBar.SetSystemMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short);
            }
        }

        public static void translateFormTitle(ref SAPbouiCOM.ItemEvent pVal)
        {
            if ((pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE & FORM_LOAD_FOR_VISIBLE || pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD) & !pVal.BeforeAction)
            {
                SAPbouiCOM.Form oForm = uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);
                string title = oForm.Title;
                int substringLength = (title.Contains("სია") ? 4 : 5);

                if (title.Contains("Item Categories"))
                {
                    oForm.Title = title.Substring(0, substringLength) + BDOSResources.getTranslate("ItemCategories");
                }
                else if (title.Contains("Drivers"))
                {
                    oForm.Title = title.Substring(0, substringLength) + BDOSResources.getTranslate("Drivers");
                }
                else if (title.Contains("Profit Tax Base"))
                {
                    oForm.Title = title.Substring(0, substringLength) + BDOSResources.getTranslate("ProfitTaxBase");
                }
                else if (title.Contains("Profit Tax Base Type"))
                {
                    oForm.Title = title.Substring(0, substringLength) + BDOSResources.getTranslate("ProfitTaxBaseType");
                }
                else if (title.Contains("Vehicles"))
                {
                    oForm.Title = title.Substring(0, substringLength) + BDOSResources.getTranslate("Vehicles");
                }
                else if (title.Contains("Profit Tax Accrual"))
                {
                    oForm.Title = title.Substring(0, substringLength) + BDOSResources.getTranslate("ProfitTaxAccrual");
                }
                else if (title.Contains("Tax Invoice Received"))
                {
                    oForm.Title = title.Substring(0, substringLength) + BDOSResources.getTranslate("TaxInvoiceReceived");
                }
                else if (title.Contains("Banking Integration Rules"))
                {
                    oForm.Title = title.Substring(0, substringLength) + BDOSResources.getTranslate("BankingIntegrationRules");
                }
                else if (title.Contains("Tax Invoice Sent"))
                {
                    oForm.Title = title.Substring(0, substringLength) + BDOSResources.getTranslate("TaxInvoiceSent");
                }
                else if (title.Contains("Waybill"))
                {
                    oForm.Title = title.Substring(0, substringLength) + BDOSResources.getTranslate("Waybill");
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE & FORM_LOAD_FOR_VISIBLE)
                {
                    FORM_LOAD_FOR_VISIBLE = false;
                }
                else
                {
                    FORM_LOAD_FOR_VISIBLE = true;
                }
            }
        }

        static void uiApp_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            //----------------------------->ლიცენზირების ფორმა<-----------------------------
                if (pVal.FormUID == "BDOSLocLicForm" && pVal.ItemUID == "3" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && !pVal.BeforeAction)
            {
                SAPbouiCOM.Form oForm = uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);
                updateProgramLicense(oForm, out errorText);
            }
            if (!localisationAddonLicensed) return;

            try
            {
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE)
                {
                    currentFormCount = pVal.FormTypeCount;
                }

                //ჩვენი ცხრილების არჩევის ფორმები
                if (pVal.FormTypeEx == "9999")
                {
                    translateFormTitle(ref pVal);
                }

                //----------------------------->Profit Tax Base Type Master Data<-----------------------------
                else if (pVal.FormTypeEx == "UDO_F_BDO_PTBT_D")
                {
                    BDO_ProfitTaxBaseType.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK)
                    {
                        removeRecordRow = pVal.Row;
                    }
                }

                //----------------------------->Profit Tax Base Master Data<-----------------------------
                else if (pVal.FormTypeEx == "UDO_F_BDO_PTBS_D")
                {
                    BDO_ProfitTaxBase.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK)
                    {
                        removeRecordRow = pVal.Row;
                    }
                }

                //----------------------------->Item Categories<-----------------------------
                else if (pVal.FormTypeEx == "UDO_F_BDOSITMCTG_D")
                {
                    BDOSItemCategories.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK)
                    {
                        removeRecordRow = pVal.Row;
                    }
                }

                //----------------------------->Statement of Cash Flows<-----------------------------
                else if (pVal.FormTypeEx == "80028")
                {
                    StatementOfCashFlows.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->Trial Balance<-----------------------------
                else if (pVal.FormTypeEx == "167")
                {
                    TrialBalace.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->Balance Sheet<-----------------------------
                else if (pVal.FormTypeEx == "163")
                {
                    BalanceSheet.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->Profit and Loss Statement<-----------------------------
                else if (pVal.FormTypeEx == "164")
                {
                    ProfitAndLossStatement.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->Waybills Analysis Received<-----------------------------
                else if (pVal.FormUID == "BDOSWBRAn" || pVal.FormUID == "BDOSSelectValues")
                {
                    BDOSWaybillsAnalysisReceived.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->Waybills Analysis Sent<-----------------------------
                else if (pVal.FormUID == "BDOSWBSAn")
                {
                    BDOSWaybillsAnalysisSent.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->Tax Analysis Received, Settlement Reconciliation Act, Tax Analysiss Sent, DownPayment Tax Analysis Received<-----------------------------
                else if (pVal.FormTypeEx == "410000100")
                {
                    BDOSTaxAnalysisReceived.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                    BDOSReportSettlementReconciliationAct.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                    BDOSTaxAnalysissSent.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                    BDOSDownPaymentTaxAnalysisReceived.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->Journal Entry<-----------------------------
                else if (pVal.FormTypeEx == "392")
                {
                    JournalEntry.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->Cart Of Accounts<-----------------------------
                else if (pVal.FormTypeEx == "804")
                {
                    ChartOfAccounts.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                else if (pVal.FormTypeEx == "70001" )
                {
                    StockRevaluation.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->AP correction invoice <------------------------------
                else if (pVal.FormTypeEx== "70002" || pVal.FormTypeEx == "0")
                {
                    APCorrectionInvoice.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->Exchange Rate Differences<-----------------------------
                else if (pVal.FormTypeEx == "369")
                {
                    ExchangeRateDiffs.UiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                else if (pVal.FormTypeEx == "370")
                {
                    ExchangeRateDiffs.SetExecDate(pVal);
                }

                //----------------------------->Landed Costs<-----------------------------
                else if (pVal.FormTypeEx == "992")
                {
                    LandedCosts.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                ////----------------------------->Stock Revaluation<-----------------------------
                //else if (pVal.FormTypeEx == "70001")
                //{
                //    StockRevaluation.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                //}

                //----------------------------->TAX Groups<-----------------------------
                else if (pVal.FormTypeEx == "895")
                {
                    VatGroup.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->მიღებული ფაქტურების ჟურნალი<-----------------------------
                else if (pVal.FormUID == "BDOSTaxRecvForm")
                {
                    BDOSTaxJournal.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->RS - ს კოდების მითითება საზომ ერთეულებზე<-----------------------------
                else if (pVal.FormUID == "BDO_RSUoMForm")
                {
                    BDO_RSUoM.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->UoM Setup<-----------------------------
                else if (pVal.FormTypeEx == "13000005")
                {
                    BDO_RSUoM.uiApp_ItemEvent_Setup(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->A/P Credit Memo<-----------------------------
                else if (pVal.FormTypeEx == "181") // || pVal.FormTypeEx == "60504"
                {
                    APCreditMemo.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->შესაბამისობის კატალოგი<-----------------------------
                else if (pVal.FormTypeEx == "993" & !pVal.BeforeAction)
                {
                    BDO_BPCatalog.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->მიღებული ზედნადებების ჟურნალი<-----------------------------
                else if (pVal.FormUID == "BDO_WaybillsReceivedForm" || pVal.FormUID == "BDO_WaybillsReceivedNewRowForm")
                {
                    BDO_WaybillsJournalReceived.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->Exchange Rates And Indexes<-----------------------------
                else if (pVal.FormTypeEx == "866" & !pVal.BeforeAction)
                {
                    ExchangeFormRatesAndIndexes.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->Import Rate<-----------------------------
                else if (pVal.FormUID == "BDO_ImportRateForm" & !pVal.BeforeAction) //60004
                {
                    BDO_ImportRateForm.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->Waybill Journal<-----------------------------
                else if (pVal.FormUID == "BDO_WaybillsSentForm")
                {
                    BDO_WaybillsJournalSent.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->BDOSOutgoingPaymentsWizard<-----------------------------
                else if (pVal.FormUID == "BDOSSOPWizzForm")
                {
                    BDOSOutgoingPaymentsWizard.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->Stock Transfer Wizard<-----------------------------
                else if (pVal.FormUID == "BDOSSTTRWZ" || pVal.FormUID == "BDOSStockTransferDetail" || pVal.FormUID == "BDOSStockTransferSplit")
                {
                    BDOSStockTransferWizard.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->BDOSDepreciationAccruing<-----------------------------
                else if (pVal.FormUID == "BDOSDepAccrForm")
                {
                    BDOSDepreciationAccrualWizard.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->AR Down Payment VAT Accrual Wizard<-----------------------------
                else if (pVal.FormUID == "BDOSVAWizzForm")
                {
                    BDOSVATAccrualWizard.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                else if (pVal.FormUID == "BDOSVATADD")
                {
                    BDOSVATAccrualWizard.uiApp_ItemEventAddForm(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->Fuel Write-Off Wizard<-----------------------------
                else if (pVal.FormUID == "BDOSFuelWOForm")
                {
                    BDOSFuelWriteOffWizard.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->VAT Reconcilation Wizard<-----------------------------
                else if (pVal.FormUID == "BDOSReconWizz")
                {
                    BDOSVATReconcilationWizard.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->Business Partner Master Data<-----------------------------
                else if (pVal.FormTypeEx == "134")
                {
                    BusinessPartners.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->Items Master Data<-----------------------------
                else if (pVal.FormTypeEx == "150")
                {
                    Items.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->ItemsGroup Master Data<-----------------------------
                else if (pVal.FormTypeEx == "63")
                {
                    ItemGroup.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->Asset Master Data<-----------------------------
                else if (pVal.FormTypeEx == "1473000075" || pVal.FormUID == "NewCostCenterForm")
                {
                    FixedAsset.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->Capitalization<-----------------------------
                else if (pVal.FormTypeEx == "1470000009")
                {
                    Capitalization.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->Company details<-----------------------------
                else if (pVal.FormTypeEx == "136")
                {
                    CompanyDetails.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->General Settings<-----------------------------
                else if (pVal.FormTypeEx == "138")
                {
                    if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
                    {
                        SAPbouiCOM.Form oForm = uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);
                        if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD & !pVal.BeforeAction)
                        {
                            SAPbouiCOM.Item oItem;
                            SAPbouiCOM.EditText oEditText;
                            oItem = oForm.Items.Item("2018");
                            oEditText = (SAPbouiCOM.EditText)oItem.Specific;
                        }
                    }
                    GeneralSettings.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->Company Details-ზე პაროლის დანიშვნის ფორმა<-----------------------------
                else if (pVal.FormUID == "BDO_SetPasForm")
                {
                    CompanyDetails.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->Fuel Transfer Wizard<-----------------------------
                else if (pVal.FormUID == "BDOSFuelTransferWizard")
                {
                    BDOSFuelTransferWizard.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->Users - Setup<-----------------------------
                else if (pVal.FormTypeEx == "20700")
                {
                    Users.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->Driver Master Data<-----------------------------
                else if (pVal.FormTypeEx == "UDO_FT_UDO_F_BDO_DRVS_D")
                {
                    BDO_Drivers.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK)
                    {
                        removeRecordRow = 1;
                    }
                }

                //----------------------------->Vehicle Master Data<-----------------------------
                else if (pVal.FormTypeEx == "UDO_FT_UDO_F_BDO_VECL_D")
                {
                    BDO_Vehicles.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK)
                    {
                        removeRecordRow = 1;
                    }
                }

                //----------------------------->Fuel Types Master Data<-----------------------------
                else if (pVal.FormTypeEx == "UDO_FT_UDO_F_BDOSFUTP_D")
                {
                    BDOSFuelTypes.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK)
                    {
                        removeRecordRow = 1;
                    }
                }

                //----------------------------->Fuel Criteria Master Data<-----------------------------
                else if (pVal.FormTypeEx == "UDO_F_BDOSFUCR_D")
                {
                    BDOSFuelCriteria.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK)
                    {
                        removeRecordRow = 1;
                    }
                }

                //----------------------------->Specification of Fuel Norm Master Data<-----------------------------
                else if (pVal.FormTypeEx == "UDO_FT_UDO_F_BDOSFUNR_D")
                {
                    BDOSFuelNormSpecification.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK)
                    {
                        removeRecordRow = 1;
                    }
                }

                //----------------------------->Fuel Consumption Act Document<-----------------------------
                else if (pVal.FormTypeEx == "UDO_FT_UDO_F_BDOSFUCN_D")
                {
                    BDOSFuelConsumptionAct.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->Waybill Document<-----------------------------
                else if (pVal.FormTypeEx == "UDO_FT_UDO_F_BDO_WBLD_D")
                {
                    BDO_Waybills.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->Warehouse Addresses<-----------------------------
                else if (pVal.FormTypeEx == "11163")
                {
                    BDOSWarehouseAddresses.UiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK)
                    {
                        removeRecordRow = 1;
                    }
                }

                //----------------------------->A/P Invoice<-----------------------------
                else if (pVal.FormTypeEx == "141") // || pVal.FormTypeEx == "60504"
                {
                    APInvoice.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->GoodsReceiptPO<-----------------------------
                else if (pVal.FormTypeEx == "143")
                {
                    GoodsReceiptPO.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->A/P Reserve Invoice<-----------------------------
                else if (pVal.FormTypeEx == "60092") // || pVal.FormTypeEx == "60504"
                {
                    APReserveInvoice.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->A/R Invoice<-----------------------------
                else if (pVal.FormTypeEx == "133")
                {
                    ARInvoice.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->Blanket Agreement<-----------------------------
                else if (pVal.FormTypeEx == "1250000100" || pVal.FormTypeEx == "1250000102")
                {
                    BlanketAgreement.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->A/R Reserve Invoice<-----------------------------
                else if (pVal.FormTypeEx == "60091")
                {
                    ARReserveInvoice.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->Inventory Transfer<-----------------------------
                else if (pVal.FormTypeEx == "940")
                {
                    StockTransfer.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->Internal Reconcillation<-----------------------------
                else if (pVal.FormTypeEx == "120060805" || pVal.FormTypeEx == "0")
                {
                    InternalReconciliation.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //-----------------------------Inventory Transfer Request<-----------------------------
                else if (pVal.FormTypeEx == "1250000940")
                {
                    StockTransferRequest.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->A/R Credit Note<-----------------------------
                else if (pVal.FormTypeEx == "179")
                {
                    ARCreditNote.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->Sales Order<-----------------------------
                else if (pVal.FormTypeEx == "139")
                {
                    SalesOrder.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->Return<-----------------------------
                else if (pVal.FormTypeEx == "180")
                {
                    Return.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->Delivery<-----------------------------
                else if (pVal.FormTypeEx == "140")
                {
                    Delivery.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->Tax Invoice Received <-----------------------------
                else if (pVal.FormTypeEx == "UDO_FT_UDO_F_BDO_TAXR_D")
                {
                    BDO_TaxInvoiceReceived.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                else if (pVal.FormUID == "BDO_TaxInvoiceReceivedChoose")
                {
                    BDO_TaxInvoiceReceived.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->Tax Invoice Sent<-----------------------------
                else if (pVal.FormTypeEx == "UDO_FT_UDO_F_BDO_TAXS_D")
                {
                    BDO_TaxInvoiceSent.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->A/R Down Payment VAT Accrual<-----------------------------
                else if (pVal.FormTypeEx == "UDO_FT_UDO_F_BDO_ARDPV_D")
                {
                    BDOSARDownPaymentVATAccrual.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->Profit Tax Accural<-----------------------------
                else if (pVal.FormTypeEx == "UDO_FT_UDO_F_BDO_TAXP_D")
                {
                    BDO_ProfitTaxAccrual.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->Fixed Asset Transfer Document<-----------------------------
                else if (pVal.FormTypeEx == "UDO_FT_UDO_F_BDOSFASTRD_D")
                {
                    BDOSFixedAssetTransfer.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->Depreciation Accrual<-----------------------------
                else if (pVal.FormTypeEx == "UDO_FT_UDO_F_BDOSDEPACR_D")
                {
                    BDOSDepreciationAccrualDocument.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->Withholding Tax<-----------------------------
                else if (pVal.FormTypeEx == "65015")
                {
                    WithholdingTax.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->House Bank Accounts<-----------------------------
                else if (pVal.FormTypeEx == "60701")
                {
                    HouseBankAccounts.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->Outgoing Paymentss<-----------------------------
                else if (pVal.FormTypeEx == "426" || pVal.FormUID == "OutgoingPaymentNewDate")
                {
                    OutgoingPayment.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->Incoming Paymentss<-----------------------------
                else if (pVal.FormTypeEx == "170")
                {
                    IncomingPayment.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->Goods Issue<-----------------------------
                else if (pVal.FormTypeEx == "720")
                {
                    GoodsIssue.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->Retirement<-----------------------------
                else if (pVal.FormTypeEx == "1470000014")
                {
                    Retirement.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->A/P Down Payment Request<-----------------------------
                else if (pVal.FormTypeEx == "65309") // || pVal.FormTypeEx == "60504"
                {
                    APDownPaymentRequest.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->A/P Down Payment Invoice<-----------------------------
                else if (pVal.FormTypeEx == "65301")
                {
                    APDownPaymentInvoice.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->აუთენთიფიკაციის ფორმა (INTERNET BANK - TBC)<-----------------------------
                else if (pVal.FormUID == "BDOSAuthenticationFormTBC")
                {
                    BDOSAuthenticationFormTBC.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->პაროლის შეცვლის ფორმა (INTERNET BANK - TBC)<-----------------------------
                else if (pVal.FormUID == "BDOSChangePasswordFormTBC")
                {
                    BDOSAuthenticationFormTBC.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->აუთენთიფიკაციის ფორმა (INTERNET BANK - BOG)<-----------------------------
                else if (pVal.FormUID == "BDOSAuthenticationFormBOG")
                {
                    BDOSAuthenticationFormBOG.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->Internet Banking<-----------------------------
                else if (pVal.FormUID == "BDOSInternetBankingForm")
                {
                    BDOSInternetBanking.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->Internet Bank documents<-----------------------------
                else if (pVal.FormUID == "BDOSINBDOC")
                {
                    BDOSInternetBankingDocuments.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->Payment Means from Outgoing Paymentss<-----------------------------
                else if (pVal.FormTypeEx == "196")
                {
                    PaymentMeans.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->Payment Means from Incoming Paymentss<-----------------------------
                else if (pVal.FormTypeEx == "146")
                {
                    PaymentMeans.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->Delete UDF<-----------------------------
                else if (pVal.FormUID == "BDOSDeleteUDFForm")
                {
                    BDOSDeleteUDF.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->A/R Down Payment Invoice<-----------------------------
                else if (pVal.FormTypeEx == "65300")
                {
                    ARDownPaymentInvoice.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->A/R Down Payment Request<-----------------------------
                else if (pVal.FormTypeEx == "65308")
                {
                    ARDownPaymentRequest.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->Internet Banking Integration Services Rules<-----------------------------
                else if (pVal.FormTypeEx == "UDO_F_BDOSINTR_D")
                {
                    BDOSInternetBankingIntegrationServicesRules.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->Document Settings<-----------------------------
                else if (pVal.FormTypeEx == "228")
                {
                    DocumentSettings.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->>WareHouses<-----------------------------
                else if (pVal.FormTypeEx == "62")
                {
                    Warehouses.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->BP Bank Accounts<-----------------------------
                else if (pVal.FormTypeEx == "65052")
                {
                    BPBankAccounts.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->>Asset Class<-----------------------------
                else if (pVal.FormTypeEx == "1472000006")
                {
                    AssetClass.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->Withholding Tax Table<-----------------------------
                else if (pVal.FormTypeEx == "60504") //from A/P Credit Memo, A/P Invoice, A/P Reserve Invoice, A/P Down Payment Request 
                {
                    WithholdingTax.openTaxTableFromAPDocs(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->Issue For Production<-----------------------------
                else if (pVal.FormTypeEx == "65213")
                {
                    IssueForProduction.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->Select Items to Copy (Issue For Production)<-----------------------------
                else if (pVal.FormTypeEx == "65212")
                {
                    if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
                    {
                        SAPbouiCOM.Form oForm = uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                        if (pVal.ItemUID == "1" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && !pVal.BeforeAction)
                            selectItemsToCopyOkClick = true;
                    }
                }

                //----------------------------->Credit Line Master Data<-----------------------------
                else if (pVal.FormTypeEx == "UDO_F_BDOSCRLN_D")
                {
                    BDOSCreditLine.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK)
                    {
                        removeRecordRow = 1;
                    }
                }

                //----------------------------->Interest Accrual Document<-----------------------------
                else if (pVal.FormTypeEx == "UDO_FT_UDO_F_BDOSINAC_D")
                {
                    BDOSInterestAccrual.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->Interest Accrual Wizard<-----------------------------
                else if (pVal.FormUID == "BDOSInterestAccrualWizard")
                {
                    BDOSInterestAccrualWizard.uiApp_ItemEvent(FormUID, ref pVal, out BubbleEvent);
                }

                //----------------------------->A/R Correction Invoice<-----------------------------
                else if (pVal.FormTypeEx == "70008" )
                {
                    ArCorrectionInvoice.UiApp_ItemEvent(ref pVal, out BubbleEvent);
                }

                //----------------------------->Batch Number Selection<-----------------------------
                else if (pVal.FormTypeEx == "42")
                {
                    BatchNumberSelection.UiApp_ItemEvent(ref pVal, out BubbleEvent);
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
                try
                {
                    uiApp.StatusBar.SetSystemMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short);
                }
                catch
                {
                }
            }
        }

        static void uiApp_RightClickEvent(ref SAPbouiCOM.ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (!localisationAddonLicensed) return;

            SAPbouiCOM.Form oForm = null;

            //wizardebi icrasheboda zogjer right click-ze
            try
            {
                oForm = uiApp.Forms.ActiveForm;

                //----------------------------->>Fixed Asset Transfer Document<-----------------------------
                if (oForm.TypeEx == "UDO_FT_UDO_F_BDOSFASTRD_D")
                {
                    BDOSFixedAssetTransfer.uiApp_RightClickEvent(oForm, eventInfo, out BubbleEvent);
                }

                //----------------------------->>Depreciation Document<-----------------------------
                else if (oForm.TypeEx == "UDO_FT_UDO_F_BDOSDEPACR_D")
                {
                    BDOSDepreciationAccrualDocument.uiApp_RightClickEvent(oForm, eventInfo, out BubbleEvent);
                }

                else
                {
                    try
                    {
                        uiApp.Menus.RemoveEx("BDOSAddRow");
                    }
                    catch { }
                    try
                    {
                        uiApp.Menus.RemoveEx("BDOSDelRow");
                    }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                uiApp.StatusBar.SetSystemMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short);
            }

            if (oForm == null)
            {
                return;
            }

            SAPbouiCOM.Item oItem = null;

            string DocEntry = "";

            try
            {
                oItem = oForm.Items.Item("0_U_E");
                DocEntry = oItem.Specific.Value;
                DocEntry = DocEntry.Trim();
            }
            catch
            {
            }

            if (eventInfo.BeforeAction)
            {
                if (!uiApp.Menus.Exists("6005") && oItem != null && DocEntry == "")
                {
                    SAPbouiCOM.MenuItem oMenuItem;
                    SAPbouiCOM.Menus oMenus;
                    SAPbouiCOM.MenuCreationParams oCreationPackage;

                    try
                    {
                        oCreationPackage = (SAPbouiCOM.MenuCreationParams)uiApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                        oCreationPackage.UniqueID = "PreviewUDOJrE";
                        oCreationPackage.String = BDOSResources.getTranslate("PreviewJournalEntry");
                        oCreationPackage.Enabled = true;
                        oCreationPackage.Position = -1;

                        oMenuItem = uiApp.Menus.Item("1280");
                        oMenus = oMenuItem.SubMenus;
                        oMenus.AddEx(oCreationPackage);

                        clickUnitedJournalEntry = true;
                    }
                    catch (Exception ex)
                    {
                        string errorText = ex.Message;
                    }
                }
                else
                {
                    try
                    {
                        uiApp.Menus.RemoveEx("PreviewUDOJrE");
                    }
                    catch { }
                }

                //აღდგენის (restore) წაშლა მარჯვენა-კლიკის კონტექსტური მენიუდან
                if (uiApp.Menus.Exists("1285"))
                {
                    try
                    {
                        uiApp.Menus.RemoveEx("1285");
                    }
                    catch { }
                }
            }
            else
            {
            }
        }

        public static void updateProgramLicense(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            int answer = uiApp.MessageBox(BDOSResources.getTranslate("AfterChangingLicenseKey"), 1, BDOSResources.getTranslate("Yes"), BDOSResources.getTranslate("No"), "");
            if (answer == 1)
            {
                string licenseKey = oForm.Items.Item("BDOSLicKey").Specific.value;

                License oLicense = new License();
                bool result = oLicense.LicenseSuccessfull(licenseKey);

                oForm.Close();
                try
                {
                    uiApp.ActivateMenuItem("1026");
                }
                catch { }

                Dictionary<string, string> CompanyLicenseInfo = CommonFunctions.getCompanyLicenseInfo();
                string licenseStatus = CompanyLicenseInfo["LicenseStatus"];
                string licenseUpdateDate = CompanyLicenseInfo["LicenseUpdateDate"];
                string licenseQuantity = CompanyLicenseInfo["LicenseQuantity"];

                if (licenseStatus == BDOSResources.getTranslate("Active"))
                {
                    if (!localisationAddonLicensed) runLocalisationAddOn();
                }
                else
                {
                    localisationAddonLicensed = false;
                    uiApp.MessageBox(BDOSResources.getTranslate("LocalisationAddonNotLicensed"));
                }
            }
        }
    }
}