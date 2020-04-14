using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace BDO_Localisation_AddOn
{
    static partial class CrystalReports
    {
        public static void addStandAloneCrystalReportForAddOn( string startupPath, out string errorText)
        {
            errorText = null;

            bool hanaDB = (Program.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB);

            string reportName = "Output VAT Analysis";
            string reportTypeCode = "RCRI"; //Use TypeCode "RCRI" to specify a Crystal Report
            string rptFilePath = startupPath + (hanaDB ? @"\CrystalReports\Report_OutputVATAnalysis.rpt" : @"\CrystalReports_SQL\Report_OutputVATAnalysis_SQL.rpt");
            string menuID = "12800"; //salesReports
            string layoutCode = null;

            layoutCode = getReportlayoutCode( reportName, reportTypeCode, out errorText);

            addCrystalReport( reportName, reportTypeCode, rptFilePath, menuID, SAPbobsCOM.ReportLayoutCategoryEnum.rlcCrystal, ref layoutCode, out errorText);

            //მოგების გადასახადი
            reportName = "Profit Tax Analysis";
            reportTypeCode = "RCRI"; //Use TypeCode "RCRI" to specify a Crystal Report
            rptFilePath = startupPath + (hanaDB ? @"\CrystalReports\Report_ProfitTax.rpt" : @"\CrystalReports_SQL\Report_ProfitTax_SQL.rpt");
            menuID = "9728"; //Financial Reports/Financials
            layoutCode = null;

            layoutCode = getReportlayoutCode( reportName, reportTypeCode, out errorText);

            addCrystalReport( reportName, reportTypeCode, rptFilePath, menuID, SAPbobsCOM.ReportLayoutCategoryEnum.rlcCrystal, ref layoutCode, out errorText);

            //მომწოდებლებთან ანგარიშსწორება
            reportName = "Supplier Liabilities By Currencies";
            reportTypeCode = "RCRI"; //Use TypeCode "RCRI" to specify a Crystal Report
            rptFilePath = startupPath + (hanaDB ? @"\CrystalReports\Report_SupplierLiabilitiesByCurrencies.rpt" : @"\CrystalReports_SQL\Report_SupplierLiabilitiesByCurrencies_SQL.rpt");
            menuID = "43536"; //Business Patrner Reports/Business Patrners
            layoutCode = null;

            layoutCode = getReportlayoutCode( reportName, reportTypeCode, out errorText);

            addCrystalReport( reportName, reportTypeCode, rptFilePath, menuID, SAPbobsCOM.ReportLayoutCategoryEnum.rlcCrystal, ref layoutCode, out errorText);

            //ურთიერთანგარიშსწორების შეჯერების აქტი
            reportName = "Settlement Reconciliation Act";
            reportTypeCode = "RCRI"; //Use TypeCode "RCRI" to specify a Crystal Report
            rptFilePath = startupPath + (hanaDB ? @"\CrystalReports\Report_SettlementReconciliationAct.rpt" : @"\CrystalReports_SQL\Report_SettlementReconciliationAct_SQL.rpt");
            menuID = "43536"; //Business Patrner Reports/Business Patrners
            layoutCode = null;

            layoutCode = getReportlayoutCode( reportName, reportTypeCode, out errorText);

            addCrystalReport( reportName, reportTypeCode, rptFilePath, menuID, SAPbobsCOM.ReportLayoutCategoryEnum.rlcCrystal, ref layoutCode, out errorText);

            //მყიდველების დავალიანებები გაყიდვების ჯგუფების მიხედვით
            reportName = "Supplier Liabilities by Currencies";
            reportTypeCode = "RCRI"; //Use TypeCode "RCRI" to specify a Crystal Report
            rptFilePath = startupPath + (hanaDB ? @"\CrystalReports\Report_SupplierLiabilitiesByCurrencies.rpt": @"\CrystalReports_SQL\Report_SupplierLiabilitiesByCurrencies_SQL.rpt");
            menuID = "13056"; //Accounting/Financial Reports/Reports
            layoutCode = null;

            layoutCode = getReportlayoutCode( reportName, reportTypeCode, out errorText);

            addCrystalReport( reportName, reportTypeCode, rptFilePath, menuID, SAPbobsCOM.ReportLayoutCategoryEnum.rlcCrystal, ref layoutCode, out errorText);

            //დეტალური ბრუნვითი უწყისი
            reportName = "Detailed Trial Balance";
            reportTypeCode = "RCRI"; //Use TypeCode "RCRI" to specify a Crystal Report
            rptFilePath = startupPath + (hanaDB ? @"\CrystalReports\Report_DetailedTrialBalance.rpt" : @"\CrystalReports_SQL\Report_DetailedTrialBalance_SQL.rpt");
            menuID = "13056"; //Accounting/Financial Reports/Reports
            layoutCode = null;

            layoutCode = getReportlayoutCode( reportName, reportTypeCode, out errorText);

            addCrystalReport( reportName, reportTypeCode, rptFilePath, menuID, SAPbobsCOM.ReportLayoutCategoryEnum.rlcCrystal, ref layoutCode, out errorText);

            //Capitalization
            reportName = "Capitalization Analysis";
            reportTypeCode = "RCRI"; //Use TypeCode "RCRI" to specify a Crystal Report
            rptFilePath = startupPath + (hanaDB ? @"\CrystalReports\Report_Capitalization.rpt" : @"\CrystalReports_SQL\Report_Capitalization_SQL.rpt");
            menuID = "43531"; //Financial Reports
            layoutCode = null;

            layoutCode = getReportlayoutCode(reportName, reportTypeCode, out errorText);

            addCrystalReport(reportName, reportTypeCode, rptFilePath, menuID, SAPbobsCOM.ReportLayoutCategoryEnum.rlcCrystal, ref layoutCode, out errorText);

            //მიღებული ა/ფ ანალიზი
            reportName = "Tax Invoice Received Analysis";
            reportTypeCode = "RCRI"; //Use TypeCode "RCRI" to specify a Crystal Report
            rptFilePath = startupPath + (hanaDB ? @"\CrystalReports\Report_TaxInvoiceReceivedAnalysis.rpt" : @"\CrystalReports_SQL\Report_TaxInvoiceReceivedAnalysis_SQL.rpt");
            menuID = "43534"; //Purchasing Reports/Purchasing
            layoutCode = null;

            layoutCode = getReportlayoutCode( reportName, reportTypeCode, out errorText);

            addCrystalReport( reportName, reportTypeCode, rptFilePath, menuID, SAPbobsCOM.ReportLayoutCategoryEnum.rlcCrystal, ref layoutCode, out errorText);

            //შესყიდვის ფაქტურების ანგარიშგება
            reportName = "Purchase Tax Invoice Analysis";
            reportTypeCode = "RCRI"; //Use TypeCode "RCRI" to specify a Crystal Report
            rptFilePath = startupPath + (hanaDB ? @"\CrystalReports\Report_PurchaseTaxInvoiceAnalysis.rpt" : @"\CrystalReports_SQL\Report_PurchaseTaxInvoiceAnalysis_SQL.rpt");
            menuID = "43534"; //Purchasing Reports/Purchasing
            layoutCode = null;

            layoutCode = getReportlayoutCode(reportName, reportTypeCode, out errorText);

            addCrystalReport(reportName, reportTypeCode, rptFilePath, menuID, SAPbobsCOM.ReportLayoutCategoryEnum.rlcCrystal, ref layoutCode, out errorText);

            //მიღებული ავანსის ა/ფ ანალიზი
            reportName = "Down Payment Tax Invoice Received Analysis";
            reportTypeCode = "RCRI"; //Use TypeCode "RCRI" to specify a Crystal Report
            rptFilePath = startupPath + (hanaDB ? @"\CrystalReports\Report_DPTaxInvoiceReceivedAnalysis.rpt" : @"\CrystalReports_SQL\Report_DPTaxInvoiceReceivedAnalysis_SQL.rpt");
            menuID = "43534"; //Purchasing Reports/Purchasing
            layoutCode = null;

            layoutCode = getReportlayoutCode( reportName, reportTypeCode, out errorText);

            addCrystalReport( reportName, reportTypeCode, rptFilePath, menuID, SAPbobsCOM.ReportLayoutCategoryEnum.rlcCrystal, ref layoutCode, out errorText);


            //გაცემული ა/ფ ანალიზი
            reportName = "Tax Invoice Sent Analysis";
            reportTypeCode = "RCRI"; //Use TypeCode "RCRI" to specify a Crystal Report
            rptFilePath = startupPath + (hanaDB ? @"\CrystalReports\Report_TaxInvoiceSentAnalysis.rpt" : @"\CrystalReports_SQL\Report_TaxInvoiceSentAnalysis_SQL.rpt");
            menuID = "12800"; //Sales Reports/Sales
            layoutCode = null;

            layoutCode = getReportlayoutCode( reportName, reportTypeCode, out errorText);

            addCrystalReport( reportName, reportTypeCode, rptFilePath, menuID, SAPbobsCOM.ReportLayoutCategoryEnum.rlcCrystal, ref layoutCode, out errorText);

            //მარაგების მოძრაობა
            reportName = "Stock Turnover";
            reportTypeCode = "RCRI"; //Use TypeCode "RCRI" to specify a Crystal Report
            rptFilePath = startupPath + (hanaDB ? @"\CrystalReports\Report_StockTurnover.rpt" : @"\CrystalReports_SQL\Report_StockTurnover_SQL.rpt");
            menuID = "1760"; //Stock Managment/Stock Reports
            layoutCode = null;

            layoutCode = getReportlayoutCode( reportName, reportTypeCode, out errorText);

            addCrystalReport( reportName, reportTypeCode, rptFilePath, menuID, SAPbobsCOM.ReportLayoutCategoryEnum.rlcCrystal, ref layoutCode, out errorText);

            //მარაგების მოძრაობა საწყობით დაჯგუფებული
            reportName = "Stock Turnover By Grouping";
            reportTypeCode = "RCRI"; //Use TypeCode "RCRI" to specify a Crystal Report
            rptFilePath = startupPath + (hanaDB ? @"\CrystalReports\Report_StockTurnoverByGrouping.rpt" : @"\CrystalReports_SQL\Report_StockTurnoverByGrouping_SQL.rpt");
            menuID = "1760"; //Stock Managment/Stock Reports
            layoutCode = null;

            layoutCode = getReportlayoutCode( reportName, reportTypeCode, out errorText);

            addCrystalReport( reportName, reportTypeCode, rptFilePath, menuID, SAPbobsCOM.ReportLayoutCategoryEnum.rlcCrystal, ref layoutCode, out errorText);

            //საერთო მოგება
            reportName = "Gross Profit";
            reportTypeCode = "RCRI"; //Use TypeCode "RCRI" to specify a Crystal Report
            rptFilePath = startupPath + (hanaDB ? @"\CrystalReports\Report_GrossProfit.rpt" : @"\CrystalReports_SQL\Report_GrossProfit_SQL.rpt");
            menuID = "2048"; //Stock Managment/Stock Reports
            layoutCode = null;

            layoutCode = getReportlayoutCode( reportName, reportTypeCode, out errorText);

            addCrystalReport( reportName, reportTypeCode, rptFilePath, menuID, SAPbobsCOM.ReportLayoutCategoryEnum.rlcCrystal, ref layoutCode, out errorText);

        }

        public static void addDocumentTypeCrystalReportForAddOn( string startupPath, out string errorText)
        {
            errorText = null;

            bool hanaDB = (Program.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB);
            
            string typeName = "Waybill document type";
            string addonName = "BDOS Localisation AddOn";
            string addonFormType = "UDO_FT_UDO_F_BDO_WBLD_D";
            string typeCode = addReportType( typeName, addonName, addonFormType, out errorText);
            string layoutCode = null;

            if (string.IsNullOrEmpty(typeCode) == false)
            {
                string reportName = "Waybill Sent (BDOS)";
                string reportTypeCode = typeCode; //"WBLD"; //Use TypeCode "RCRI" to specify a Crystal Report
                string rptFilePath = startupPath + (hanaDB ? @"\CrystalReports\WBLD.rpt" : @"\CrystalReports_SQL\WBLD_SQL.rpt");
                string menuID = null;

                layoutCode = getReportlayoutCode( reportName, reportTypeCode, out errorText);

                addCrystalReport( reportName, reportTypeCode, rptFilePath, menuID, SAPbobsCOM.ReportLayoutCategoryEnum.rlcCrystal, ref layoutCode, out errorText);

                if (string.IsNullOrEmpty(layoutCode) == false && string.IsNullOrEmpty(errorText) == true)
                {
                    setDefaultReport( typeCode, layoutCode, out errorText);
                }
            }

            //საქონლის საბეჭდი ფორმა Goods List (BDOS)
            typeName = "Goods List";
            addonName = "BDOS Localisation AddOn";

            Dictionary<string, string> listDocuments = new Dictionary<string, string>();
            listDocuments.Add("720", "IGE1"); //Goods Isshue
            listDocuments.Add("133", "INV2"); //AR Invoice
            listDocuments.Add("141", "PCH2"); //AP Invoice
            listDocuments.Add("940", "WTR1"); //Inventory Transfer
            listDocuments.Add("140", "DLN2"); // Delivery

            foreach (KeyValuePair<string, string> keyValue in listDocuments)
            {
                layoutCode = null;

                string reportName = BDOSResources.getTranslate("GoodsList") + " (BDOS)";
                string reportTypeCode = keyValue.Value;
                string rptFilePath = startupPath + (hanaDB ? @"\CrystalReports\GoodsList.rpt" : @"\CrystalReports_SQL\GoodsList_SQL.rpt");
                string menuID = null;

                layoutCode = getReportlayoutCode( reportName, reportTypeCode, out errorText);

                addCrystalReport( reportName, reportTypeCode, rptFilePath, menuID, SAPbobsCOM.ReportLayoutCategoryEnum.rlcCrystal, ref layoutCode, out errorText);

                if (string.IsNullOrEmpty(layoutCode) == false && string.IsNullOrEmpty(errorText) == true)
                {
                    setDefaultReport( typeCode, layoutCode, out errorText);
                }
            }

            //Incoming Payment
            typeName = "Incoming Payment";
            addonName = "BDOS Localisation AddOn";
            addonFormType = "170";
            //typeCode = addReportType( typeName, addonName, addonFormType, out errorText);
            layoutCode = null;

            if (string.IsNullOrEmpty(typeCode) == false)
            {
                string reportName = BDOSResources.getTranslate("IncomingPayment") + " (BDOS)";
                string reportTypeCode = "RCT1";
                string rptFilePath = startupPath + (hanaDB ? @"\CrystalReports\IncomingPayment.rpt" : @"\CrystalReports_SQL\IncomingPayment_SQL.rpt");
                string menuID = null;

                layoutCode = getReportlayoutCode( reportName, reportTypeCode, out errorText);

                addCrystalReport( reportName, reportTypeCode, rptFilePath, menuID, SAPbobsCOM.ReportLayoutCategoryEnum.rlcCrystal, ref layoutCode, out errorText);

                if (string.IsNullOrEmpty(layoutCode) == false && string.IsNullOrEmpty(errorText) == true)
                {
                    setDefaultReport( typeCode, layoutCode, out errorText);
                }
            }

            //Outgoing Payment
            typeName = "Outgoing Payment";
            addonName = "BDOS Localisation AddOn";
            addonFormType = "170";
            //typeCode = addReportType( typeName, addonName, addonFormType, out errorText);
            layoutCode = null;

            if (string.IsNullOrEmpty(typeCode) == false)
            {
                string reportName = BDOSResources.getTranslate("OutgoingPayment") + " (BDOS)";
                string reportTypeCode = "VPM1";
                string rptFilePath = startupPath + (hanaDB ? @"\CrystalReports\OutgoingPayment.rpt" : @"\CrystalReports_SQL\OutgoingPayment_SQL.rpt");
                string menuID = null;

                layoutCode = getReportlayoutCode( reportName, reportTypeCode, out errorText);

                addCrystalReport( reportName, reportTypeCode, rptFilePath, menuID, SAPbobsCOM.ReportLayoutCategoryEnum.rlcCrystal, ref layoutCode, out errorText);

                if (string.IsNullOrEmpty(layoutCode) == false && string.IsNullOrEmpty(errorText) == true)
                {
                    setDefaultReport( typeCode, layoutCode, out errorText);
                }
            }
        }

        private static void addCrystalReport( string reportName, string reportTypeCode, string rptFilePath, string menuID, SAPbobsCOM.ReportLayoutCategoryEnum category, ref string layoutCode, out string errorText)
        {
            errorText = null;

            if (string.IsNullOrEmpty(layoutCode) == true)
            {
                SAPbobsCOM.ReportLayoutsService oReportLayoutsService = (SAPbobsCOM.ReportLayoutsService)Program.oCompany.GetCompanyService().GetBusinessService(SAPbobsCOM.ServiceTypes.ReportLayoutsService);
                SAPbobsCOM.ReportLayout oReportLayout = (SAPbobsCOM.ReportLayout)oReportLayoutsService.GetDataInterface(SAPbobsCOM.ReportLayoutsServiceDataInterfaces.rlsdiReportLayout);

                //Initialize critical properties 
                // 
                // Use TypeCode "RCRI" to specify a Crystal Report. 
                // Use other TypeCode to specify a layout for a document type. 
                // List of TypeCode types are in table RTYP. 
                oReportLayout.Name = reportName;
                oReportLayout.TypeCode = reportTypeCode;
                oReportLayout.Author = Program.oCompany.UserName;
                oReportLayout.Category = category;
                oReportLayout.ForeignLanguageReport = SAPbobsCOM.BoYesNoEnum.tYES;

                try
                {
                    // Add new object 
                    SAPbobsCOM.ReportLayoutParams oNewReportLayoutParams;

                    if (string.IsNullOrEmpty(menuID) == true)
                    {
                        oNewReportLayoutParams = oReportLayoutsService.AddReportLayout(oReportLayout); //LostReports
                    }
                    else
                    {
                        oNewReportLayoutParams = oReportLayoutsService.AddReportLayoutToMenu(oReportLayout, menuID);
                    }
                    // Get code of the added ReportLayout object 
                    layoutCode = oNewReportLayoutParams.LayoutCode;
                }

                catch (Exception ex)
                {
                    int errCode;
                    string errMsg;
                    Program.oCompany.GetLastError(out errCode, out errMsg);
                    errorText = BDOSResources.getTranslate("ErrorDescription") + " : " + errMsg + "! " + BDOSResources.getTranslate("Code") + " : " + errCode + "! " + BDOSResources.getTranslate("OtherInfo") + " : " + ex.Message;
                    return;
                }
            }
            // Wpload .rpt file using SetBlob interface 

            SAPbobsCOM.CompanyService oCompanyService = Program.oCompany.GetCompanyService();
            // Specify the table and field to update 
            SAPbobsCOM.BlobParams oBlobParams = (SAPbobsCOM.BlobParams)oCompanyService.GetDataInterface(SAPbobsCOM.CompanyServiceDataInterfaces.csdiBlobParams);
            oBlobParams.Table = "RDOC";
            oBlobParams.Field = "Template";

            // Specify the record whose blob field is to be set 
            SAPbobsCOM.BlobTableKeySegment oKeySegment = oBlobParams.BlobTableKeySegments.Add();
            oKeySegment.Name = "DocCode";
            oKeySegment.Value = layoutCode;

            SAPbobsCOM.Blob oBlob = (SAPbobsCOM.Blob)oCompanyService.GetDataInterface(SAPbobsCOM.CompanyServiceDataInterfaces.csdiBlob);

            // Put the rpt file into buffer 
            FileStream oFile = new FileStream(rptFilePath, System.IO.FileMode.Open, FileAccess.Read);
            int fileSize = (int)oFile.Length;
            byte[] buf = new byte[fileSize];
            oFile.Read(buf, 0, fileSize);
            oFile.Close();

            // Convert memory buffer to Base64 string 
            oBlob.Content = Convert.ToBase64String(buf, 0, fileSize);

            try
            {
                //Upload Blob to database 
                oCompanyService.SetBlob(oBlobParams, oBlob);
            }
            catch (Exception ex)
            {
                int errCode;
                string errMsg;
                Program.oCompany.GetLastError(out errCode, out errMsg);
                errorText = BDOSResources.getTranslate("ErrorDescription") + " : " + errMsg + "! " + BDOSResources.getTranslate("Code") + " : " + errCode + "! " + BDOSResources.getTranslate("OtherInfo") + " : " + ex.Message;
                return;
            }
        }

        private static string addReportType( string typeName, string addonName, string addonFormType, out string errorText)
        {
            errorText = null;
            try
            {
                string reportTypeCode = getReportTypeCode( addonName, addonFormType, out errorText);
                if (string.IsNullOrEmpty(reportTypeCode) == false)
                {
                    return reportTypeCode;
                }

                //1. Add a new report type => ReportTypesService (new in 8.81).
                SAPbobsCOM.ReportTypesService rptTypeService = (SAPbobsCOM.ReportTypesService)Program.oCompany.GetCompanyService().GetBusinessService(SAPbobsCOM.ServiceTypes.ReportTypesService);
                SAPbobsCOM.ReportTypesParams oReportTypeParams = rptTypeService.GetReportTypeList();

                SAPbobsCOM.ReportType newType = (SAPbobsCOM.ReportType)rptTypeService.GetDataInterface(SAPbobsCOM.ReportTypesServiceDataInterfaces.rtsReportType);

                newType.TypeName = typeName;
                newType.AddonName = addonName;
                newType.AddonFormType = addonFormType;

                SAPbobsCOM.ReportTypeParams newTypeParam = rptTypeService.AddReportType(newType);
                return newTypeParam.TypeCode;
            }
            catch (Exception ex)
            {
                int errCode;
                string errMsg;
                Program.oCompany.GetLastError(out errCode, out errMsg);
                errorText = BDOSResources.getTranslate("ErrorDescription") + " : " + errMsg + "! " + BDOSResources.getTranslate("Code") + " : " + errCode + "! " + BDOSResources.getTranslate("OtherInfo") + " : " + ex.Message;
                return null;
            }
            finally
            {
                GC.Collect();
            }
        }

        public static string getReportTypeCode( string addonName, string addonFormType, out string errorText)
        {
            errorText = null;

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                string query = "SELECT \"RTYP\".\"CODE\" " +
                    "FROM \"RTYP\" AS \"RTYP\" " +
                    "WHERE \"RTYP\".\"FRM_TYPE\" = '" + addonFormType + "' AND \"RTYP\".\"ADD_NAME\" = '" + addonName + "'";

                oRecordSet.DoQuery(query);
                while (!oRecordSet.EoF)
                {
                    return oRecordSet.Fields.Item("CODE").Value.ToString();
                }
                return null;
            }
            catch (Exception ex)
            {
                int errCode;
                string errMsg;
                Program.oCompany.GetLastError(out errCode, out errMsg);
                errorText = BDOSResources.getTranslate("ErrorDescription") + " : " + errMsg + "! " + BDOSResources.getTranslate("Code") + " : " + errCode + "! " + BDOSResources.getTranslate("OtherInfo") + " : " + ex.Message;
                return null;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                GC.Collect();
            }
        }

        public static void setDefaultReport( string reportCode, string layoutCode, out string errorText)
        {
            errorText = null;

            SAPbobsCOM.CompanyService oCmpSrv;
            SAPbobsCOM.ReportLayoutsService oReportLayoutService;
            SAPbobsCOM.DefaultReportParams oDefaultReportParams;

            try
            {
                //'get company service
                oCmpSrv = Program.oCompany.GetCompanyService();

                //'get report layout service
                oReportLayoutService = oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.ReportLayoutsService);

                //'get report layout params
                oDefaultReportParams = oReportLayoutService.GetDataInterface(SAPbobsCOM.ReportLayoutsServiceDataInterfaces.rlsdiDefaultReportParams);

                //'set the report layout code
                oDefaultReportParams.LayoutCode = layoutCode;

                //'set the report code
                //'the report code is the document type code (e.g. POR2=PurchaseOrder)
                oDefaultReportParams.ReportCode = reportCode;

                //'set the user code
                oDefaultReportParams.UserID = Program.oCompany.UserSignature;

                //'delete the report layout
                oReportLayoutService.SetDefaultReport(oDefaultReportParams);
            }
            catch (Exception ex)
            {
                int errCode;
                string errMsg;
                Program.oCompany.GetLastError(out errCode, out errMsg);
                errorText = BDOSResources.getTranslate("ErrorDescription") + " : " + errMsg + "! " + BDOSResources.getTranslate("Code") + " : " + errCode + "! " + BDOSResources.getTranslate("OtherInfo") + " : " + ex.Message;
                return;
            }
            finally
            {
                GC.Collect();
            }
        }

        public static string getDefaultReportLayoutCode( string addonName, string addonFormType, out string errorText)
        {
            errorText = null;

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                string query = "SELECT \"RTYP\".\"DEFLT_REP\" " +
                    "FROM \"RTYP\" AS \"RTYP\" " +
                    "WHERE \"RTYP\".\"FRM_TYPE\" = '" + addonFormType + "' AND \"RTYP\".\"ADD_NAME\" = '" + addonName + "'";

                oRecordSet.DoQuery(query);
                while (!oRecordSet.EoF)
                {
                    return oRecordSet.Fields.Item("DEFLT_REP").Value.ToString();
                }
                return null;
            }
            catch (Exception ex)
            {
                int errCode;
                string errMsg;
                Program.oCompany.GetLastError(out errCode, out errMsg);
                errorText = BDOSResources.getTranslate("ErrorDescription") + " : " + errMsg + "! " + BDOSResources.getTranslate("Code") + " : " + errCode + "! " + BDOSResources.getTranslate("OtherInfo") + " : " + ex.Message;
                return null;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                GC.Collect();
            }
        }

        public static string getReportlayoutCode( string docName, string typeCode, out string errorText)
        {
            errorText = null;

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                string query = "SELECT \"RDOC\".\"DocCode\" " +
                    "FROM \"RDOC\" AS \"RDOC\" " +
                    "WHERE \"RDOC\".\"DocName\" = '" + docName + "' AND \"RDOC\".\"TypeCode\" = '" + typeCode + "'";

                oRecordSet.DoQuery(query);
                while (!oRecordSet.EoF)
                {
                    return oRecordSet.Fields.Item("DocCode").Value.ToString();
                }
                return null;
            }
            catch (Exception ex)
            {
                int errCode;
                string errMsg;
                Program.oCompany.GetLastError(out errCode, out errMsg);
                errorText = BDOSResources.getTranslate("ErrorDescription") + " : " + errMsg + "! " + BDOSResources.getTranslate("Code") + " : " + errCode + "! " + BDOSResources.getTranslate("OtherInfo") + " : " + ex.Message;
                return null;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                GC.Collect();
            }
        }

        public static void printCrystalReport( string layoutCode, int docEntry, out string errorText)
        {
            errorText = null;

            SAPbobsCOM.CompanyService oCmpSrv;
            SAPbobsCOM.ReportLayoutsService oReportLayoutService;
            SAPbobsCOM.ReportLayoutPrintParams oPrintParam;
            oCmpSrv = Program.oCompany.GetCompanyService();
            oReportLayoutService = oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.ReportLayoutsService);

            try
            {
                //'SETUP THE REPORT
                oPrintParam = oReportLayoutService.GetDataInterface(SAPbobsCOM.ReportLayoutsServiceDataInterfaces.rlsdiReportLayoutPrintParams);
                oPrintParam.LayoutCode = layoutCode;
                oPrintParam.DocEntry = docEntry;

                //'PRINT WITH DEFAULT SETTINGS
                oReportLayoutService.Print(oPrintParam);
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
        }
    }
}
