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
    class BDOSTaxAnalysissSent
    {
        public static bool RSDataImported = false;

        public static void createUDO( out string errorText)
        {
            errorText = null;

            string tableName = "BDOSTXANS";
            string description = "Tax Sent (Analysis)";

            SAPbobsCOM.UserObjectsMD oUserObjectsMD = null;
            oUserObjectsMD = (SAPbobsCOM.UserObjectsMD)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectsMD);

            int result = UDO.addUserTable( tableName, description, SAPbobsCOM.BoUTBTableType.bott_NoObject, out errorText);

            if (result != 0)
            {
                return;
            }

            Dictionary<string, object> fieldskeysMap = new Dictionary<string, object>();

            fieldskeysMap = new Dictionary<string, object>(); // overhead_no
            fieldskeysMap.Add("Name", "Overheadno");
            fieldskeysMap.Add("TableName", "BDOSTXANS");
            fieldskeysMap.Add("Description", "Overheadno");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // User
            fieldskeysMap.Add("Name", "User");
            fieldskeysMap.Add("TableName", "BDOSTXANS");
            fieldskeysMap.Add("Description", "User");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields( fieldskeysMap, out errorText);
            fieldskeysMap = new Dictionary<string, object>(); // ID
            fieldskeysMap.Add("Name", "ID");
            fieldskeysMap.Add("TableName", "BDOSTXANS");
            fieldskeysMap.Add("Description", "ID");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // SELLER_UN_ID
            fieldskeysMap.Add("Name", "SELLERUNID");
            fieldskeysMap.Add("TableName", "BDOSTXANS");
            fieldskeysMap.Add("Description", "SELLER_UN_ID");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // SEQ_NUM_B
            fieldskeysMap.Add("Name", "SEQNUMB");
            fieldskeysMap.Add("TableName", "BDOSTXANS");
            fieldskeysMap.Add("Description", "SEQ_NUM_B");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // STATUS
            fieldskeysMap.Add("Name", "STATUS");
            fieldskeysMap.Add("TableName", "BDOSTXANS");
            fieldskeysMap.Add("Description", "STATUS");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // WAS_REF
            fieldskeysMap.Add("Name", "WASREF");
            fieldskeysMap.Add("TableName", "BDOSTXANS");
            fieldskeysMap.Add("Description", "WAS_REF");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // F_SERIES
            fieldskeysMap.Add("Name", "FSERIES");
            fieldskeysMap.Add("TableName", "BDOSTXANS");
            fieldskeysMap.Add("Description", "F_SERIES");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // F_NUMBER
            fieldskeysMap.Add("Name", "FNUMBER");
            fieldskeysMap.Add("TableName", "BDOSTXANS");
            fieldskeysMap.Add("Description", "F_NUMBER");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields( fieldskeysMap, out errorText);
            fieldskeysMap = new Dictionary<string, object>(); // REG_DT
            fieldskeysMap.Add("Name", "REGDT");
            fieldskeysMap.Add("TableName", "BDOSTXANS");
            fieldskeysMap.Add("Description", "REG_DT");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Date);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // OPERATION_DT
            fieldskeysMap.Add("Name", "OPERATDT");
            fieldskeysMap.Add("TableName", "BDOSTXANS");
            fieldskeysMap.Add("Description", "OPERATION_DT");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Date);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // S_USER_ID
            fieldskeysMap.Add("Name", "SUSERID");
            fieldskeysMap.Add("TableName", "BDOSTXANS");
            fieldskeysMap.Add("Description", "S_USER_ID");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // B_S_USER_ID
            fieldskeysMap.Add("Name", "BSUSERID");
            fieldskeysMap.Add("TableName", "BDOSTXANS");
            fieldskeysMap.Add("Description", "B_S_USER_ID");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // DOC_MOS_NOM_B
            fieldskeysMap.Add("Name", "DOCMOSNOMB");
            fieldskeysMap.Add("TableName", "BDOSTXANS");
            fieldskeysMap.Add("Description", "DOC_MOS_NOM_B");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // SA_IDENT_NO
            fieldskeysMap.Add("Name", "SAIDENTNO");
            fieldskeysMap.Add("TableName", "BDOSTXANS");
            fieldskeysMap.Add("Description", "SA_IDENT_NO");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // ORG_NAME
            fieldskeysMap.Add("Name", "ORGNAME");
            fieldskeysMap.Add("TableName", "BDOSTXANS");
            fieldskeysMap.Add("Description", "ORG_NAME");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // NOTES
            fieldskeysMap.Add("Name", "NOTES");
            fieldskeysMap.Add("TableName", "BDOSTXANS");
            fieldskeysMap.Add("Description", "NOTES");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // TANXA
            fieldskeysMap.Add("Name", "TANXA");
            fieldskeysMap.Add("TableName", "BDOSTXANS");
            fieldskeysMap.Add("Description", "TANXA");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // VAT
            fieldskeysMap.Add("Name", "VAT");
            fieldskeysMap.Add("TableName", "BDOSTXANS");
            fieldskeysMap.Add("Description", "VAT");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // K_ID
            fieldskeysMap.Add("Name", "KID");
            fieldskeysMap.Add("TableName", "BDOSTXANS");
            fieldskeysMap.Add("Description", "K_ID");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // AGREE_DATE
            fieldskeysMap.Add("Name", "AGREEDATE");
            fieldskeysMap.Add("TableName", "BDOSTXANS");
            fieldskeysMap.Add("Description", "AGREE_DATE");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Date);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // AGREE_S_USER_ID
            fieldskeysMap.Add("Name", "AGREESUSID");
            fieldskeysMap.Add("TableName", "BDOSTXANS");
            fieldskeysMap.Add("Description", "AGREE_S_USER_ID");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // REF_DATE
            fieldskeysMap.Add("Name", "REFDATE");
            fieldskeysMap.Add("TableName", "BDOSTXANS");
            fieldskeysMap.Add("Description", "REF_DATE");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Date);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // REF_S_USER_ID
            fieldskeysMap.Add("Name", "REFUSID");
            fieldskeysMap.Add("TableName", "BDOSTXANS");
            fieldskeysMap.Add("Description", "REF_S_USER_ID");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // Wb_Corredted_Tx
            fieldskeysMap.Add("Name", "WbCorrTx");
            fieldskeysMap.Add("TableName", "BDOSTXANS");
            fieldskeysMap.Add("Description", "Wb_Corredted_Tx");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

        }

        public static void fillUserCode( SAPbouiCOM.Form oForm)
        {
            oForm.Items.Item("1000033").Specific.Value = Program.oCompany.UserName;
            oForm.Items.Item("1000003").Click();
            oForm.Items.Item("1000033").Enabled = false;
        }

        public static void addRecord(  SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            SAPbobsCOM.UserTable oUserTable = Program.oCompany.UserTables.Item("BDOSTXANS");

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            SAPbobsCOM.Recordset oRecordSetinsert = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string insertquery = "";

            //არსებული ჩანაწერების წაშლა ცხრილიდან
            string clearquery = @"DELETE FROM ""@BDOSTXANS"" WHERE ""U_User"" = '" + Program.oCompany.UserName + @"'";
            oRecordSet.DoQuery(clearquery);

            //int keyCode = 0;
            //string maxquery = @"SELECT ""Code"" FROM ""@BDOSTXANS"" ORDER BY ""Code"" DESC";
            //oRecordSet.DoQuery(maxquery);
            //if (!oRecordSet.EoF)
            //{
            //    keyCode = Convert.ToInt32(oRecordSet.Fields.Item("Code").Value.ToString());
            //}

            DateTime startDate;
            string startDateStr = oForm.Items.Item("1000003").Specific.Value;
            DateTime BeginDate = new DateTime(1, 1, 1);

            if (DateTime.TryParseExact(startDateStr, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture, DateTimeStyles.None, out startDate))
            {
                BeginDate = startDate;
            }

            DateTime endDate;
            string endDateStr = oForm.Items.Item("1000009").Specific.Value;
            DateTime EndDate = DateTime.Today;

            if (DateTime.TryParseExact(endDateStr, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture, DateTimeStyles.None, out endDate))
            {
                EndDate = endDate;
            }
            endDate = endDate.AddDays(1).AddSeconds(-1);
            string invoice_no = oForm.Items.Item("1000021").Specific.Value;

            Dictionary<string, string> rsSettings = CompanyDetails.getRSSettings( out errorText);
            if (errorText != null)
            {
                Program.uiApp.StatusBar.SetSystemMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return;
            }

            string su = rsSettings["SU"];
            string sp = rsSettings["SP"];

            TaxInvoice oTaxInvoice = new TaxInvoice(rsSettings["ProtocolType"]);

            bool chek_service_user = oTaxInvoice.check_usr( su, sp, out errorText);
            if (chek_service_user == false)
            {
                errorText = BDOSResources.getTranslate("ServiceUserPasswordNotCorrect");
                Program.uiApp.StatusBar.SetSystemMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return;
            }

            string sa_identNum = "";
            DateTime startDateOp = new DateTime(1, 1, 1);
            DateTime endDateOp = new DateTime(DateTime.Today.Year, DateTime.Today.Month, DateTime.DaysInMonth(DateTime.Today.Year, DateTime.Today.Month));
            endDateOp = endDateOp.AddDays(1).AddSeconds(-1);
            DataTable TaxDataTable = oTaxInvoice.get_seller_invoices(startDate, endDate, startDateOp, endDateOp, invoice_no, sa_identNum, "", "", out errorText);

            try
            {
                Parallel.ForEach(TaxDataTable.AsEnumerable(), TaxDataRow =>
                {
                    DataTable k_invoiceTableLines = new DataTable();
                    decimal TANXA_First = 0;
                    decimal VAT_First = 0;
                    bool k_invoiceTableLines_False = true;

                    string k_ID = TaxDataRow["K_ID"].ToString();
                    if (k_ID != "-1")
                    {
                        k_invoiceTableLines = oTaxInvoice.get_ntos_invoices_inv_nos(Convert.ToInt64(k_ID));
                        if (k_invoiceTableLines != null)
                        {
                            k_invoiceTableLines_False = false;
                        }

                        DataTable k_taxDataTable = oTaxInvoice.get_invoice_desc(Convert.ToInt64(k_ID));
                        Parallel.ForEach(k_taxDataTable.AsEnumerable(), k_taxDeclRow =>
                        {
                            if (k_taxDataTable.Columns.Contains("full_amount"))
                            {

                                TANXA_First = TANXA_First + Convert.ToDecimal(k_taxDeclRow["full_amount"], CultureInfo.InvariantCulture); //თანხა დღგ-ის და აქციზის ჩათვლით
                            }

                            if (k_taxDataTable.Columns.Contains("drg_amount"))
                            {
                                decimal drg_amount = Convert.ToDecimal(k_taxDeclRow["drg_amount"], CultureInfo.InvariantCulture); //დღგ
                                if (drg_amount > 0)
                                {
                                    VAT_First = VAT_First + drg_amount;
                                }
                            }
                        });
                    }

                    string inv_ID = TaxDataRow["ID"].ToString();
                    DataTable invoiceTableLines = oTaxInvoice.get_ntos_invoices_inv_nos(Convert.ToInt64(inv_ID));
                    if (invoiceTableLines != null || k_invoiceTableLines_False == false)
                    {
                        fillValues( ref insertquery, TaxDataRow, "Error getting Waybill numbers", TANXA_First, VAT_First, false);
                    }
                    else if (invoiceTableLines.Rows.Count > 0)
                    {
                        Parallel.ForEach(invoiceTableLines.AsEnumerable(), currDataRow =>
                        {
                            string overhead_no = currDataRow["overhead_no"].ToString();
                            DataRow[] foundRows = k_invoiceTableLines.Select(@"overhead_no = '" + overhead_no + "'");
                            bool WbCorrTx = (foundRows.Length > 0);

                            fillValues( ref insertquery, TaxDataRow, overhead_no, TANXA_First, VAT_First, WbCorrTx);
                        });
                    }
                    else
                    {
                        fillValues( ref insertquery, TaxDataRow, "", TANXA_First, VAT_First, false);
                    }

                    try
                    {
                        oRecordSetinsert.DoQuery(insertquery);
                    }
                    catch (Exception ex)
                    {
                        int errCode;
                        string errMsg;

                        Program.oCompany.GetLastError(out errCode, out errMsg);
                        string erText = BDOSResources.getTranslate("ErrorDescription") + " : " + errMsg + "! " + BDOSResources.getTranslate("Code") + " : " + errCode + "! " + BDOSResources.getTranslate("OtherInfo") + " : " + ex.Message;
                    }
                });
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
                Marshal.ReleaseComObject(oUserTable);
                oUserTable = null;
                RSDataImported = true;
                oForm.Items.Item("1").Click();
            }
        }

        public static void fillValues_OLD( SAPbobsCOM.UserTable oUserTable, DataRow TaxDataRow, string overhead_no, decimal TANXA_First, decimal Vat_First, bool WbCorrTx, out string errorText)
        {
            errorText = null;

            SAPbobsCOM.SBObob oSBOBob = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);

            oUserTable.UserFields.Fields.Item("U_Overheadno").Value = overhead_no;
            oUserTable.UserFields.Fields.Item("U_User").Value = Program.oCompany.UserName;
            oUserTable.UserFields.Fields.Item("U_ID").Value = TaxDataRow["ID"].ToString();
            //try
            //{
            //    oUserTable.UserFields.Fields.Item("U_SELLERUNID").Value = TaxDataRow["SELLER_UN_ID"].ToString();
            //}
            //catch { }

            //try
            //{
            //    oUserTable.UserFields.Fields.Item("U_SEQNUMB").Value = TaxDataRow["SEQ_NUM_B"].ToString();
            //}
            //catch { }

            string K_ID = TaxDataRow["K_ID"].ToString(); // კორექტირების ანგარიშ-ფაქტურის ID
            string WAS_REF = TaxDataRow["WAS_REF"].ToString(); // უარყოფილი მეორე მხარის მიერ 0 - არა 1 - კი
            bool corrInv = K_ID != "-1"; //თუ არის კორექტირების ა/ფ
            bool refInv = WAS_REF == "1"; //თუ არის უარყოფილი ა/ფ

            string statusRS = TaxDataRow["STATUS"].ToString();
            oUserTable.UserFields.Fields.Item("U_STATUS").Value = BDO_TaxInvoiceSent.getStatusValueByStatusNumber(statusRS, corrInv, refInv);
            //oUserTable.UserFields.Fields.Item("U_WASREF").Value = WAS_REF;
            oUserTable.UserFields.Fields.Item("U_FSERIES").Value = TaxDataRow["F_SERIES"].ToString();
            oUserTable.UserFields.Fields.Item("U_FNUMBER").Value = TaxDataRow["F_NUMBER"].ToString();

            try
            {
                DateTime REG_DT = new DateTime();
                REG_DT = DateTime.TryParse(TaxDataRow["REG_DT"].ToString(), out REG_DT) == false ? new DateTime() : REG_DT; // რეგისტრაციის თარიღი
                oUserTable.UserFields.Fields.Item("U_REGDT").Value = REG_DT;
            }
            catch { }

            try
            {
                DateTime OPERATION_DT = new DateTime();
                OPERATION_DT = DateTime.TryParse(TaxDataRow["OPERATION_DT"].ToString(), out OPERATION_DT) == false ? new DateTime() : OPERATION_DT; // ოპერაციის განხორციელების თარიღი
                oUserTable.UserFields.Fields.Item("U_OPERATDT").Value = OPERATION_DT;
            }
            catch { }

            //oUserTable.UserFields.Fields.Item("U_SUSERID").Value = TaxDataRow["S_USER_ID"].ToString();
            //try
            //{
            //    oUserTable.UserFields.Fields.Item("U_BSUSERID").Value = TaxDataRow["B_S_USER_ID"].ToString();
            //}
            //catch { }

            //try
            //{
            //    oUserTable.UserFields.Fields.Item("U_DOCMOSNOMB").Value = TaxDataRow["DOC_MOS_NOM_B"].ToString();
            //}
            //catch { }

            oUserTable.UserFields.Fields.Item("U_SAIDENTNO").Value = TaxDataRow["SA_IDENT_NO"].ToString();
            //oUserTable.UserFields.Fields.Item("U_ORGNAME").Value = TaxDataRow["ORG_NAME"].ToString();
            //oUserTable.UserFields.Fields.Item("U_NOTES").Value = TaxDataRow["NOTES"].ToString();

            string TANXA = TaxDataRow["TANXA"].ToString();
            if (TANXA != "")
            {
                decimal Amount = Convert.ToDecimal(TaxDataRow["TANXA"], CultureInfo.InvariantCulture) - TANXA_First;
                oUserTable.UserFields.Fields.Item("U_TANXA").Value = Convert.ToDouble(Amount, System.Globalization.CultureInfo.InvariantCulture);
            }

            string VAT = TaxDataRow["VAT"].ToString();
            if (VAT != "")
            {
                decimal VatAmount = Convert.ToDecimal(TaxDataRow["VAT"], CultureInfo.InvariantCulture) - Vat_First;
                oUserTable.UserFields.Fields.Item("U_VAT").Value = Convert.ToDouble(VatAmount, System.Globalization.CultureInfo.InvariantCulture);
            }

            //oUserTable.UserFields.Fields.Item("U_KID").Value = K_ID;

            //try
            //{
            //    DateTime AGREE_DATE = new DateTime();
            //    AGREE_DATE = DateTime.TryParse(TaxDataRow["AGREE_DATE"].ToString(), out AGREE_DATE) == false ? new DateTime() : AGREE_DATE; // დადასტურების თარიღი
            //    oUserTable.UserFields.Fields.Item("U_AGREEDATE").Value = AGREE_DATE;
            //}
            //catch { }

            //oUserTable.UserFields.Fields.Item("U_AGREESUSID").Value = TaxDataRow["AGREE_S_USER_ID"].ToString();

            //try
            //{
            //    DateTime REF_DATE = new DateTime();
            //    REF_DATE = DateTime.TryParse(TaxDataRow["REF_DATE"].ToString(), out REF_DATE) == false ? new DateTime() : REF_DATE; // უარყოფის თარიღი
            //    oUserTable.UserFields.Fields.Item("U_REFDATE").Value = REF_DATE;
            //}
            //catch { }

            //oUserTable.UserFields.Fields.Item("U_REFUSID").Value = TaxDataRow["REF_S_USER_ID"].ToString();
            //oUserTable.UserFields.Fields.Item("U_WbCorrTx").Value = WbCorrTx.ToString();
            int returnCode = oUserTable.Add();

            if (returnCode != 0)
            {
                int errCode;
                string errMsg;

                Program.oCompany.GetLastError(out errCode, out errMsg);
                errorText = "Error description : " + errMsg + "! Code : " + errCode;
            }
        }

        public static void fillValues( ref string insertquery, DataRow TaxDataRow, string overhead_no, decimal TANXA_First, decimal Vat_First, bool WbCorrTx)
        {
            string statusRS = TaxDataRow["STATUS"].ToString();
            string K_ID = TaxDataRow["K_ID"].ToString(); // კორექტირების ანგარიშ-ფაქტურის ID
            string WAS_REF = TaxDataRow["WAS_REF"].ToString(); // უარყოფილი მეორე მხარის მიერ 0 - არა 1 - კი
            bool corrInv = K_ID != "-1"; //თუ არის კორექტირების ა/ფ
            bool refInv = WAS_REF == "1"; //თუ არის უარყოფილი ა/ფ

            DateTime REG_DT = new DateTime();
            DateTime OPERATION_DT = new DateTime();
            decimal Amount = 0;
            decimal VatAmount = 0;

            try
            {
                REG_DT = DateTime.TryParse(TaxDataRow["REG_DT"].ToString(), out REG_DT) == false ? new DateTime() : REG_DT; // რეგისტრაციის თარიღი
            }
            catch { }

            try
            {
                OPERATION_DT = DateTime.TryParse(TaxDataRow["OPERATION_DT"].ToString(), out OPERATION_DT) == false ? new DateTime() : OPERATION_DT; // ოპერაციის განხორციელების თარიღი
            }
            catch { }

            string TANXA = TaxDataRow["TANXA"].ToString();
            if (TANXA != "")
            {
                Amount = Convert.ToDecimal(TaxDataRow["TANXA"], CultureInfo.InvariantCulture) - TANXA_First;
            }

            string VAT = TaxDataRow["VAT"].ToString();
            if (VAT != "")
            {
                VatAmount = Convert.ToDecimal(TaxDataRow["VAT"], CultureInfo.InvariantCulture) - Vat_First;
            }

            StringBuilder Sbuilder = new StringBuilder();

            Sbuilder.Append(@"INSERT INTO ""@BDOSTXANS""  (""Code"", 
                                                        ""Name"",
                                                        ""U_Overheadno"", 
                                                        ""U_User"", 
                                                        ""U_ID"", 
                                                        ""U_STATUS"", 
                                                        ""U_FSERIES"", 
                                                        ""U_FNUMBER"", 
                                                        ""U_REGDT"",
                                                        ""U_OPERATDT"",                                                        
                                                        ""U_SAIDENTNO"",
                                                        ""U_TANXA"",
                                                        ""U_VAT"")
                                                        VALUES (");

            string KeyCode = System.Guid.NewGuid().ToString();
            Sbuilder.Append("'");
            Sbuilder.Append(KeyCode);
            Sbuilder.Append("',");

            Sbuilder.Append("'");
            Sbuilder.Append(KeyCode);
            Sbuilder.Append("',");

            Sbuilder.Append("'");
            Sbuilder.Append(overhead_no);
            Sbuilder.Append("',");

            Sbuilder.Append("'");
            CommonFunctions.AppendXML(Sbuilder, Program.oCompany.UserName);
            Sbuilder.Append("',");

            Sbuilder.Append("'");
            CommonFunctions.AppendXML(Sbuilder, TaxDataRow["ID"].ToString());
            Sbuilder.Append("',");

            Sbuilder.Append("'");
            CommonFunctions.AppendXML(Sbuilder, BDO_TaxInvoiceSent.getStatusValueByStatusNumber(statusRS, corrInv, refInv));
            Sbuilder.Append("',");

            Sbuilder.Append("N'");
            CommonFunctions.AppendXML(Sbuilder, TaxDataRow["F_SERIES"].ToString());
            Sbuilder.Append("',");

            Sbuilder.Append("'");
            CommonFunctions.AppendXML(Sbuilder, TaxDataRow["F_NUMBER"].ToString());
            Sbuilder.Append("',");

            Sbuilder.Append("'");
            CommonFunctions.AppendXML(Sbuilder, REG_DT.ToString("yyyyMMdd"));
            Sbuilder.Append("',");

            Sbuilder.Append("'");
            CommonFunctions.AppendXML(Sbuilder, OPERATION_DT.ToString("yyyyMMdd"));
            Sbuilder.Append("',");

            Sbuilder.Append("'");
            CommonFunctions.AppendXML(Sbuilder, TaxDataRow["SA_IDENT_NO"].ToString());
            Sbuilder.Append("',");

            Sbuilder.Append(Amount.ToString(CultureInfo.InvariantCulture));
            Sbuilder.Append(",");

            Sbuilder.Append(VatAmount.ToString(CultureInfo.InvariantCulture));
            Sbuilder.Append(");");

            insertquery = Sbuilder.ToString();

        }

        public static void uiApp_ItemEvent(  string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            if (pVal.ItemChanged & pVal.BeforeAction == false)
            {
                RSDataImported = false;
            }

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (oForm.Title.Contains("Tax Invoice Sent Analysis") != true)
                {
                    return;
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD & pVal.BeforeAction == true)
                {
                    fillUserCode( oForm);
                }

                if (pVal.ItemUID == "1" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK & pVal.BeforeAction == true)
                {
                    if (RSDataImported == false)
                    {
                        addRecord( oForm, out errorText);
                    }
                }

            }
        }

    }
}
