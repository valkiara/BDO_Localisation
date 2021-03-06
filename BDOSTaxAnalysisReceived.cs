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
    class BDOSTaxAnalysisReceived
    {
        public static void createUDO(out string errorText)
        {
            errorText = null;

            string tableName = "BDOSTXANR";
            string description = "Tax Received (Analysis)";

            SAPbobsCOM.UserObjectsMD oUserObjectsMD = null;
            oUserObjectsMD = (SAPbobsCOM.UserObjectsMD)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectsMD);

            int result = UDO.addUserTable(tableName, description, SAPbobsCOM.BoUTBTableType.bott_NoObject, out errorText);

            if (result != 0)
            {
                return;
            }

            Dictionary<string, object> fieldskeysMap = new Dictionary<string, object>();

            fieldskeysMap = new Dictionary<string, object>(); // overhead_no
            fieldskeysMap.Add("Name", "Overheadno");
            fieldskeysMap.Add("TableName", "BDOSTXANR");
            fieldskeysMap.Add("Description", "Overheadno");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // User
            fieldskeysMap.Add("Name", "User");
            fieldskeysMap.Add("TableName", "BDOSTXANR");
            fieldskeysMap.Add("Description", "User");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // ID
            fieldskeysMap.Add("Name", "ID");
            fieldskeysMap.Add("TableName", "BDOSTXANR");
            fieldskeysMap.Add("Description", "ID");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // SELLER_UN_ID
            fieldskeysMap.Add("Name", "SELLERUNID");
            fieldskeysMap.Add("TableName", "BDOSTXANR");
            fieldskeysMap.Add("Description", "SELLER_UN_ID");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // SEQ_NUM_B
            fieldskeysMap.Add("Name", "SEQNUMB");
            fieldskeysMap.Add("TableName", "BDOSTXANR");
            fieldskeysMap.Add("Description", "SEQ_NUM_B");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // STATUS
            fieldskeysMap.Add("Name", "STATUS");
            fieldskeysMap.Add("TableName", "BDOSTXANR");
            fieldskeysMap.Add("Description", "STATUS");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // WAS_REF
            fieldskeysMap.Add("Name", "WASREF");
            fieldskeysMap.Add("TableName", "BDOSTXANR");
            fieldskeysMap.Add("Description", "WAS_REF");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // F_SERIES
            fieldskeysMap.Add("Name", "FSERIES");
            fieldskeysMap.Add("TableName", "BDOSTXANR");
            fieldskeysMap.Add("Description", "F_SERIES");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // F_NUMBER
            fieldskeysMap.Add("Name", "FNUMBER");
            fieldskeysMap.Add("TableName", "BDOSTXANR");
            fieldskeysMap.Add("Description", "F_NUMBER");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);
            fieldskeysMap = new Dictionary<string, object>(); // REG_DT
            fieldskeysMap.Add("Name", "REGDT");
            fieldskeysMap.Add("TableName", "BDOSTXANR");
            fieldskeysMap.Add("Description", "REG_DT");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Date);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // OPERATION_DT
            fieldskeysMap.Add("Name", "OPERATDT");
            fieldskeysMap.Add("TableName", "BDOSTXANR");
            fieldskeysMap.Add("Description", "OPERATION_DT");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Date);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // S_USER_ID
            fieldskeysMap.Add("Name", "SUSERID");
            fieldskeysMap.Add("TableName", "BDOSTXANR");
            fieldskeysMap.Add("Description", "S_USER_ID");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // B_S_USER_ID
            fieldskeysMap.Add("Name", "BSUSERID");
            fieldskeysMap.Add("TableName", "BDOSTXANR");
            fieldskeysMap.Add("Description", "B_S_USER_ID");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // DOC_MOS_NOM_B
            fieldskeysMap.Add("Name", "DOCMOSNOMB");
            fieldskeysMap.Add("TableName", "BDOSTXANR");
            fieldskeysMap.Add("Description", "DOC_MOS_NOM_B");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // SA_IDENT_NO
            fieldskeysMap.Add("Name", "SAIDENTNO");
            fieldskeysMap.Add("TableName", "BDOSTXANR");
            fieldskeysMap.Add("Description", "SA_IDENT_NO");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // ORG_NAME
            fieldskeysMap.Add("Name", "ORGNAME");
            fieldskeysMap.Add("TableName", "BDOSTXANR");
            fieldskeysMap.Add("Description", "ORG_NAME");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // NOTES
            fieldskeysMap.Add("Name", "NOTES");
            fieldskeysMap.Add("TableName", "BDOSTXANR");
            fieldskeysMap.Add("Description", "NOTES");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // TANXA
            fieldskeysMap.Add("Name", "TANXA");
            fieldskeysMap.Add("TableName", "BDOSTXANR");
            fieldskeysMap.Add("Description", "TANXA");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // VAT
            fieldskeysMap.Add("Name", "VAT");
            fieldskeysMap.Add("TableName", "BDOSTXANR");
            fieldskeysMap.Add("Description", "VAT");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // K_ID
            fieldskeysMap.Add("Name", "KID");
            fieldskeysMap.Add("TableName", "BDOSTXANR");
            fieldskeysMap.Add("Description", "K_ID");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // AGREE_DATE
            fieldskeysMap.Add("Name", "AGREEDATE");
            fieldskeysMap.Add("TableName", "BDOSTXANR");
            fieldskeysMap.Add("Description", "AGREE_DATE");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Date);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // AGREE_S_USER_ID
            fieldskeysMap.Add("Name", "AGREESUSID");
            fieldskeysMap.Add("TableName", "BDOSTXANR");
            fieldskeysMap.Add("Description", "AGREE_S_USER_ID");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // REF_DATE
            fieldskeysMap.Add("Name", "REFDATE");
            fieldskeysMap.Add("TableName", "BDOSTXANR");
            fieldskeysMap.Add("Description", "REF_DATE");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Date);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // REF_S_USER_ID
            fieldskeysMap.Add("Name", "REFUSID");
            fieldskeysMap.Add("TableName", "BDOSTXANR");
            fieldskeysMap.Add("Description", "REF_S_USER_ID");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);


        }

        public static void fillUserCode(SAPbouiCOM.Form oForm)
        {
            oForm.Items.Item("1000033").Specific.Value = Program.oCompany.UserName;
            oForm.Items.Item("1000003").Click();
            oForm.Items.Item("1000033").Enabled = false;
        }

        public static void addRecord(SAPbouiCOM.Form oForm)
        {
            SAPbobsCOM.UserTable oUserTable = Program.oCompany.UserTables.Item("BDOSTXANR");

            //არსებული ჩანაწერების წაშლა ცხრილიდან
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string clearQuery = @"DELETE FROM ""@BDOSTXANR"" WHERE ""U_User"" = '" + Program.oCompany.UserName + @"'";
            oRecordSet.DoQuery(clearQuery);

            //int keyCode = 0;
            //string maxquery = @"SELECT ""Code"" FROM ""@BDOSTXANR"" ORDER BY ""Code"" DESC";
            //oRecordSet.DoQuery(maxquery);
            //if (!oRecordSet.EoF)
            //{
            //    keyCode = Convert.ToInt32(oRecordSet.Fields.Item("Code").Value.ToString());
            //}

            DateTime startDate;
            string startDateStr = oForm.Items.Item("1000003").Specific.Value;
            startDate = DateTime.TryParseExact(startDateStr, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out startDate) ? startDate : DateTime.MinValue;

            DateTime endDate;
            string endDateStr = oForm.Items.Item("1000009").Specific.Value;
            endDate = DateTime.TryParseExact(endDateStr, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out endDate) ? endDate : DateTime.MinValue;
            endDate = endDate.AddDays(1).AddSeconds(-1);

            string invoice_no = oForm.Items.Item("1000021").Specific.Value;

            string errorText;
            Dictionary<string, string> rsSettings = CompanyDetails.getRSSettings(out errorText);
            if (errorText != null)
            {
                Program.uiApp.StatusBar.SetSystemMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short);
                return;
            }

            string su = rsSettings["SU"];
            string sp = rsSettings["SP"];

            TaxInvoice oTaxInvoice = new TaxInvoice(rsSettings["ProtocolType"]);

            bool chek_service_user = oTaxInvoice.check_usr(su, sp, out errorText);
            if (chek_service_user == false)
            {
                errorText = BDOSResources.getTranslate("ServiceUserPasswordNotCorrect");
                Program.uiApp.StatusBar.SetSystemMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short);
                return;
            }

            string sa_identNum = "";
            DateTime startDateOp = new DateTime(1, 1, 1);
            DateTime endDateOp = new DateTime(DateTime.Today.Year, DateTime.Today.Month, DateTime.DaysInMonth(DateTime.Today.Year, DateTime.Today.Month));
            endDateOp = endDateOp.AddDays(1).AddSeconds(-1);

            DataTable TaxDataTable = oTaxInvoice.get_buyer_invoices(startDate, endDate, startDateOp, endDateOp, invoice_no, sa_identNum, "", "", out errorText);

            SAPbobsCOM.Recordset oRecordSetInsert = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                string insertQuery = string.Empty;

                foreach (DataRow TaxDataRow in TaxDataTable.AsEnumerable())
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
                            k_invoiceTableLines_False = false;

                        DataTable k_taxDataTable = oTaxInvoice.get_invoice_desc(Convert.ToInt64(k_ID));

                        foreach (DataRow k_taxDeclRow in k_taxDataTable.AsEnumerable())
                        {
                            if (k_taxDataTable.Columns.Contains("full_amount"))
                            {
                                TANXA_First += Convert.ToDecimal(k_taxDeclRow["full_amount"], CultureInfo.InvariantCulture); //თანხა დღგ-ის და აქციზის ჩათვლით
                            }

                            if (k_taxDataTable.Columns.Contains("drg_amount"))
                            {
                                decimal drg_amount = Convert.ToDecimal(k_taxDeclRow["drg_amount"], CultureInfo.InvariantCulture); //დღგ
                                if (drg_amount > 0)
                                {
                                    VAT_First += drg_amount;
                                }
                            }
                        };
                    }

                    string inv_ID = TaxDataRow["ID"].ToString();
                    DataTable invoiceTableLines = oTaxInvoice.get_ntos_invoices_inv_nos(Convert.ToInt64(inv_ID));
                    if (invoiceTableLines != null || k_invoiceTableLines_False == false)
                    {
                        fillValues(ref insertQuery, TaxDataRow, "Error getting Waybill numbers", TANXA_First, VAT_First);
                    }
                    else if (invoiceTableLines.Rows.Count > 0)
                    {
                        int quantityAddRows = 0;

                        foreach (DataRow currDataRow in invoiceTableLines.AsEnumerable())
                        {
                            string overhead_no = currDataRow["overhead_no"].ToString();
                            DataRow[] foundRows = k_invoiceTableLines.Select(@"overhead_no = '" + overhead_no + "'");
                            if (foundRows.Length > 0)
                            {
                                fillValues(ref insertQuery, TaxDataRow, overhead_no, TANXA_First, VAT_First);
                                quantityAddRows++;
                            }
                        };

                        if (quantityAddRows == 0)
                        {
                            fillValues(ref insertQuery, TaxDataRow, "", TANXA_First, VAT_First);
                        }
                    }
                    else
                    {
                        fillValues(ref insertQuery, TaxDataRow, "", TANXA_First, VAT_First);
                    }

                    try
                    {
                        oRecordSetInsert.DoQuery(insertQuery);  
                    }
                    catch (Exception ex)
                    {
                        Program.uiApp.StatusBar.SetSystemMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short);
                        return;
                    }
                }
                oForm.Items.Item("1").Click();
            }
            catch (Exception ex)
            {
                Program.uiApp.StatusBar.SetSystemMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short);
                return;
            }
            finally
            {
                Marshal.ReleaseComObject(oRecordSet);
                Marshal.ReleaseComObject(oRecordSetInsert);
                Marshal.ReleaseComObject(oUserTable);
            }
        }

        public static void fillValues_OLD(SAPbobsCOM.UserTable oUserTable, DataRow TaxDataRow, string overhead_no, decimal TANXA_First, decimal Vat_First, out string errorText)
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

            string statusRS = TaxDataRow["STATUS"].ToString();
            oUserTable.UserFields.Fields.Item("U_STATUS").Value = BDO_TaxInvoiceReceived.getStatusValueByStatusNumber(statusRS);
            oUserTable.UserFields.Fields.Item("U_WASREF").Value = TaxDataRow["WAS_REF"].ToString();
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

            oUserTable.UserFields.Fields.Item("U_SUSERID").Value = TaxDataRow["S_USER_ID"].ToString();

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
            oUserTable.UserFields.Fields.Item("U_ORGNAME").Value = TaxDataRow["ORG_NAME"].ToString();
            oUserTable.UserFields.Fields.Item("U_NOTES").Value = TaxDataRow["NOTES"].ToString();

            string TANXA = TaxDataRow["TANXA"].ToString();
            if (TANXA != "")
            {
                decimal Amount = Convert.ToDecimal(TaxDataRow["TANXA"], CultureInfo.InvariantCulture) - TANXA_First;
                oUserTable.UserFields.Fields.Item("U_TANXA").Value = Convert.ToDouble(Amount, CultureInfo.InvariantCulture);
            }

            string VAT = TaxDataRow["VAT"].ToString();
            if (VAT != "")
            {
                decimal VatAmount = Convert.ToDecimal(TaxDataRow["VAT"], CultureInfo.InvariantCulture) - Vat_First;
                oUserTable.UserFields.Fields.Item("U_VAT").Value = Convert.ToDouble(VatAmount, CultureInfo.InvariantCulture);
            }

            oUserTable.UserFields.Fields.Item("U_KID").Value = TaxDataRow["K_ID"].ToString();

            try
            {
                DateTime AGREE_DATE = new DateTime();
                AGREE_DATE = DateTime.TryParse(TaxDataRow["AGREE_DATE"].ToString(), out AGREE_DATE) == false ? new DateTime() : AGREE_DATE; // დადასტურების თარიღი
                oUserTable.UserFields.Fields.Item("U_AGREEDATE").Value = AGREE_DATE;
            }
            catch { }

            oUserTable.UserFields.Fields.Item("U_AGREESUSID").Value = TaxDataRow["AGREE_S_USER_ID"].ToString();

            try
            {
                DateTime REF_DATE = new DateTime();
                REF_DATE = DateTime.TryParse(TaxDataRow["REF_DATE"].ToString(), out REF_DATE) == false ? new DateTime() : REF_DATE; // უარყოფის თარიღი
                oUserTable.UserFields.Fields.Item("U_REFDATE").Value = REF_DATE;
            }
            catch { }

            oUserTable.UserFields.Fields.Item("U_REFUSID").Value = TaxDataRow["REF_S_USER_ID"].ToString();

            int returnCode = oUserTable.Add();

            if (returnCode != 0)
            {
                int errCode;
                string errMsg;

                Program.oCompany.GetLastError(out errCode, out errMsg);
                errorText = "Error description : " + errMsg + "! Code : " + errCode;
            }
        }

        public static void fillValues(ref string insertquery, DataRow TaxDataRow, string overhead_no, decimal TANXA_First, decimal Vat_First)
        {
            string statusRS = TaxDataRow["STATUS"].ToString();
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
            Sbuilder.Append(@"INSERT INTO ""@BDOSTXANR""  (""Code"", 
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
            CommonFunctions.AppendXML(Sbuilder, BDO_TaxInvoiceReceived.getStatusValueByStatusNumber(statusRS));
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
            //insertquery = insertquery + Environment.NewLine + Sbuilder.ToString();            
        }

        public static void uiApp_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (oForm.Title.Contains("Tax Invoice Received Analysis") != true || oForm.Title.Contains("Down Payment Tax Invoice Received Analysis"))
                {
                    return;
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD && pVal.BeforeAction)
                {
                    fillUserCode(oForm);
                }

                else if (pVal.ItemUID == "1" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.BeforeAction && !pVal.InnerEvent)
                {
                    addRecord(oForm);
                }
            }
        }
    }
}
