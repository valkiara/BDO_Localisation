using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Collections;
using System.Data;

namespace BDO_Localisation_AddOn
{
    static partial class CommonFunctions
    {
        public static StringBuilder AppendXML(StringBuilder XML, string str)
        {
            str = str.Replace("&", "&amp;");
            str = str.Replace("\"", "&quot;");
            str = str.Replace("'", "&apos;");
            str = str.Replace("<", "&lt;");
            str = str.Replace(">", "&gt;");

            return XML.Append(str);
        }

        public static StringBuilder AddCellXML(StringBuilder XML, string columnUid, string value)
        {
            XML.Append("<Cell> <ColumnUid>");
            XML.Append(columnUid);
            XML.Append("</ColumnUid> <Value>");
            AppendXML(XML, value);
            XML.Append("</Value></Cell>");

            return XML;
        }

        public static void StartTransaction()
        {
            if (Program.oCompany.InTransaction != true)
            {
                Program.oCompany.StartTransaction();
                Program.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                Program.oCompany.StartTransaction();
            }
        }

        public static void EndTransaction(SAPbobsCOM.BoWfTransOpt EndType)
        {
            if (Program.oCompany.InTransaction)
            {
                Program.oCompany.EndTransaction(EndType);
            }
        }

        public static DataTable GetOACTTable()
        {

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            DataTable AccountTable = new DataTable();
            AccountTable.Columns.Add("AcctCode");
            AccountTable.Columns.Add("ActType");
            AccountTable.Columns.Add("U_BDOSEmpAct");

            try
            {
                string query = @"SELECT * FROM
	                         ""OACT""";

                oRecordSet.DoQuery(query);

                DataRow AccountTableRow = null;
                SAPbobsCOM.Fields oRecordSetFields = null;

                while (!oRecordSet.EoF)
                {
                    oRecordSetFields = oRecordSet.Fields;
                    AccountTableRow = AccountTable.Rows.Add();
                    AccountTableRow["AcctCode"] = oRecordSetFields.Item("AcctCode").Value;
                    AccountTableRow["ActType"] = oRecordSetFields.Item("ActType").Value;
                    AccountTableRow["U_BDOSEmpAct"] = oRecordSetFields.Item("U_BDOSEmpAct").Value;

                    oRecordSet.MoveNext();
                }
            }

            catch
            {
                return null;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                oRecordSet = null;
            }

            return AccountTable;


        }

        public static Dictionary<string, string> getCashFlowLineItemsList()
        {
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                Dictionary<string, string> CurList = new Dictionary<string, string>();

                string query = @"SELECT ""CFWId"", ""CFWName"" FROM ""OCFW"" WHERE ""Postable"" = 'Y'";

                oRecordSet.DoQuery(query);

                while (!oRecordSet.EoF)
                {
                    CurList.Add(oRecordSet.Fields.Item("CFWId").Value.ToString(), oRecordSet.Fields.Item("CFWName").Value.ToString());
                    oRecordSet.MoveNext();
                }

                return CurList;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
            }
        }

        public static bool IsDevelopment()
        {
            return UDO.UserDefinedFieldExists("OVPM", "BDOSBdgCf") && (getOADM("U_BDOSDevCmp").ToString().Trim() == "Y") && UDO.UserDefinedFieldExists("OINV", "BDOSLnOpTp");
        }

        public static string GetAccountDetermination(string Code, string Value)
        {
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string Query = @"select """ + Value + @""" from ""OADT""
                            where ""Code"" = '" + Code + "'";

            oRecordSet.DoQuery(Query);

            if (!oRecordSet.EoF)
            {
                return oRecordSet.Fields.Item(Value).Value.ToString();
            }


            return "";
        }

        public static decimal GetVatGroupRate(string VatGroup, string ItemCode)
        {
            NumberFormatInfo Nfi = new NumberFormatInfo() { NumberDecimalSeparator = "." };
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string query = "";
            if (VatGroup != "")
            {
                query = "SELECT " +
                            "* " +
                            "FROM \"OVTG\" " +
                            "WHERE \"OVTG\".\"Code\"='" + VatGroup + "'";
            }
            else
            {
                query = @"SELECT ""OVTG"".""Rate"" 
                        FROM ""OITM"" 
                        LEFT JOIN ""OVTG"" ON ""OITM"".""VatGourpSa"" = ""OVTG"".""Code"" 
                        WHERE ""OITM"".""ItemCode""=N'" + ItemCode.Replace("'", "''") + "'";
            }

            oRecordSet.DoQuery(query);

            if (!oRecordSet.EoF)
            {
                return Convert.ToDecimal(oRecordSet.Fields.Item("Rate").Value, CultureInfo.InvariantCulture);
            }
            else
            {
                return 0;
            }
        }

        public static string accountParse(string account, out string currency)
        {
            currency = null;

            if (string.IsNullOrEmpty(account) == false)
            {
                account = string.Join("", account.Split(default(string[]), StringSplitOptions.RemoveEmptyEntries));
                char[] lastChar = account.Substring(account.Length - 1).ToCharArray();
                if (char.IsDigit(lastChar[0]) == false)
                {
                    currency = account.Substring(account.Length - 3);
                    return account.Substring(0, account.Length - 3);
                }
            }
            return account;
        }

        public static string accountParse(string account)
        {
            if (string.IsNullOrEmpty(account) == false)
            {
                account = string.Join("", account.Split(default(string[]), StringSplitOptions.RemoveEmptyEntries));
                char[] lastChar = account.Substring(account.Length - 1).ToCharArray();
                if (char.IsDigit(lastChar[0]) == false)
                {
                    return account.Substring(0, account.Length - 3);
                }
            }
            return account;
        }

        public static object getOADM(string settingName)
        {
            string query = @"select """ + settingName + @""" from ""OADM""";

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            oRecordSet.DoQuery(query);

            if (!oRecordSet.EoF)
            {
                var value = oRecordSet.Fields.Item(settingName).Value;
                Marshal.FinalReleaseComObject(oRecordSet);
                oRecordSet = null;
                return value;
            }
            else
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                oRecordSet = null;
                return null;
            }
        }

        public static string getCurrencyInternationalCode(string currCode)
        {
            string currency = null;
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                string query = @"SELECT ""OCRN"".""DocCurrCod"" FROM ""OCRN"" WHERE ""OCRN"".""CurrCode"" = '" + currCode + "'";

                oRecordSet.DoQuery(query);

                while (!oRecordSet.EoF)
                {
                    currency = oRecordSet.Fields.Item("DocCurrCod").Value.ToString();

                    return currency;
                }
            }
            catch
            {
                return currency;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
            }
            return currency;
        }

        public static string getCurrencySapCode(string currInternationalCode)
        {
            string currency = null;
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                string query = @"SELECT ""OCRN"".""CurrCode"" FROM ""OCRN"" WHERE ""OCRN"".""DocCurrCod"" = '" + currInternationalCode + "'";

                oRecordSet.DoQuery(query);

                while (!oRecordSet.EoF)
                {
                    currency = oRecordSet.Fields.Item("CurrCode").Value.ToString();

                    return currency;
                }
            }
            catch
            {
                return currency;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
            }
            return currency;
        }

        public static string getPeriodsCategory(string column, string year)
        {
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                string query = @"SELECT """ + column + @""" FROM ""OACP"" WHERE ""Year"" = '" + year + "'";

                oRecordSet.DoQuery(query);

                if (!oRecordSet.EoF)
                {
                    return oRecordSet.Fields.Item("" + column + "").Value.ToString();
                }
            }
            catch
            {
                return null;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                oRecordSet = null;
            }
            return null;
        }

        public static SAPbobsCOM.Recordset getBPBankInfo(string account, string licTradNum, string cardType)
        {
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                if (string.IsNullOrEmpty(account) == false)
                {
                    string query = @"SELECT
                    	 ""OCRB"".""CardCode"",
                    	 ""OCRD"".""CardName"",
                    	 ""OCRD"".""CardType"",
                    	 ""OCRD"".""LicTradNum"",
                         ""OCRD"".""DebPayAcct"",
                         ""OCRD"".""ProjectCod"",
                         CAST(""OCRD"".""DflAgrmnt"" AS NVARCHAR) AS ""BlnkAgr"",
                         ""OCRD"".""Currency"",
                    	 ""OCRB"".""BankCode"",
                    	 ""OCRB"".""Country"",
                    	 ""OCRB"".""Account"",
                    	 ""OCRB"".""AcctName"",
                         ""OCRB"".""U_treasury""
                    FROM ""OCRB"" 
                    INNER JOIN ""OCRD"" ON ""OCRB"".""CardCode"" = ""OCRD"".""CardCode"" 
                    WHERE ""Account"" = '" + account + @"' AND ""OCRD"".""CardType"" = '" + cardType + "'";

                    if (!string.IsNullOrEmpty(licTradNum))
                    {
                        query = query + @" AND ""OCRD"".""LicTradNum"" = '" + licTradNum + "'";
                    }

                    oRecordSet.DoQuery(query);
                    if (!oRecordSet.EoF)
                    {
                        return oRecordSet;
                    }
                }
                return null;
            }
            catch
            {
                return null;
            }
        }

        public static SAPbobsCOM.Recordset getBPBankInfo(string account, string cardCode)
        {
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                if (!string.IsNullOrEmpty(account) && !string.IsNullOrEmpty(cardCode))
                {
                    string query = @"SELECT
                    	 ""OCRB"".""CardCode"",
                    	 ""OCRD"".""CardType"",
                         ""OCRD"".""Currency"",
                    	 ""OCRB"".""BankCode"",
                    	 ""OCRB"".""Country"",
                    	 ""OCRB"".""Account"",
                    	 ""OCRB"".""AcctName"",
                         ""OCRB"".""U_treasury""
                    FROM ""OCRB"" 
                    INNER JOIN ""OCRD"" ON ""OCRB"".""CardCode"" = ""OCRD"".""CardCode"" 
                    WHERE ""Account"" = '" + account + @"' AND ""OCRD"".""CardCode"" = '" + cardCode + "'";

                    oRecordSet.DoQuery(query);
                    if (!oRecordSet.EoF)
                    {
                        return oRecordSet;
                    }
                }
                return null;
            }
            catch
            {
                return null;
            }
        }

        public static SAPbobsCOM.Recordset getBPBankInfo(string cardCode)
        {
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                if (!string.IsNullOrEmpty(cardCode))
                {
                    StringBuilder query = new StringBuilder();
                    query.Append("SELECT \"CardCode\", \n");
                    query.Append("\"CardName\", \n");
                    query.Append("\"DebPayAcct\", \n");
                    query.Append("\"Currency\", \n");
                    query.Append("\"BankCountr\", \n");
                    query.Append("\"BankCode\", \n");
                    query.Append("\"DflAccount\" \n");
                    query.Append("FROM   \"OCRD\" \n");
                    query.Append("WHERE  \"CardCode\" = '" + cardCode + "'");

                    oRecordSet.DoQuery(query.ToString());
                    if (!oRecordSet.EoF)
                    {
                        return oRecordSet;
                    }
                }
                return null;
            }
            catch
            {
                return null;
            }
        }

        public static bool isBPAccountTreasury(string cardCode, string bankCode, string account)
        {
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                if (!string.IsNullOrEmpty(cardCode) && !string.IsNullOrEmpty(bankCode) && !string.IsNullOrEmpty(account))
                {
                    string query = @"SELECT ""Account""
                    FROM ""OCRB""
                    WHERE ""CardCode"" = '" + cardCode + @"' AND ""BankCode"" = '" + bankCode + @"' AND ""Account"" = '" + account + @"' AND ""U_treasury"" = 'Y'";

                    oRecordSet.DoQuery(query);
                    if (!oRecordSet.EoF)
                    {
                        return true;
                    }
                }
                return false;
            }
            catch
            {
                return false;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                oRecordSet = null;
            }
        }

        public static SAPbobsCOM.Recordset getEmployeeInfo(string govID)
        {
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                if (string.IsNullOrEmpty(govID) == false)
                {
                    string query = @"SELECT ""OHEM"".""empID""
                         FROM ""OHEM""
                         WHERE ""OHEM"".""govID"" = '" + govID + @"'";

                    oRecordSet.DoQuery(query);
                    if (!oRecordSet.EoF)
                    {
                        return oRecordSet;
                    }
                }
                Marshal.FinalReleaseComObject(oRecordSet);
                oRecordSet = null;
                return null;
            }
            catch
            {
                return null;
            }
        }

        public static Dictionary<string, string> getCurrencyListForValidValues()
        {
            string query = @"SELECT 
            ""CurrCode"", 
            ""CurrName"" 
            FROM ""OCRN""";

            Dictionary<string, string> listValidValuesDict = new Dictionary<string, string>();

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oRecordSet.DoQuery(query);

                listValidValuesDict.Add("", "");

                while (!oRecordSet.EoF)
                {
                    listValidValuesDict.Add(oRecordSet.Fields.Item("CurrCode").Value.ToString(), oRecordSet.Fields.Item("CurrName").Value.ToString());

                    oRecordSet.MoveNext();
                }
                return listValidValuesDict;
            }
            catch
            {
                return listValidValuesDict;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
            }
        }

        public static Dictionary<string, string> getCashFlowItemListForValidValues()
        {
            string query = @"SELECT 
            ""CFWId"", 
            ""CFWName"" 
            FROM ""OCFW""";

            Dictionary<string, string> listValidValuesDict = new Dictionary<string, string>();

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oRecordSet.DoQuery(query);

                listValidValuesDict.Add("", "");

                while (!oRecordSet.EoF)
                {
                    listValidValuesDict.Add(oRecordSet.Fields.Item("CFWId").Value.ToString(), oRecordSet.Fields.Item("CFWName").Value.ToString());

                    oRecordSet.MoveNext();
                }
                return listValidValuesDict;
            }
            catch
            {
                return listValidValuesDict;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
            }
        }

        public static string getTransferAccount(string account)
        {
            string transferAccount = null;

            //BOG
            if (account.IndexOf("RUB") >= 0 && account.IndexOf("GE51BG") >= 0)
            {
                account = account.Replace("RUB", "RUR");

            }


            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                string query = @"SELECT ""DSC1"".""GLAccount"" FROM ""DSC1"" WHERE ""DSC1"".""Account"" = '" + account + "'";
                oRecordSet.DoQuery(query);
                if (!oRecordSet.EoF)
                {
                    transferAccount = oRecordSet.Fields.Item("GLAccount").Value.ToString();
                    return transferAccount;
                }
                return transferAccount;
            }
            catch
            {
                return transferAccount;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                oRecordSet = null;
            }
        }

        public static bool isAccountInHouseBankAccount(string account)
        {
            //BOG
            if (account.IndexOf("RUB") >= 0 && account.IndexOf("GE51BG") >= 0)
            {
                account = account.Replace("RUB", "RUR");

            }

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                string query = @"SELECT ""DSC1"".""GLAccount"" FROM ""DSC1"" WHERE ""DSC1"".""Account"" = '" + account + "'";
                oRecordSet.DoQuery(query);
                if (!oRecordSet.EoF)
                {
                    return true;
                }
                return false;
            }
            catch
            {
                return false;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                oRecordSet = null;
            }
        }

        public static bool isAccountCashFlowRelevant(string GLAccount)
        {
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                string query = @"SELECT ""OACT"".""CfwRlvnt"" FROM ""OACT"" WHERE ""OACT"".""AcctCode"" = '" + GLAccount + @"' AND ""OACT"".""CfwRlvnt"" = 'Y'";
                oRecordSet.DoQuery(query);
                if (!oRecordSet.EoF)
                {
                    return true;
                }
                return false;
            }
            catch
            {
                return false;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                oRecordSet = null;
            }
        }

        public static string getHouseBankAccount(string bankCode, string currency)
        {
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                string query = @"SELECT ""DSC1"".""Account"" FROM ""DSC1"" 
                   WHERE ""DSC1"".""BankCode"" = '" + bankCode + "'" +
                   @"AND LOCATE(""DSC1"".""Account"",
                   '" + currency + "') > 0";

                oRecordSet.DoQuery(query);
                if (!oRecordSet.EoF)
                {
                    return oRecordSet.Fields.Item("Account").Value.ToString();
                }
                return null;
            }
            catch
            {
                return null;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                oRecordSet = null;
            }
        }

        public static string getBankProgram(string trsfrAcct = null, string account = null)
        {
            string program = null;

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                string query = @"SELECT ""DSC1"".""U_program"" FROM ""DSC1"" WHERE ";

                if (string.IsNullOrEmpty(trsfrAcct) == false)
                    query = query + @" ""DSC1"".""GLAccount"" = '" + trsfrAcct + "'";
                else
                { query = query + @" ""DSC1"".""Account"" = '" + account + "'"; }

                oRecordSet.DoQuery(query);
                if (!oRecordSet.EoF)
                {
                    program = oRecordSet.Fields.Item("U_program").Value.ToString();
                    return program;
                }
                return program;
            }
            catch
            {
                return program;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                oRecordSet = null;
            }
        }

        public static string getBankCode(string trsfrAcct = null, string account = null)
        {
            string bankCode = null;

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                string query = @"SELECT ""DSC1"".""BankCode"" FROM ""DSC1"" WHERE ";

                if (string.IsNullOrEmpty(trsfrAcct) == false)
                    query = query + @" ""DSC1"".""GLAccount"" = '" + trsfrAcct + "'";
                else
                { query = query + @" ""DSC1"".""Account"" = '" + account + "'"; }

                oRecordSet.DoQuery(query);
                if (!oRecordSet.EoF)
                {
                    bankCode = oRecordSet.Fields.Item("BankCode").Value.ToString();
                    return bankCode;
                }
                return bankCode;
            }
            catch
            {
                return bankCode;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                oRecordSet = null;
            }
        }

        public static string getBankName(string bankCode)
        {
            string bankName = null;

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                string query = @"SELECT ""ODSC"".""BankName"" FROM ""ODSC"" WHERE ""ODSC"".""BankCode"" = '" + bankCode + "'";

                oRecordSet.DoQuery(query);
                if (!oRecordSet.EoF)
                {
                    bankName = oRecordSet.Fields.Item("BankName").Value.ToString();
                    return bankName;
                }
                return bankName;
            }
            catch
            {
                return bankName;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                oRecordSet = null;
            }
        }

        public static string getAccountCurrency(string accountCode)
        {
            string accountCurrency = null;

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                string query = @"SELECT ""OACT"".""ActCurr"" FROM ""OACT"" WHERE ""OACT"".""AcctCode"" = '" + accountCode + "'";

                oRecordSet.DoQuery(query);
                if (!oRecordSet.EoF)
                {
                    accountCurrency = oRecordSet.Fields.Item("ActCurr").Value.ToString();
                    return accountCurrency;
                }
                return accountCurrency;
            }
            catch
            {
                return accountCurrency;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                oRecordSet = null;
            }
        }

        public static string getServiceUrlForInternetBanking(string program, out string clientID, out int port, out string errorText)
        {
            errorText = null;
            clientID = null;
            port = 0;

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string mode;
            string wsdl;
            string url;

            try
            {
                string query = @"SELECT 
                ""U_program"" AS ""program"",
                ""U_mode"" AS ""mode"",
                ""U_WSDL"" AS ""WSDL"",
                ""U_ID"" AS ""ID"",
                ""U_URL"" AS ""URL"",
                ""U_port"" AS ""port""
                FROM ""@BDO_INTB""
                WHERE ""U_program"" = '" + program + "'";

                oRecordSet.DoQuery(query);
                if (!oRecordSet.EoF)
                {
                    mode = oRecordSet.Fields.Item("mode").Value;
                    wsdl = oRecordSet.Fields.Item("WSDL").Value;
                    clientID = oRecordSet.Fields.Item("ID").Value;
                    url = oRecordSet.Fields.Item("URL").Value;
                    port = oRecordSet.Fields.Item("port").Value;

                    if (string.IsNullOrEmpty(mode))
                    {
                        errorText = "ინტერნეტბანკის გაცვლის რეჟიმი არ არის შევსებული" + "! (" + program + ")";
                        return null;
                    }
                    if ((program == "TBC" && string.IsNullOrEmpty(wsdl)) || (program == "BOG" && string.IsNullOrEmpty(url)))
                    {
                        errorText = "ინტერნეტბანკის გაცვლის მისამართი არ არის შევსებული" + "! (" + program + ")";
                        return null;
                    }
                    if (program == "BOG" && mode == "test" && port == 0)
                    {
                        errorText = "ინტერნეტბანკის გაცვლისთვის პორტი არ არის შევსებული" + "! (" + program + ")";
                        return null;
                    }
                    if (program == "BOG")
                    {
                        return url;
                    }

                    return wsdl;
                }
                return null;
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
                return null;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
            }
        }

        public static Dictionary<string, string> getCompanyInfo()
        {
            Dictionary<string, string> CompanyInfo = new Dictionary<string, string>();
            CompanyInfo.Add("CompnyName", "");
            CompanyInfo.Add("CompnyAddr", "");
            CompanyInfo.Add("FreeZoneNo", "");
            CompanyInfo.Add("DflBnkCode", "");

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                string query = @"SELECT ""OADM"".""CompnyName"", ""OADM"".""CompnyAddr"", ""OADM"".""FreeZoneNo"", ""OADM"".""DflBnkCode"" FROM ""OADM""";

                oRecordSet.DoQuery(query);
                if (!oRecordSet.EoF)
                {
                    CompanyInfo["CompnyName"] = oRecordSet.Fields.Item("CompnyName").Value.ToString();
                    CompanyInfo["CompnyAddr"] = oRecordSet.Fields.Item("CompnyAddr").Value.ToString();
                    CompanyInfo["FreeZoneNo"] = oRecordSet.Fields.Item("FreeZoneNo").Value.ToString();
                    CompanyInfo["DflBnkCode"] = oRecordSet.Fields.Item("DflBnkCode").Value.ToString();

                    //return CompanyInfo;
                }
                //return CompanyInfo;
            }
            catch
            {
                //return CompanyInfo;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                oRecordSet = null;
            }

            return CompanyInfo;
        }

        public static Dictionary<string, string> getCompanyLicenseInfo()
        {
            Dictionary<string, string> CompanyLicenseInfo = new Dictionary<string, string>();
            CompanyLicenseInfo.Add("LicenseKey", "");
            CompanyLicenseInfo.Add("LicenseStatus", "");
            CompanyLicenseInfo.Add("LicenseUpdateDate", "");
            CompanyLicenseInfo.Add("LicenseQuantity", "");

            License oLicense = new License();
            string result = getOADM("U_BDOSLocLic").ToString();
            if (result != "")
            {
                string deCryptText = oLicense.CryptText(result, 0);

                CompanyLicenseInfo["LicenseKey"] = oLicense.GetValueTeg("СерииныйНомер", deCryptText);

                string LicenseUpdateDt = oLicense.GetValueTeg("ДатаПоследнегоЗапроса", deCryptText);

                SAPbobsCOM.SBObob oSBOBob = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);
                SAPbobsCOM.IRecordset oIRecordset = oSBOBob.Format_StringToDate(LicenseUpdateDt);
                DateTime LiceUpdateDate = oIRecordset.Fields.Item("Date").Value;
                LiceUpdateDate = LiceUpdateDate.AddSeconds(10 * 24 * 3600);
                if (oLicense.GetValueTeg("ЛицензияАктивна", deCryptText) != BDOSResources.getTranslate("Active") || DateTime.Today > LiceUpdateDate)
                {
                    CompanyLicenseInfo["LicenseStatus"] = BDOSResources.getTranslate("Inactive");
                }
                else
                {
                    CompanyLicenseInfo["LicenseStatus"] = BDOSResources.getTranslate("Active");
                }

                CompanyLicenseInfo["LicenseUpdateDate"] = LicenseUpdateDt;

                string licenseQuantity = oLicense.GetValueTeg("КоличествоЛицензии", deCryptText);
                CompanyLicenseInfo["LicenseQuantity"] = (licenseQuantity == "" ? "0" : licenseQuantity);

                Marshal.ReleaseComObject(oSBOBob);
                oSBOBob = null;
                Marshal.ReleaseComObject(oIRecordset);
                oIRecordset = null;
            }

            return CompanyLicenseInfo;

        }

        public static decimal roundAmountByGeneralSettings(decimal amount, string DecType, RoundingDirection roundingDir = RoundingDirection.Other)
        {
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                string query = @"SELECT * FROM ""OADM""";
                int sumDec = 2;
                oRecordSet.DoQuery(query);
                if (!oRecordSet.EoF)
                {
                    sumDec = Convert.ToInt32(oRecordSet.Fields.Item(DecType + "Dec").Value);
                }
                if (roundingDir == RoundingDirection.Other)
                {
                    return Math.Round(amount, sumDec);
                }
                return Round(Convert.ToDouble(amount, CultureInfo.InvariantCulture), sumDec, roundingDir);
            }
            catch
            {
                if (roundingDir == RoundingDirection.Other)
                {
                    return Math.Round(amount, 2);
                }
                return Round(Convert.ToDouble(amount, CultureInfo.InvariantCulture), 2, roundingDir);
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                oRecordSet = null;
            }
        }

        private delegate double RoundingFunction(double value);

        public enum RoundingDirection { Up, Down, Other }

        private static decimal Round(double value, int precision, RoundingDirection roundingDirection)
        {
            RoundingFunction roundingFunction;
            if (roundingDirection == RoundingDirection.Up)
                roundingFunction = Math.Ceiling;
            else
                roundingFunction = Math.Floor;
            value *= Math.Pow(10, precision);
            value = roundingFunction(value);
            return Convert.ToDecimal(value * Math.Pow(10, -1 * precision), CultureInfo.InvariantCulture);
        }

        public static string getRegistrationCountryCode(string account, string table)
        {
            string registrationCountryCode = null;

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                string query = @"SELECT """ + table + @""".""Country"" FROM """ + table + @""" WHERE """ + table + @""".""Account"" = '" + account + "'";

                oRecordSet.DoQuery(query);
                if (!oRecordSet.EoF)
                {
                    registrationCountryCode = oRecordSet.Fields.Item("Country").Value.ToString();
                    return registrationCountryCode;
                }
                return registrationCountryCode;
            }
            catch
            {
                return registrationCountryCode;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                oRecordSet = null;
            }
        }

        public static bool codeIsUsed(Dictionary<string, string> listTables, string code)
        {
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string query = "";

            foreach (KeyValuePair<string, string> keyValue in listTables)
            {
                string tableName = keyValue.Key;
                string fieldName = keyValue.Value;
                query = query + @" SELECT """ + fieldName + @""" FROM """ + tableName + @""" WHERE """ + fieldName.Replace("'", "''") + @""" = '" + code + "' ";
                query = query + @" UNION ALL ";
            }

            query = query.Remove(query.Length - 11);

            oRecordSet.DoQuery(query);
            if (!oRecordSet.EoF)
            {
                return true;
            }

            return false;
        }

        public static string getLocalCurrency()
        {
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                string query = @"SELECT ""MainCurncy"" FROM ""OADM""";
                oRecordSet.DoQuery(query);
                if (!oRecordSet.EoF)
                {
                    return oRecordSet.Fields.Item("MainCurncy").Value;
                }
                return null;
            }
            catch
            {
                return null;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                oRecordSet = null;
            }
        }

        public static Dictionary<string, string> getActiveDimensionsList(out string errorText)
        {
            errorText = null;

            Dictionary<string, string> activeDimensionsList = new Dictionary<string, string>();
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string query = @"SELECT * FROM ""ODIM"" WHERE ""DimActive""='Y'";

            oRecordSet.DoQuery(query);

            while (!oRecordSet.EoF)
            {
                activeDimensionsList.Add(oRecordSet.Fields.Item("DimCode").Value.ToString(), oRecordSet.Fields.Item("DimDesc").Value.ToString());
                oRecordSet.MoveNext();
            }

            return activeDimensionsList;
        }

        public static void fillDocRate(SAPbouiCOM.Form oForm, string documentTable, string paymentTable)
        {
            decimal DocRate = 0;
            string errorText = null;

            SAPbouiCOM.DBDataSource DBDataSource11 = oForm.DataSources.DBDataSources.Item(paymentTable);
            SAPbouiCOM.DBDataSource DBDataSourceO = oForm.DataSources.DBDataSources.Item(documentTable);
            string DocDateStr = DBDataSourceO.GetValue("DocDate", 0);
            DateTime DocDate = DateTime.TryParseExact(DocDateStr, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DocDate) == false ? DateTime.Now : DocDate;

            string DocCurr = "";

            if (documentTable == "OINV" || documentTable == "OPCH" || documentTable == "ODPI" || documentTable == "ODPO")
            {
                DocCurr = DBDataSourceO.GetValue("DocCur", 0);
            }
            else DocCurr = DBDataSourceO.GetValue("DocCurr", 0);
            if (DocCurr == getLocalCurrency())
            {
                return;
            }

            if (DBDataSourceO.GetValue("CANCELED", 0) != "N")
            {
                return;
            }


            string UseBlaAgRt = DBDataSourceO.GetValue("U_UseBlaAgRt", 0);
            string BlaAgrDocEntryStr = DBDataSourceO.GetValue("AgrNo", 0);
            if (BlaAgrDocEntryStr != "")
            {
                int BlaAgrDocEntry = Convert.ToInt32(BlaAgrDocEntryStr);
                string docCurr;
                if (UseBlaAgRt == "Y")
                {
                    DocRate = BlanketAgreement.GetBlAgremeentCurrencyRate(BlaAgrDocEntry, out docCurr, DocDate);

                }
            }
            if (documentTable == "OINV" || documentTable == "OPCH")
            {
                if (DBDataSource11.Size > 0)
                {
                    decimal GrossSum = 0;
                    decimal GrossFCSum = 0;
                    for (int row = 0; row < DBDataSource11.Size; row++)
                    {
                        decimal Gross = FormsB1.cleanStringOfNonDigits(DBDataSource11.GetValue("BaseGross", row));
                        decimal GrossFC = FormsB1.cleanStringOfNonDigits(DBDataSource11.GetValue("BaseGrossF", row));


                        if (Gross > 0)
                        {
                            GrossSum = GrossSum + Gross;
                            GrossFCSum = GrossFCSum + GrossFC;

                        }
                    }

                    decimal DocTotalFC = FormsB1.cleanStringOfNonDigits(DBDataSourceO.GetValue("DocTotalFC", 0));
                    decimal DocTotal = FormsB1.cleanStringOfNonDigits(DBDataSourceO.GetValue("DocTotal", 0));




                    if (UseBlaAgRt == "N")
                    {
                        SAPbobsCOM.SBObob oSBOBob = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);
                        SAPbobsCOM.Recordset RateRecordset = oSBOBob.GetCurrencyRate(DocCurr, DocDate);

                        while (!RateRecordset.EoF)
                        {
                            DocRate = Convert.ToDecimal(RateRecordset.Fields.Item("CurrencyRate").Value);
                            RateRecordset.MoveNext();
                        }
                    }



                    DocRate = (roundAmountByGeneralSettings(DocTotalFC * DocRate, "Sum") + GrossSum) / (DocTotalFC + GrossFCSum);

                }
            }

            if (DocRate > 0)
            {

                if (documentTable == "OINV" || documentTable == "OPCH" || documentTable == "ODPI" || documentTable == "ODPO")
                {
                    if (oForm.Items.Item("64").Enabled)
                    {
                        oForm.Items.Item("64").Specific.Value = FormsB1.ConvertDecimalToString(DocRate);

                    }
                }

                else if ((documentTable == "OVPM" || documentTable == "ORCT") && (oForm.TypeEx == "196" || oForm.TypeEx == "146"))
                {
                    if (oForm.Items.Item("95").Enabled && oForm.Items.Item("95").Visible)
                    {
                        oForm.Items.Item("95").Specific.Value = FormsB1.ConvertDecimalToString(DocRate);
                    }
                }

                else
                {
                    if (oForm.Items.Item("21").Enabled && oForm.Items.Item("21").Visible)
                    {
                        oForm.Items.Item("21").Specific.Value = FormsB1.ConvertDecimalToString(DocRate);
                    }
                }
            }

        }

        public static void fillPhysicalEntityTaxes(string objType, SAPbouiCOM.Form oFormWtax, SAPbouiCOM.Form oForm, string docDBSourcesName, string tableDBSourcesName, out string errorText)
        {
            errorText = "";

            oForm.Freeze(true);

            try
            {
                SAPbouiCOM.DBDataSources docDBSources = oForm.DataSources.DBDataSources;
                string wtCode = docDBSources.Item("OCRD").GetValue("WTCode", 0).Trim();

                bool physicalEntityTax = (docDBSources.Item("OCRD").GetValue("WTLiable", 0).Trim() == "Y" &&
                                            getValue("OWHT", "U_BDOSPhisTx", "WTCode", wtCode).ToString() == "Y");

                Dictionary<string, decimal> PhysicalEntityPensionRates;

                string errorTextCheck;
                string docDatestr = docDBSources.Item(docDBSourcesName).GetValue("DocDate", 0).Trim();
                bool frgn = docDBSources.Item(docDBSourcesName).GetValue("DocCur", 0).Trim() != getLocalCurrency();
                if (physicalEntityTax)
                {
                    if (string.IsNullOrEmpty(docDatestr))
                    {
                        errorText = BDOSResources.getTranslate("DocDate") + " " + BDOSResources.getTranslate("YouCantLeaveEmpty");
                        return;
                    }

                    DateTime DocDate = DateTime.ParseExact(docDatestr, "yyyyMMdd", CultureInfo.InvariantCulture);
                    PhysicalEntityPensionRates = WithholdingTax.getPhysicalEntityPensionRates(DocDate, wtCode, out errorTextCheck);

                    if (!string.IsNullOrEmpty(errorTextCheck))
                    {
                        errorText = errorTextCheck;
                        return;
                    }
                }
                else
                {
                    PhysicalEntityPensionRates = new Dictionary<string, decimal>();
                    PhysicalEntityPensionRates.Add("WTRate", 0);
                    PhysicalEntityPensionRates.Add("PensionWTaxRate", 0);
                    PhysicalEntityPensionRates.Add("PensionCoWTaxRate", 0);
                }

                string docType = docDBSources.Item(docDBSourcesName).GetValue("DocType", 0).Trim();
                SAPbouiCOM.Matrix oMatrix;

                if (docType == "I")
                {
                    oMatrix = oForm.Items.Item("38").Specific;
                }
                else
                {
                    oMatrix = oForm.Items.Item("39").Specific;
                }

                SAPbouiCOM.Matrix oMatrixWtax = oFormWtax.Items.Item("6").Specific;
                SAPbouiCOM.DBDataSource DBDataSourceTable = docDBSources.Item(tableDBSourcesName);

                decimal totalTaxes = 0;
                decimal PensPhAm;
                decimal WhtAmt;
                decimal PensCoAm;
                decimal GrossAmount;
                decimal PensPhAmFC;
                decimal WhtAmtFC;
                decimal GrossAmountFC;

                for (int row = 0; row < DBDataSourceTable.Size; row++)
                {
                    PensPhAm = 0;
                    WhtAmt = 0;
                    PensCoAm = 0;
                    GrossAmount = 0;
                    GrossAmountFC = 0;

                    if (physicalEntityTax && DBDataSourceTable.GetValue("WtLiable", row).Trim() == "Y")
                    {
                        GrossAmount = Convert.ToDecimal(getChildOrDbDataSourceValue(DBDataSourceTable, null, null, "LineTotal", row), CultureInfo.InvariantCulture);

                        PensPhAm = roundAmountByGeneralSettings(GrossAmount * PhysicalEntityPensionRates["PensionWTaxRate"] / 100, "Sum");
                        WhtAmt = roundAmountByGeneralSettings((GrossAmount - PensPhAm) * PhysicalEntityPensionRates["WTRate"] / 100, "Sum");
                        PensCoAm = roundAmountByGeneralSettings(GrossAmount * PhysicalEntityPensionRates["PensionCoWTaxRate"] / 100, "Sum");

                        if (frgn)
                        {
                            GrossAmountFC = Convert.ToDecimal(getChildOrDbDataSourceValue(DBDataSourceTable, null, null, "TotalFrgn", row), CultureInfo.InvariantCulture);
                            PensPhAmFC = roundAmountByGeneralSettings(GrossAmountFC * PhysicalEntityPensionRates["PensionWTaxRate"] / 100, "Sum");
                            WhtAmtFC = roundAmountByGeneralSettings((GrossAmountFC - PensPhAmFC) * PhysicalEntityPensionRates["WTRate"] / 100, "Sum");
                            totalTaxes = totalTaxes + PensPhAmFC + WhtAmtFC;
                        }
                        else
                        {
                            totalTaxes = totalTaxes + PensPhAm + WhtAmt;
                        }
                    }

                    int rowNumber = row + 1;//Convert.ToInt32(DBDataSourceTable.GetValue("LineNum", row));

                    oMatrix.Columns.Item("U_BDOSWhtAmt").Cells.Item(rowNumber).Specific.String = FormsB1.ConvertDecimalToStringForEditboxStrings(WhtAmt);
                    oMatrix.Columns.Item("U_BDOSPnPhAm").Cells.Item(rowNumber).Specific.String = FormsB1.ConvertDecimalToStringForEditboxStrings(PensPhAm);
                    oMatrix.Columns.Item("U_BDOSPnCoAm").Cells.Item(rowNumber).Specific.String = FormsB1.ConvertDecimalToStringForEditboxStrings(PensCoAm);
                }

                if (objType != "204") //A/P Reserve Invoice, A/P Invoice, A/P Credit Memo
                {
                    decimal taxableAmt = FormsB1.cleanStringOfNonDigits(oMatrixWtax.Columns.Item("7").Cells.Item(1).Specific.Value);
                    PensPhAm = roundAmountByGeneralSettings(taxableAmt * PhysicalEntityPensionRates["PensionWTaxRate"] / 100, "Sum");
                    WhtAmt = roundAmountByGeneralSettings((taxableAmt - PensPhAm) * PhysicalEntityPensionRates["WTRate"] / 100, "Sum");
                    totalTaxes = PensPhAm + WhtAmt;
                }

                if (physicalEntityTax)
                {
                    decimal oldWhtAmt = Convert.ToDecimal(getChildOrDbDataSourceValue(docDBSources.Item(docDBSourcesName), null, null, "WTSum", 0), CultureInfo.InvariantCulture);
                    if (oldWhtAmt != totalTaxes)
                    {
                        if (frgn)
                        {
                            oMatrixWtax.Columns.Item("28").Cells.Item(1).Specific.String = FormsB1.ConvertDecimalToStringForEditboxStrings(totalTaxes);
                        }
                        else
                        {
                            oMatrixWtax.Columns.Item("14").Cells.Item(1).Specific.String = FormsB1.ConvertDecimalToStringForEditboxStrings(totalTaxes);
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
                oForm.Freeze(false);
            }
        }

        public static void blockAssetInvoice(SAPbouiCOM.Form oForm, string docDBSourcesName, string tableDBSourcesName, string whsFieldName, out bool rejection)
        {
            rejection = false;
            SAPbouiCOM.DBDataSource DocDBSource = oForm.DataSources.DBDataSources.Item(docDBSourcesName);
            SAPbouiCOM.DBDataSource DocDBSourceTable = oForm.DataSources.DBDataSources.Item(tableDBSourcesName);


            DateTime DocDate = DateTime.ParseExact(DocDBSource.GetValue("DocDate", 0), "yyyyMMdd", CultureInfo.InvariantCulture);
            DocDate = new DateTime(DocDate.Year, DocDate.Month, 1);
            DocDate = DocDate.AddMonths(1).AddDays(-1);

            string ItemCodes = "";

            for (int i = 0; i < DocDBSourceTable.Size; i++)
            {
                ItemCodes = ItemCodes + "'" + DocDBSourceTable.GetValue("ItemCode", i).ToString() + "'";
                ItemCodes = ItemCodes + (i == DocDBSourceTable.Size - 1 ? "" : ",");
            }


            string query = BDOSDepreciationAccrualWizard.BatchDepreciaionQuery(DocDate, ItemCodes, "", "", false);
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);


            oRecordSet.DoQuery(query);
            while (!oRecordSet.EoF)
            {
                string ItemCode = oRecordSet.Fields.Item("ItemCode").Value;
                string DistNumber = oRecordSet.Fields.Item("DistNumber").Value;

                decimal futureDeprAmt = Convert.ToDecimal(oRecordSet.Fields.Item("FutureDeprAmt").Value, CultureInfo.InvariantCulture);
                decimal CurrDeprAmt = Convert.ToDecimal(oRecordSet.Fields.Item("CurrDeprAmt").Value, CultureInfo.InvariantCulture);
                if (CurrDeprAmt > 0 || futureDeprAmt > 0)
                {
                    rejection = true;
                    Program.uiApp.SetStatusBarMessage(BDOSResources.getTranslate("ThereIsDepreciationAmountsInCurrentMonthForItem") + " " + ItemCode + ": " + DistNumber);
                }

                oRecordSet.MoveNext();
            }
        }

        public static void blockNegativeStockByDocDate(SAPbouiCOM.Form oForm, string docDBSourcesName, string tableDBSourcesName, string whsFieldName, out bool rejection)
        {
            rejection = false;

            oForm.Freeze(true);

            try
            {
                SAPbouiCOM.DBDataSources docDBSources = oForm.DataSources.DBDataSources;
                string docType = docDBSources.Item(docDBSourcesName).GetValue("DocType", 0).Trim();

                if (docType == "I")
                {
                    string docDatestr = docDBSources.Item(docDBSourcesName).GetValue("DocDate", 0).Trim();

                    string blockStock = getOADM("U_BDOSBlcPDt").ToString().Trim();
                    bool blockStockByCompany = (blockStock == "ByCompany");
                    bool blockStockByWarehouse = (blockStock == "ByWarehouse");

                    if (blockStockByCompany || blockStockByWarehouse)
                    {
                        DataTable docLines = new DataTable();
                        docLines.Columns.Add("ItemCode", typeof(string));
                        docLines.Columns.Add("WhsCode", typeof(string));
                        docLines.Columns.Add("Quantity", typeof(decimal));
                        DataRow docLinesRow = null;

                        DataTable stockLines = new DataTable();
                        stockLines.Columns.Add("ItemCode", typeof(string));
                        stockLines.Columns.Add("WhsCode", typeof(string));
                        stockLines.Columns.Add("FinalQty", typeof(decimal));
                        DataRow stockLinesRow = null;

                        decimal quantity;
                        string itemCode;
                        string whsCode;

                        StringBuilder queryBuilder = new StringBuilder();
                        queryBuilder.Append(@"SELECT ");
                        if (blockStockByWarehouse)
                        {
                            queryBuilder.Append(@" ""LocCode"" AS ""WhsCode"", ");
                        }
                        queryBuilder.Append(@"""ItemCode"",
                                              SUM(""InQty"" - ""OutQty"") AS ""FinalQty""
                                             FROM ""OIVL"" 
                                            WHERE ""DocDate"" <='");
                        queryBuilder.Append(docDatestr);
                        queryBuilder.Append("' ");

                        int rowQty = 0;
                        SAPbouiCOM.DBDataSource DBDataSourceTable = docDBSources.Item(tableDBSourcesName);
                        for (int row = 0; row < DBDataSourceTable.Size; row++)
                        {
                            itemCode = getChildOrDbDataSourceValue(DBDataSourceTable, null, null, "ItemCode", row).ToString().Trim();
                            if (getValue("OITM", "InvntItem", "ItemCode", itemCode).ToString() == "Y")
                            {
                                quantity = Convert.ToDecimal(getChildOrDbDataSourceValue(DBDataSourceTable, null, null, "Quantity", row), CultureInfo.InvariantCulture);
                                whsCode = getChildOrDbDataSourceValue(DBDataSourceTable, null, null, whsFieldName, row).ToString().Trim();

                                if (row == 0)
                                {
                                    queryBuilder.Append(" AND ( ");
                                }
                                else
                                {
                                    queryBuilder.Append(" OR ");
                                }

                                queryBuilder.Append(@"""ItemCode"" = N'");
                                queryBuilder.Append(itemCode);
                                queryBuilder.Append("'");

                                docLinesRow = docLines.Rows.Add();
                                docLinesRow["ItemCode"] = itemCode;
                                docLinesRow["Quantity"] = quantity;

                                if (blockStockByWarehouse)
                                {
                                    queryBuilder.Append(@" AND ""LocCode"" = N'");
                                    queryBuilder.Append(whsCode);

                                    docLinesRow["WhsCode"] = whsCode;
                                    queryBuilder.Append("'");
                                }

                                rowQty++;
                            }
                        }

                        if (rowQty == 0) return;

                        queryBuilder.Append(" ) ");
                        queryBuilder.Append(" GROUP BY ");
                        if (blockStockByWarehouse)
                        {
                            queryBuilder.Append(@" ""LocCode"", ");
                        }
                        queryBuilder.Append(@" ""ItemCode"" ");

                        //ნაშთები
                        SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        string query = queryBuilder.ToString();

                        oRecordSet.DoQuery(query);
                        while (!oRecordSet.EoF)
                        {
                            stockLinesRow = stockLines.Rows.Add();
                            stockLinesRow["ItemCode"] = oRecordSet.Fields.Item("ItemCode").Value;
                            stockLinesRow["FinalQty"] = Convert.ToDecimal(oRecordSet.Fields.Item("FinalQty").Value, CultureInfo.InvariantCulture);
                            if (blockStockByWarehouse)
                            {
                                stockLinesRow["WhsCode"] = oRecordSet.Fields.Item("WhsCode").Value;
                            }
                            oRecordSet.MoveNext();
                        }
                        Marshal.ReleaseComObject(oRecordSet);
                        oRecordSet = null;

                        //დოკუმენტის სტრიქონების დაჯამება
                        if (docLines.Rows.Count > 0)
                        {
                            docLines = docLines.AsEnumerable().GroupBy(row => new
                            {
                                ItemCode = row.Field<string>("ItemCode"),
                                WhsCode = row.Field<string>("WhsCode")
                            })
                                          .Select(g =>
                                          {
                                              var row = docLines.NewRow();
                                              row["ItemCode"] = g.Key.ItemCode;
                                              row["WhsCode"] = g.Key.WhsCode;
                                              row["Quantity"] = g.Sum(r => r.Field<decimal>("Quantity"));
                                              return row;
                                          }).CopyToDataTable();
                        }

                        string errorText;
                        string filtString;
                        DataRow[] foundRows;
                        decimal finalQty;
                        decimal docQty;
                        bool tempCheck;

                        for (int i = 0; i < docLines.Rows.Count; i++)
                        {
                            tempCheck = true;
                            itemCode = (string)docLines.Rows[i]["ItemCode"];
                            docQty = (decimal)docLines.Rows[i]["Quantity"];
                            whsCode = "";

                            filtString = @"ItemCode = '" + itemCode + "'";
                            if (blockStockByWarehouse)
                            {
                                whsCode = (string)docLines.Rows[i]["WhsCode"];
                                filtString = filtString + @" AND WhsCode = '" + whsCode + "'";
                            }

                            foundRows = stockLines.Select(filtString);
                            if (foundRows.Length > 0)
                            {
                                finalQty = (decimal)foundRows[0]["FinalQty"];
                                if (docQty <= finalQty) //თუ ნაკლებია დოკ-ის რაოდენობა, არ შეწყდება ივენთი
                                {
                                    tempCheck = false;
                                }
                            }

                            if (tempCheck)
                            {
                                errorText = BDOSResources.getTranslate("InsufficientStockBalanceOnPostingDate") + ", " + BDOSResources.getTranslate("ItemCode") + ": " + itemCode;
                                if (blockStockByWarehouse)
                                {
                                    errorText = errorText + ", " + BDOSResources.getTranslate("Warehouse") + ": " + whsCode;
                                }
                                Program.uiApp.StatusBar.SetSystemMessage(errorText);
                                rejection = true;
                            }
                        }
                    }
                }
            }
            catch { }
            finally
            {
                oForm.Freeze(false);
            }
        }

        public static void nullsToEmptyString<T>(T instance)
        {
            var type = instance.GetType();
            var properties = type.GetProperties();
            foreach (var propertyInfo in properties)
            {
                var property = type.GetProperty(propertyInfo.Name, typeof(string));
                if (property != null)
                {
                    var value = (string)property.GetValue(instance);

                    if (value == null)
                    {
                        property.SetValue(instance, "");
                    }
                }
            }
        }

        public static string[] getNumberArrayFromText(string text)
        {
            IEnumerable<string> numberArray = Regex.Split(text, @"[^0-9\.]+").Where(c => c != "." && c.Trim() != "");
            return numberArray.ToArray();
        }

        public static bool isLocalisationAddOnConnected()
        {
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                string query = @"SELECT * 
                FROM ""SBOCOMMON"".""SEWH1"" 
                WHERE ""SEWH1"".""CompDbNam"" = '" + Program.oCompany.CompanyDB + @"' 
                AND LOCATE(""SEWH1"".""Name"",
                	 'Localisation') > 0 
                AND ""SEWH1"".""Status"" = 'Connected'";

                oRecordSet.DoQuery(query);

                if (!oRecordSet.EoF)
                {
                    return true;
                }
                return false;
            }
            catch
            {
                return false;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                oRecordSet = null;
            }
        }

        public static bool isHRAddOnConnected()
        {
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                string query = @"SELECT * 
                FROM ""SBOCOMMON"".""SEWH1"" 
                WHERE ""SEWH1"".""CompDbNam"" = '" + Program.oCompany.CompanyDB + @"' 
                AND LOCATE(""SEWH1"".""Name"",
                	 'HR') > 0 
                AND ""SEWH1"".""Status"" = 'Connected'";

                oRecordSet.DoQuery(query);

                if (!oRecordSet.EoF)
                {
                    return true;
                }
                return false;
            }
            catch
            {
                return false;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                oRecordSet = null;
            }
        }

        public static object getChildOrDbDataSourceValue(SAPbouiCOM.DBDataSource DBDataSourceTable, SAPbobsCOM.GeneralDataCollection ChildTable, DataTable DTSource, string FieldName, int index)
        {
            object fieldValue = "";

            if (DBDataSourceTable != null)
            {
                fieldValue = DBDataSourceTable.GetValue(FieldName, index);
            }
            else if (ChildTable != null)
            {
                fieldValue = ChildTable.Item(index).GetProperty(FieldName);
            }
            else if (DTSource != null)
            {
                fieldValue = DTSource.Rows[index][FieldName];
            }

            System.Type fieldType = fieldValue.GetType();

            if (fieldType.Name == "Decimal" || fieldType.Name == "Double")
            {
                return fieldValue;
            }
            else if (fieldType.Name == "DateTime")
            {
                DateTime dateValue = ((System.DateTime)(fieldValue));
                return (new DateTime(dateValue.Date.Year, dateValue.Date.Month, 1)).ToString("yyyyMMdd");
            }
            else
            {
                return fieldValue.ToString().Trim();
            }
        }

        public static string getUDFValue(string TableName, string FieldName, string KeyValue)
        {
            SAPbobsCOM.UserTable oUserTable = null;
            oUserTable = Program.oCompany.UserTables.Item(TableName);

            if (oUserTable.GetByKey(KeyValue))
            {
                string str = oUserTable.UserFields.Fields.Item(FieldName).Value;
                return str;

            }
            else
            {
                return null;
            }
        }

        public static string getDateFormat()
        {
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string DateFormat;
            try
            {
                string query = @"SELECT ""DateFormat"" FROM ""OADM""";
                oRecordSet.DoQuery(query);
                if (!oRecordSet.EoF)
                {
                    DateFormat = oRecordSet.Fields.Item("DateFormat").Value.ToString();
                    switch (DateFormat)
                    {
                        case "0": return "DD/MM/YY";
                        case "1": return "DD/MM/CCYY";
                        case "2": return "MM/DD/YY";
                        case "3": return "MM/DD/CCYY";
                        case "4": return "CCYY/MM/DD";
                        case "5": return "DD/Month/YYYY";
                        case "6": return "YY/MM/DD";
                    }
                }
                return null;
            }
            catch
            {
                return null;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                oRecordSet = null;
            }
        }

        public static List<string> checkDuplicatesInDBDataSources(SAPbouiCOM.DBDataSource oDBDataSource, Dictionary<string, SAPbouiCOM.DBDataSource> oKeysDictionary, out string errorText)
        {
            errorText = null;
            List<string> oList = new List<string>();
            Dictionary<string, object> oDictionary = null;
            List<string> duplicates = new List<string>();

            try
            {
                for (int i = 0; i < oDBDataSource.Size; i++)
                {
                    oDictionary = new Dictionary<string, object>();
                    foreach (var pair in oKeysDictionary)
                    {
                        oDictionary.Add(pair.Key, pair.Value.GetValue(pair.Key, i).Trim());
                    }
                    oList.Add(string.Join(",", oDictionary.Values));
                }

                if (oList.Count > 0)
                {
                    duplicates = oList.GroupBy(s => s).SelectMany(grp => grp.Skip(1)).ToList();
                    if (duplicates.Count > 0)
                    {
                        errorText = BDOSResources.getTranslate("ThereAreDuplicatesRowsInTheTable") + "!";
                    }
                }
                return duplicates;
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
                return duplicates;
            }
        }

        public static object getValue(string tableName, string fieldName, string filterFieldName, string filterFieldValue)
        {
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                string query = @"SELECT """ + fieldName + @""" 
                FROM """ + tableName + @""" 
                WHERE """ + filterFieldName + @""" = '" + filterFieldValue + "'";

                oRecordSet.DoQuery(query);

                if (!oRecordSet.EoF)
                {
                    return oRecordSet.Fields.Item("" + fieldName + "").Value;
                }
                return "";
            }
            catch
            {
                return "";
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                oRecordSet = null;
            }
        }

        public static decimal getInStockByWarehouseAndDate(string itemCode, string warehouse, string docDate)
        {
            SAPbobsCOM.Recordset oRecordset = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                string query = @"SELECT 
                ""ItemCode"", 
                ""Dscription"", 
                ""Warehouse"",
                SUM(""InQty"" - ""OutQty"") AS ""InStock""
                FROM ""OINM""
                WHERE
                ""ItemCode"" = '" + itemCode + @"'
                AND ""Warehouse"" = '" + warehouse + @"' AND ""DocDate"" <= '" + docDate + @"'
                GROUP BY ""ItemCode"", ""Dscription"", ""Warehouse""";

                oRecordset.DoQuery(query);
                if (!oRecordset.EoF)
                {
                    return Convert.ToDecimal(oRecordset.Fields.Item("InStock").Value);
                }
                return 0;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordset);
            }
        }
    }
}
