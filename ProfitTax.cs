using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Runtime.InteropServices;
using System.Globalization;

namespace BDO_Localisation_AddOn
{
    class ProfitTax
    {
        public static void createUDO( out string errorText)
        {
            errorText = null;

            string tableName = "BDOSPRTX";
            string description = "Profit Tax";

            SAPbobsCOM.UserObjectsMD oUserObjectsMD = null;
            oUserObjectsMD = (SAPbobsCOM.UserObjectsMD)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectsMD);

            int result = UDO.addUserTable( tableName, description, SAPbobsCOM.BoUTBTableType.bott_NoObjectAutoIncrement, out errorText);

            if (result != 0)
            {
                return;
            }

            Dictionary<string, object> fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap = new Dictionary<string, object>(); // docType
            fieldskeysMap.Add("Name", "docType");
            fieldskeysMap.Add("TableName", "BDOSPRTX");
            fieldskeysMap.Add("Description", "Posting Date");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // memo
            fieldskeysMap.Add("Name", "memo");
            fieldskeysMap.Add("TableName", "BDOSPRTX");
            fieldskeysMap.Add("Description", "Memo");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields( fieldskeysMap, out errorText);


            fieldskeysMap = new Dictionary<string, object>(); // docEntry 
            fieldskeysMap.Add("Name", "docEntry");
            fieldskeysMap.Add("TableName", "BDOSPRTX");
            fieldskeysMap.Add("Description", "Posting Date");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Numeric);
            fieldskeysMap.Add("EditSize", 11);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // docNum
            fieldskeysMap.Add("Name", "docNum");
            fieldskeysMap.Add("TableName", "BDOSPRTX");
            fieldskeysMap.Add("Description", "Posting Date");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Numeric);
            fieldskeysMap.Add("EditSize", 11);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // DocDate 
            fieldskeysMap.Add("Name", "docDate");
            fieldskeysMap.Add("TableName", "BDOSPRTX");
            fieldskeysMap.Add("Description", "Posting Date");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Date);
            fieldskeysMap.Add("EditSize", 15);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // დაბეგვრის ობიექტი
            fieldskeysMap.Add("Name", "prBase"); 
            fieldskeysMap.Add("TableName", "BDOSPRTX");
            fieldskeysMap.Add("Description", "Profit Base");
            fieldskeysMap.Add("EditSize", 15);
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // დებეტის
            fieldskeysMap.Add("Name", "dbAcct");
            fieldskeysMap.Add("TableName", "BDOSPRTX");
            fieldskeysMap.Add("Description", "Debit Account");
            fieldskeysMap.Add("EditSize", 15);
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // კრედიტის
            fieldskeysMap.Add("Name", "crAcct");
            fieldskeysMap.Add("TableName", "BDOSPRTX");
            fieldskeysMap.Add("Description", "Credit Account");
            fieldskeysMap.Add("EditSize", 15);
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            //Dictionary<string, string> listValidValuesDict = new Dictionary<string, string>();
            //listValidValuesDict = new Dictionary<string, string>();
            //listValidValuesDict.Add("Uncrediting", "Uncrediting");
            //listValidValuesDict.Add("Accrual", "Accrual");

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "txType"); //დაბეგვრის ტიპი
            fieldskeysMap.Add("TableName", "BDOSPRTX");
            fieldskeysMap.Add("Description", "Base Type");
            fieldskeysMap.Add("EditSize", 50);
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            //fieldskeysMap.Add("ValidValues", listValidValuesDict);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "amtTx"); //დასაბეგრი თანხა
            fieldskeysMap.Add("TableName", "BDOSPRTX");
            fieldskeysMap.Add("Description", "Amount Taxable");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "amtPrTx"); //მოგების გადასახადი
            fieldskeysMap.Add("TableName", "BDOSPRTX");
            fieldskeysMap.Add("Description", "Profit Tax Amount");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields( fieldskeysMap, out errorText);


        }

        public static DataTable ProfitTaxTable()
        {
            DataTable reLines = new DataTable();
            reLines.Columns.Add("docEntry", typeof(int));
            reLines.Columns.Add("docNum", typeof(int));
            reLines.Columns.Add("docDate", typeof(DateTime));
            reLines.Columns.Add("debitAccount", typeof(string));
            reLines.Columns.Add("creditAccount", typeof(string));
            reLines.Columns.Add("prBase", typeof(string));
            reLines.Columns.Add("txType", typeof(string));
            reLines.Columns.Add("amtTx", typeof(float));
            reLines.Columns.Add("amtPrTx", typeof(float));

            return reLines;
        }

        public static void AddRecord( DataTable reLines, string docType, string memo, out string errorText)
        {
            errorText = null;
            int returnCode;

            SAPbobsCOM.UserTable oUserTable = Program.oCompany.UserTables.Item("BDOSPRTX");
            DataRow reLine;

            try
            {
                for (int i = 0; i < reLines.Rows.Count; i++)
                {
                    reLine = reLines.Rows[i];

                    oUserTable.UserFields.Fields.Item("U_docType").Value = docType;
                    oUserTable.UserFields.Fields.Item("U_memo").Value = memo;
                    oUserTable.UserFields.Fields.Item("U_docEntry").Value = reLine["docEntry"].ToString();
                    oUserTable.UserFields.Fields.Item("U_docNum").Value = reLine["docNum"].ToString();
                    oUserTable.UserFields.Fields.Item("U_docDate").Value = reLine["docDate"];                  
                    oUserTable.UserFields.Fields.Item("U_dbAcct").Value = reLine["debitAccount"].ToString();
                    oUserTable.UserFields.Fields.Item("U_crAcct").Value = reLine["creditAccount"].ToString();
                    oUserTable.UserFields.Fields.Item("U_prBase").Value = reLine["prBase"].ToString();
                    oUserTable.UserFields.Fields.Item("U_txType").Value = reLine["txType"].ToString();

                    oUserTable.UserFields.Fields.Item("U_amtTx").Value = Convert.ToDouble( CommonFunctions.roundAmountByGeneralSettings(Convert.ToDecimal(reLine["amtTx"]), "Sum"));
                    oUserTable.UserFields.Fields.Item("U_amtPrTx").Value = Convert.ToDouble(CommonFunctions.roundAmountByGeneralSettings( Convert.ToDecimal(reLine["amtPrTx"]), "Sum"));

                    returnCode = oUserTable.Add();

                    if (returnCode != 0)
                    {
                        int errCode;
                        string errMsg;

                        Program.oCompany.GetLastError(out errCode, out errMsg);
                        errorText = "Error description : " + errMsg + "! Code : " + errCode;
                    }
                }
            }
            catch(Exception ex)
            {
                errorText = ex.Message;

            }
            finally
            {
                Marshal.ReleaseComObject(oUserTable);
                oUserTable = null;
            }
        }

        public static decimal GetProfitTaxRate()
        {
            NumberFormatInfo Nfi = new NumberFormatInfo() { NumberDecimalSeparator = "." };
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string query = @"SELECT ""U_BDO_PrTxRt"" AS Rate FROM ""OADM""";

            oRecordSet.DoQuery(query);

            if (!oRecordSet.EoF)
            {
                return Convert.ToDecimal(oRecordSet.Fields.Item("Rate").Value);
            }
            else
            {
                return 0;
            }
        }

        public static bool TaxAccountsIsEmpty()
        {
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string query = @"SELECT ""U_BDO_CapAcc"", ""U_BDO_TaxAcc"" FROM ""OADM"" WHERE ""U_BDO_CapAcc"" = '' OR ""U_BDO_TaxAcc"" = ''";

            oRecordSet.DoQuery(query);

            if (!oRecordSet.EoF)
            {
                return true;
            }
            else
            {
                return false ;
            }
        }
 
        public static bool ProfitTaxTypeIsSharing()
        {
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string query = @"SELECT ""U_BDO_TaxTyp"" FROM ""OADM"" WHERE ""U_BDO_TaxTyp"" = '1'";

            oRecordSet.DoQuery(query);

            if (!oRecordSet.EoF)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

 
    }
}
