using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BDO_Localisation_AddOn
{
    public static class BDOSTablesLog
    {
        public struct LogRecord
        {
            public string ExecutionID, TableName, FieldName, UDOName;

            public DateTime Date;        
        }

        public static void CreateTable(out string errorText)
        {
            errorText = null; 
            
            string tableName = "BDOSLOGS";
            string description = "Field and Table Logs";
            
            /*var oUserObjectsMD = (SAPbobsCOM.UserObjectsMD)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectsMD);*/

            int result = UDO.addUserTable( tableName, description, SAPbobsCOM.BoUTBTableType.bott_NoObject, out errorText);

            if (result != 0)
            {         
                return;
            }

            Dictionary<string, object> fieldskeysMap = new Dictionary<string, object>();

            fieldskeysMap = new Dictionary<string, object>(); // 
            fieldskeysMap.Add("Name", "BDOSExID");
            fieldskeysMap.Add("TableName", "BDOSLOGS");
            fieldskeysMap.Add("Description", "Execution ID"); //ISO Date with Second Precision , ex: 2018-12-03T07:12:33.167Z
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 250);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // 
            fieldskeysMap.Add("Name", "BDOSTbNm");
            fieldskeysMap.Add("TableName", "BDOSLOGS");
            fieldskeysMap.Add("Description", "Table Name");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 250);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // 
            fieldskeysMap.Add("Name", "BDOSFdNm");
            fieldskeysMap.Add("TableName", "BDOSLOGS");
            fieldskeysMap.Add("Description", "Field Name");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 250);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // 
            fieldskeysMap.Add("Name", "BDOSUDO");
            fieldskeysMap.Add("TableName", "BDOSLOGS");
            fieldskeysMap.Add("Description", "UDO Name");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 250);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // 
            fieldskeysMap.Add("Name", "BDOSStts");
            fieldskeysMap.Add("TableName", "BDOSLOGS");
            fieldskeysMap.Add("Description", "Status (Success = Y)");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 250);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // 
            fieldskeysMap.Add("Name", "BDOSPcUsr");
            fieldskeysMap.Add("TableName", "BDOSLOGS");
            fieldskeysMap.Add("Description", "PC User");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 250);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // 
            fieldskeysMap.Add("Name", "BDOSB1Usr");
            fieldskeysMap.Add("TableName", "BDOSLOGS");
            fieldskeysMap.Add("Description", "B1 User");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 250);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // REG_DT
            fieldskeysMap.Add("Name", "BDOSDt");
            fieldskeysMap.Add("TableName", "BDOSLOGS");
            fieldskeysMap.Add("Description", "Add-on execution date");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Date);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSStCd");
            fieldskeysMap.Add("TableName", "BDOSLOGS");
            fieldskeysMap.Add("Description", "Status Code");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Numeric);
            fieldskeysMap.Add("EditSize", 11);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>(); // 
            fieldskeysMap.Add("Name", "BDOSDesc");
            fieldskeysMap.Add("TableName", "BDOSLOGS");
            fieldskeysMap.Add("Description", "Description");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Memo);
            fieldskeysMap.Add("EditSize", 500);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            Program.UserDefinedTablesCurrentCompany = UDO.UserDefinedTablesCurrentCompany();
            Program.UserDefinedFieldsCurrentCompany = UDO.UserDefinedFieldsCurrentCompany();

        }

    }
}
