using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BDO_Localisation_AddOn
{
    class BDOSTasksForApproval
    {
        public static void createNoObjectUDO(out string errorText)
        {
            string tableName = "BDOSAPRT";
            string description = "Tasks for Approval";

            int result = UDO.addUserTable(tableName, description, SAPbobsCOM.BoUTBTableType.bott_NoObjectAutoIncrement, out errorText);

            if (result != 0)
            {
                return;
            }

            Dictionary<string, object> fieldskeysMap;

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "Addressee");
            fieldskeysMap.Add("TableName", "BDOSAPRT");
            fieldskeysMap.Add("Description", "Addressee Id");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 25);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "Status");
            fieldskeysMap.Add("TableName", "BDOSAPRT");
            fieldskeysMap.Add("Description", "Status");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 1);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "AprPrcCode");
            fieldskeysMap.Add("TableName", "BDOSAPRT");
            fieldskeysMap.Add("Description", "Approval Procedure Code");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "UserCode");
            fieldskeysMap.Add("TableName", "BDOSAPRT");
            fieldskeysMap.Add("Description", "Authorizer Code");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 25);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "Date");
            fieldskeysMap.Add("TableName", "BDOSAPRT");
            fieldskeysMap.Add("Description", "Date");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Date);
            fieldskeysMap.Add("EditSize", 15);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "ExecDate");
            fieldskeysMap.Add("TableName", "BDOSAPRT");
            fieldskeysMap.Add("Description", "Execute Date");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Date);
            fieldskeysMap.Add("EditSize", 15);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "DocCode");
            fieldskeysMap.Add("TableName", "BDOSAPRT");
            fieldskeysMap.Add("Description", "Document Code");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 100);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "ObjectCode");
            fieldskeysMap.Add("TableName", "BDOSAPRT");
            fieldskeysMap.Add("Description", "User Defined Object Code");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 20);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "Remarks");
            fieldskeysMap.Add("TableName", "BDOSAPRT");
            fieldskeysMap.Add("Description", "Remarks");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 254);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            GC.Collect();
        }
    }
}
