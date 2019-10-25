using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BDO_Localisation_AddOn
{
    class BDOSApprovalStages
    {
        public static void createNoObjectUDO(out string errorText)
        {
            string tableName = "BDOSAPRS";
            string description = "Approval Stages";

            int result = UDO.addUserTable(tableName, description, SAPbobsCOM.BoUTBTableType.bott_NoObjectAutoIncrement, out errorText);

            if (result != 0)
            {
                return;
            }

            Dictionary<string, object> fieldskeysMap;

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "AprPrcCode");
            fieldskeysMap.Add("TableName", "BDOSAPRS");
            fieldskeysMap.Add("Description", "Approval Procedure Code");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "StgLineId");
            fieldskeysMap.Add("TableName", "BDOSAPRS");
            fieldskeysMap.Add("Description", "Stage Line Id");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Numeric);
            fieldskeysMap.Add("EditSize", 11);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "UserCode");
            fieldskeysMap.Add("TableName", "BDOSAPRS");
            fieldskeysMap.Add("Description", "Authorizer Code");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 25);

            UDO.addUserTableFields(fieldskeysMap, out errorText);
            
            GC.Collect();
        }
    }
}
