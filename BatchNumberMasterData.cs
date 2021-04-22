using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BDO_Localisation_AddOn
{
    class BatchNumberMasterData
    {
        public static void createUserFields()
        {
            Dictionary<string, object>  fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "CaptDate");
            fieldskeysMap.Add("TableName", "OBTN");
            fieldskeysMap.Add("Description", "Capitalization Date");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Date);

            UDO.addUserTableFields(fieldskeysMap, out var errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "HistAPC");
            fieldskeysMap.Add("TableName", "OBTN");
            fieldskeysMap.Add("Description", "Historical Acquisition & Production Cost");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "APC");
            fieldskeysMap.Add("TableName", "OBTN");
            fieldskeysMap.Add("Description", "Acquisition & Production Cost");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "HistAccmDprAmt");
            fieldskeysMap.Add("TableName", "OBTN");
            fieldskeysMap.Add("Description", "Historical Accumulated Depreciation Amount");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "HistSupplier");
            fieldskeysMap.Add("TableName", "OBTN");
            fieldskeysMap.Add("Description", "Historical Supplier");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 15);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "UsefulLife");
            fieldskeysMap.Add("TableName", "OBTN");
            fieldskeysMap.Add("Description", "Useful Life");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Numeric);

            UDO.addUserTableFields(fieldskeysMap, out errorText);
        }
    }
}
