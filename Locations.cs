using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BDO_Localisation_AddOn
{
    class Locations
    {
        public static void createUserFields(out string errorText)
        {
            Dictionary<string, object> fieldskeysMap;

            fieldskeysMap = new Dictionary<string, object>();

            fieldskeysMap.Add("Name", "BDOSAddres");
            fieldskeysMap.Add("TableName", "OLCT");
            fieldskeysMap.Add("Description", "Address");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 150);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            GC.Collect();
        }
    }
}
