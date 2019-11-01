using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BDO_Localisation_AddOn
{
    class BDOSApprovalProcedures
    {
        public static void createMasterDataUDO(out string errorText)
        {
            string tableName = "BDOSAPRP";
            string description = "Approval Procedures";

            int result = UDO.addUserTable(tableName, description, SAPbobsCOM.BoUTBTableType.bott_MasterData, out errorText);

            if (result != 0)
            {
                return;
            }

            Dictionary<string, object> fieldskeysMap;
          
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "AllProject");
            fieldskeysMap.Add("TableName", "BDOSAPRP");
            fieldskeysMap.Add("Description", "All Project");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 1);
            fieldskeysMap.Add("DefaultValue", "N");

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "ObjectCode");
            fieldskeysMap.Add("TableName", "BDOSAPRP");
            fieldskeysMap.Add("Description", "User Defined Object Code");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 20);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            tableName = "BDOSAPR1";
            description = "Approval Procedures Child1";

            result = UDO.addUserTable(tableName, description, SAPbobsCOM.BoUTBTableType.bott_MasterDataLines, out errorText);

            if (result != 0)
            {
                return;
            }

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "Stage");
            fieldskeysMap.Add("TableName", "BDOSAPR1");
            fieldskeysMap.Add("Description", "Stage");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 100);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "OrderId");
            fieldskeysMap.Add("TableName", "BDOSAPR1");
            fieldskeysMap.Add("Description", "Order Id");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Numeric);
            fieldskeysMap.Add("EditSize", 11);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "MaxReqr");
            fieldskeysMap.Add("TableName", "BDOSAPR1");
            fieldskeysMap.Add("Description", "No. of Authorizers");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Numeric);
            fieldskeysMap.Add("EditSize", 11);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            tableName = "BDOSAPR2";
            description = "Approval Procedures Child2";

            result = UDO.addUserTable(tableName, description, SAPbobsCOM.BoUTBTableType.bott_MasterDataLines, out errorText);

            if (result != 0)
            {
                return;
            }

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "PrjCode");
            fieldskeysMap.Add("TableName", "BDOSAPR2");
            fieldskeysMap.Add("Description", "Project Code");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 20);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            GC.Collect();
        }

        public static void registerUDO()
        {
            SAPbobsCOM.UserObjectsMD oUserObjectMD = null;
            SAPbobsCOM.UserObjectMD_FindColumns oUDOFind = null;
            SAPbobsCOM.UserObjectMD_FormColumns oUDOForm = null;
            SAPbobsCOM.IUserObjectMD_ChildTables oUDOChildTables = null;
            GC.Collect();
            oUserObjectMD = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD) as SAPbobsCOM.UserObjectsMD;
            oUDOFind = oUserObjectMD.FindColumns;
            oUDOForm = oUserObjectMD.FormColumns;
            oUDOChildTables = oUserObjectMD.ChildTables;

            var retval = oUserObjectMD.GetByKey("UDO_F_BDOSAPRP_D");

            if (!retval)
            {
                oUserObjectMD.Code = "UDO_F_BDOSAPRP_D";
                oUserObjectMD.Name = "Approval Procedures";
                oUserObjectMD.TableName = "BDOSAPRP";
                oUserObjectMD.ObjectType = SAPbobsCOM.BoUDOObjType.boud_MasterData;
                oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tYES;

                //Find
                oUDOFind.ColumnAlias = "Code";
                oUDOFind.ColumnDescription = "Code";
                oUDOFind.Add();
                oUDOFind.ColumnAlias = "Name";
                oUDOFind.ColumnDescription = "Name";
                oUDOFind.Add();
                oUDOFind.ColumnAlias = "U_AllProject";
                oUDOFind.ColumnDescription = "All Project";
                oUDOFind.Add();

                //Form
                oUDOForm.FormColumnAlias = "Code";
                oUDOForm.FormColumnDescription = "Code";
                oUDOForm.Editable = SAPbobsCOM.BoYesNoEnum.tYES;
                oUDOForm.Add();
                oUDOForm.FormColumnAlias = "Name";
                oUDOForm.FormColumnDescription = "Name";
                oUDOForm.Editable = SAPbobsCOM.BoYesNoEnum.tYES;
                oUDOForm.Add();

                oUDOChildTables.Add();
                oUDOChildTables.SetCurrentLine(oUDOChildTables.Count - 1);
                oUDOChildTables.TableName = "BDOSAPR1";
                oUDOChildTables.ObjectName = "BDOSAPR1";

                oUDOChildTables.Add();
                oUDOChildTables.SetCurrentLine(oUDOChildTables.Count - 1);
                oUDOChildTables.TableName = "BDOSAPR2";
                oUDOChildTables.ObjectName = "BDOSAPR2";

                if (!retval)
                {
                    if ((oUserObjectMD.Add() != 0))
                    {
                        Program.uiApp.MessageBox(Program.oCompany.GetLastErrorDescription());
                    }
                }
            }
        }
    }
}
