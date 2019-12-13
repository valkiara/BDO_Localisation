using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace BDO_Localisation_AddOn
{
    class BDOSInterestAccrual
    {
        public static void createDocumentUDO(out string errorText)
        {
            string tableName = "BDOSINAC";
            string description = "Interest Accrual Document";

            int result = UDO.addUserTable(tableName, description, SAPbobsCOM.BoUTBTableType.bott_Document, out errorText);

            if (result != 0)
            {
                return;
            }

            Dictionary<string, object> fieldskeysMap;

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "DocDate");
            fieldskeysMap.Add("TableName", "BDOSINAC");
            fieldskeysMap.Add("Description", "Posting Date");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Date);
            fieldskeysMap.Add("EditSize", 15);
            fieldskeysMap.Add("Mandatory", SAPbobsCOM.BoYesNoEnum.tYES);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "AccrMnth");
            fieldskeysMap.Add("TableName", "BDOSINAC");
            fieldskeysMap.Add("Description", "Accrual Month");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Date);
            fieldskeysMap.Add("EditSize", 15);
            fieldskeysMap.Add("Mandatory", SAPbobsCOM.BoYesNoEnum.tYES);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BankCode");
            fieldskeysMap.Add("TableName", "BDOSINAC");
            fieldskeysMap.Add("Description", "Bank Code");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 30);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "TransId");
            fieldskeysMap.Add("TableName", "BDOSINAC");
            fieldskeysMap.Add("Description", "Transaction Number");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Numeric);
            fieldskeysMap.Add("EditSize", 11);

            UDO.addUserTableFields(fieldskeysMap, out errorText);     

            tableName = "BDOSINA1";
            description = "Interest Accrual Child1";

            result = UDO.addUserTable(tableName, description, SAPbobsCOM.BoUTBTableType.bott_DocumentLines, out errorText);

            if (result != 0)
            {
                return;
            }

            //Credit Line Master Data
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "CRLNCode");
            fieldskeysMap.Add("TableName", "BDOSINA1");
            fieldskeysMap.Add("Description", "Fuel Type Code");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            //საწვავის ერთეული
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "FuUomEntry");
            fieldskeysMap.Add("TableName", "BDOSINA1");
            fieldskeysMap.Add("Description", "Fuel UoM Abs. Entry");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Numeric);
            fieldskeysMap.Add("EditSize", 11);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            //საწვავის ერთეული
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "FuUomCode");
            fieldskeysMap.Add("TableName", "BDOSINA1");
            fieldskeysMap.Add("Description", "Fuel UoM Code");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 20);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            //წვა 100 კმ-ში
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "FuPerKm");
            fieldskeysMap.Add("TableName", "BDOSINA1");
            fieldskeysMap.Add("Description", "Per 100 km");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            //წვა საათში
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "FuPerHr");
            fieldskeysMap.Add("TableName", "BDOSINA1");
            fieldskeysMap.Add("Description", "Per Hour");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            //ოდომეტრის საწყისი ჩვენება (კმ)
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "OdmtrStart");
            fieldskeysMap.Add("TableName", "BDOSINA1");
            fieldskeysMap.Add("Description", "Starting Value of Odometer");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            //ოდომეტრის საბოლოო ჩვენება (კმ)
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "OdmtrEnd");
            fieldskeysMap.Add("TableName", "BDOSINA1");
            fieldskeysMap.Add("Description", "Ending Value of Odometer");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            //ნამუშევარი საათები
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "HrsWorked");
            fieldskeysMap.Add("TableName", "BDOSINA1");
            fieldskeysMap.Add("Description", "Hours Worked");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            //ხარჯვა ნორმის მიხედვით 
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "NormCn");
            fieldskeysMap.Add("TableName", "BDOSINA1");
            fieldskeysMap.Add("Description", "Norm Consumption");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            //ფაქტიური ხარჯვა
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "ActuallyCn");
            fieldskeysMap.Add("TableName", "BDOSINA1");
            fieldskeysMap.Add("Description", "Actually Consumption");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Float);
            fieldskeysMap.Add("SubType", SAPbobsCOM.BoFldSubTypes.st_Sum);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "Dimension1");
            fieldskeysMap.Add("TableName", "BDOSINA1");
            fieldskeysMap.Add("Description", "Dimension1");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "Dimension2");
            fieldskeysMap.Add("TableName", "BDOSINA1");
            fieldskeysMap.Add("Description", "Dimension2");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "Dimension3");
            fieldskeysMap.Add("TableName", "BDOSINA1");
            fieldskeysMap.Add("Description", "Dimension3");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "Dimension4");
            fieldskeysMap.Add("TableName", "BDOSINA1");
            fieldskeysMap.Add("Description", "Dimension4");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "Dimension5");
            fieldskeysMap.Add("TableName", "BDOSINA1");
            fieldskeysMap.Add("Description", "Dimension5");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "DocEntryGI");
            fieldskeysMap.Add("TableName", "BDOSINA1");
            fieldskeysMap.Add("Description", "Goods Issue");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Numeric);
            fieldskeysMap.Add("EditSize", 11);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            //List<string> oColumnAlias = new List<string>();
            //oColumnAlias.Add("DocEntry");
            //oColumnAlias.Add("LineId");
            //oColumnAlias.Add("ItemCode");
            //UDO.AddUserKey("BDOSINA1", "DOC_ITM", oColumnAlias, out errorText);

            GC.Collect();
        }

        public static void registerUDO()
        {
            SAPbobsCOM.UserObjectsMD oUserObjectMD = null;
            SAPbobsCOM.UserObjectMD_FindColumns oUDOFind = null;
            SAPbobsCOM.UserObjectMD_FormColumns oUDOForm = null;
            SAPbobsCOM.IUserObjectMD_ChildTables oUDOChildTables = null;
            GC.Collect();
            oUserObjectMD = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);
            Marshal.ReleaseComObject(oUserObjectMD);
            oUserObjectMD = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD) as SAPbobsCOM.UserObjectsMD;
            oUDOFind = oUserObjectMD.FindColumns;
            oUDOForm = oUserObjectMD.FormColumns;
            oUDOChildTables = oUserObjectMD.ChildTables;

            var retval = oUserObjectMD.GetByKey("UDO_F_BDOSINAC_D");

            if (!retval)
            {
                oUserObjectMD.Code = "UDO_F_BDOSINAC_D";
                oUserObjectMD.Name = "Interest Accrual Document";
                oUserObjectMD.TableName = "BDOSINAC";
                oUserObjectMD.ObjectType = SAPbobsCOM.BoUDOObjType.boud_Document;
                oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO;
                oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tNO;

                //Find
                oUDOFind.ColumnAlias = "DocEntry";
                oUDOFind.ColumnDescription = "Internal Number";
                oUDOFind.Add();
                oUDOFind.ColumnAlias = "DocNum";
                oUDOFind.ColumnDescription = "Document Number";
                oUDOFind.Add();
                oUDOFind.ColumnAlias = "CreateDate";
                oUDOFind.ColumnDescription = "Create Date";
                oUDOFind.Add();
                oUDOFind.ColumnAlias = "UpdateDate";
                oUDOFind.ColumnDescription = "Update Date";
                oUDOFind.Add();
                oUDOFind.ColumnAlias = "Status";
                oUDOFind.ColumnDescription = "Status";
                oUDOFind.Add();
                oUDOFind.ColumnAlias = "Canceled";
                oUDOFind.ColumnDescription = "Canceled";
                oUDOFind.Add();
                oUDOFind.ColumnAlias = "U_DocDate";
                oUDOFind.ColumnDescription = "Posting Date";
                oUDOFind.Add();
                oUDOFind.ColumnAlias = "U_DateFrom";
                oUDOFind.ColumnDescription = "Date From";
                oUDOFind.Add();
                oUDOFind.ColumnAlias = "U_DateTo";
                oUDOFind.ColumnDescription = "Date To";
                oUDOFind.Add();
                oUDOFind.ColumnAlias = "U_PrjCode";
                oUDOFind.ColumnDescription = "Project Code";
                oUDOFind.Add();
                oUDOFind.ColumnAlias = "U_FuNrCode";
                oUDOFind.ColumnDescription = "Specification of Fuel Norm Code";
                oUDOFind.Add();
                oUDOFind.ColumnAlias = "Remark";
                oUDOFind.ColumnDescription = "Remark";
                oUDOFind.Add();

                //Form
                oUDOForm.FormColumnAlias = "DocEntry";
                oUDOForm.FormColumnDescription = "Internal Number";
                oUDOForm.Add();

                oUDOChildTables.Add();
                oUDOChildTables.SetCurrentLine(oUDOChildTables.Count - 1);
                oUDOChildTables.TableName = "BDOSINA1";
                oUDOChildTables.ObjectName = "BDOSINA1";

                if (!retval)
                {
                    if ((oUserObjectMD.Add() != 0))
                    {
                        Program.uiApp.MessageBox(Program.oCompany.GetLastErrorDescription());
                    }
                }
            }
            Marshal.ReleaseComObject(oUserObjectMD);
        }

        public static void addMenus()
        {
            try
            {
                SAPbouiCOM.MenuItem fatherMenuItem = Program.uiApp.Menus.Item("1536");
                // Add a pop-up menu item
                SAPbouiCOM.MenuCreationParams oCreationPackage = (SAPbouiCOM.MenuCreationParams)Program.uiApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                oCreationPackage.Checked = false;
                oCreationPackage.Enabled = true;
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "UDO_F_BDOSINAC_D";
                oCreationPackage.String = BDOSResources.getTranslate("InterestAccrualDocument");
                oCreationPackage.Position = fatherMenuItem.SubMenus.Count - 1;

                SAPbouiCOM.MenuItem menuItem = fatherMenuItem.SubMenus.AddEx(oCreationPackage);
            }
            catch
            {
                //Program.uiApp.MessageBox(ex.Message);
            }
        }
    }
}
