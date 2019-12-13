using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace BDO_Localisation_AddOn
{
    static partial class UDO
    {
        public static void allUDOForAddOn(  out string errorText)
        {
            SAPbobsCOM.UserObjectsMD oUserObjectsMD = null;
            oUserObjectsMD = (SAPbobsCOM.UserObjectsMD)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);

            BDOSInternetBankingIntegrationServicesRules.createMasterDataUDO( out errorText);
            if (oUserObjectsMD.GetByKey("UDO_F_BDOSINTR_D") == false)
            {
                BDOSInternetBankingIntegrationServicesRules.registerUDO( out errorText);
            }
            
            BDOSItemCategories.createMasterDataUDO( out errorText);
            if (oUserObjectsMD.GetByKey("UDO_F_BDOSITMCTG_D") == false)
            {
                BDOSItemCategories.registerUDO( out errorText);
            }

            BDO_ProfitTaxBaseType.createMasterDataUDO( out errorText);
            if (oUserObjectsMD.GetByKey("UDO_F_BDO_PTBT_D") == false)
            {
                BDO_ProfitTaxBaseType.registerUDO( out errorText);
            }

            BDO_ProfitTaxBase.createMasterDataUDO( out errorText);
            if (oUserObjectsMD.GetByKey("UDO_F_BDO_PTBS_D") == false)
            {
                BDO_ProfitTaxBase.registerUDO( out errorText);
            }

            BDO_Drivers.createMasterDataUDO( out errorText);
            if (oUserObjectsMD.GetByKey("UDO_F_BDO_DRVS_D") == false)
            {
                BDO_Drivers.registerUDO( out errorText);
            }

            BDOSFuelTypes.createMasterDataUDO(out errorText);
            if (oUserObjectsMD.GetByKey("UDO_F_BDOSFUTP_D") == false)
            {
                BDOSFuelTypes.registerUDO();
            }

            BDOSFuelCriteria.createMasterDataUDO(out errorText);
            if (oUserObjectsMD.GetByKey("UDO_F_BDOSFUCR_D") == false)
            {
                BDOSFuelCriteria.registerUDO();
            }

            BDOSFuelNormSpecification.createMasterDataUDO(out errorText);
            if (oUserObjectsMD.GetByKey("UDO_F_BDOSFUNR_D") == false)
            {
                BDOSFuelNormSpecification.registerUDO();
            }

            BDOSFuelConsumptionAct.createDocumentUDO(out errorText);
            if (oUserObjectsMD.GetByKey("UDO_F_BDOSFUCN_D") == false)
            {
                BDOSFuelConsumptionAct.registerUDO();
            }

            BDO_Vehicles.createMasterDataUDO( out errorText);
            if (oUserObjectsMD.GetByKey("UDO_F_BDO_VECL_D") == false)
            {
                BDO_Vehicles.registerUDO( out errorText);
            }

            BDO_Waybills.createDocumentUDO( out errorText);
            if (oUserObjectsMD.GetByKey("UDO_F_BDO_WBLD_D") == false)
            {
                BDO_Waybills.registerUDO( out errorText);
            }

            BDO_TaxInvoiceReceived.createDocumentUDO( out errorText);
            if (oUserObjectsMD.GetByKey("UDO_F_BDO_TAXR_D") == false)
            {
                BDO_TaxInvoiceReceived.registerUDO( out errorText);
            }

            BDOSARDownPaymentVATAccrual.createDocumentUDO( out errorText);
            if (oUserObjectsMD.GetByKey("UDO_F_BDO_ARDPV_D") == false)
            {
                BDOSARDownPaymentVATAccrual.registerUDO( out errorText);
            }

            BDO_TaxInvoiceSent.createDocumentUDO( out errorText);
            if (oUserObjectsMD.GetByKey("UDO_F_BDO_TAXS_D") == false)
            {
                BDO_TaxInvoiceSent.registerUDO( out errorText);
            }

            BDO_ProfitTaxAccrual.createDocumentUDO( out errorText);
            if (oUserObjectsMD.GetByKey("UDO_F_BDO_TAXP_D") == false)
            {
                BDO_ProfitTaxAccrual.registerUDO( out errorText);
            }

            BDOSFixedAssetTransfer.createDocumentUDO(out errorText);
            if (oUserObjectsMD.GetByKey("UDO_F_BDOSFASTRD_D") == false)
            {
                BDOSFixedAssetTransfer.registerUDO(out errorText);
            }

            BDOSDepreciationAccrualDocument.createDocumentUDO(out errorText);
            if (oUserObjectsMD.GetByKey("UDO_F_BDOSDEPACR_D") == false)
            {
                BDOSDepreciationAccrualDocument.registerUDO(out errorText);
            }

            BDOSCreditLine.createMasterDataUDO(out errorText);
            if (!oUserObjectsMD.GetByKey("UDO_F_BDOSCRLN_D"))
            {
                BDOSCreditLine.registerUDO();
            }

            //Persona Tables
            BDOSApprovalProcedures.createMasterDataUDO(out errorText);
            if (oUserObjectsMD.GetByKey("UDO_F_BDOSAPRP_D") == false)
            {
                BDOSApprovalProcedures.registerUDO();
            }
            BDOSApprovalStages.createNoObjectUDO(out errorText);
            BDOSTasksForApproval.createNoObjectUDO(out errorText);
            //Persona Tables 

            //მოგების გადასახადის ცხრილი (ბევრი დოკუმენტი გააკეთებს ჩანაწერებს)
            ProfitTax.createUDO( out errorText);

            BDOSVATReconcilationWizard.createUDO( out errorText);

            //მიღებული ფაქტურების ცხრილი (Crystal Report - სთვის)
            BDOSTaxAnalysisReceived.createUDO( out errorText);

            //გაცემული ფაქტურების ცხრილი (Crystal Report - სთვის)
            BDOSTaxAnalysissSent.createUDO( out errorText);

            //----------------------------------------------->ინტერნეტბანკი<-----------------------------------------------
            string tableName = "BDO_INTB";
            string description = "Internet Banking (WSDL)";

            oUserObjectsMD = null;
            oUserObjectsMD = (SAPbobsCOM.UserObjectsMD)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectsMD);

            int result = UDO.addUserTable( tableName, description, SAPbobsCOM.BoUTBTableType.bott_NoObjectAutoIncrement, out errorText);

            if (result != 0)
            {
                return;
            }

            Dictionary<string, string> listValidValuesDict = new Dictionary<string, string>();
            listValidValuesDict.Add("TBC", "TBC (Web-Service)");
            listValidValuesDict.Add("BOG", "BOG (Web-Service)");

            Dictionary<string, object> fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "program");
            fieldskeysMap.Add("TableName", "BDO_INTB");
            fieldskeysMap.Add("Description", "program");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);
            fieldskeysMap.Add("ValidValues", listValidValuesDict);
            fieldskeysMap.Add("Mandatory", SAPbobsCOM.BoYesNoEnum.tYES);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            bool resultTmp = addNewValidValuesUserFieldsMD( "@BDO_INTB", "program", "BOG", "BOG (Web-Service)", out errorText);

            listValidValuesDict = new Dictionary<string, string>();
            listValidValuesDict.Add("test", "test");
            listValidValuesDict.Add("real", "real");

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "mode");
            fieldskeysMap.Add("TableName", "BDO_INTB");
            fieldskeysMap.Add("Description", "mode");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);
            fieldskeysMap.Add("ValidValues", listValidValuesDict);
            fieldskeysMap.Add("Mandatory", SAPbobsCOM.BoYesNoEnum.tYES);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "WSDL");
            fieldskeysMap.Add("TableName", "BDO_INTB");
            fieldskeysMap.Add("Description", "WSDL");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 254);
            fieldskeysMap.Add("Mandatory", SAPbobsCOM.BoYesNoEnum.tYES);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "ID");
            fieldskeysMap.Add("TableName", "BDO_INTB");
            fieldskeysMap.Add("Description", "ID");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "URL");
            fieldskeysMap.Add("TableName", "BDO_INTB");
            fieldskeysMap.Add("Description", "URL");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 254);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "port");
            fieldskeysMap.Add("TableName", "BDO_INTB");
            fieldskeysMap.Add("Description", "port");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Numeric);
            fieldskeysMap.Add("EditSize", 11);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            SAPbobsCOM.UserKeysMD oUserKeysMD = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserKeys);

            oUserKeysMD.TableName = "BDO_INTB";
            oUserKeysMD.KeyName = "program";
            oUserKeysMD.Elements.ColumnAlias = "program";
            oUserKeysMD.Unique = SAPbobsCOM.BoYesNoEnum.tYES;

            int returnCode = oUserKeysMD.Add();

            Marshal.ReleaseComObject(oUserKeysMD);

            addUpdateRecord_BDO_INTB();

            if (returnCode != 0)
            {
                int errCode;
                string errMsg;

                Program.oCompany.GetLastError(out errCode, out errMsg);
                errorText = BDOSResources.getTranslate("ErrorDescription") + " : " + errMsg + "! " + BDOSResources.getTranslate("Code") + " : " + errCode;
                
            }

            //----------------------------------------------->ედონის ვერსია<-----------------------------------------------
            tableName = "BDOSAVRS";
            description = "AddOn version";

            oUserObjectsMD = null;
            oUserObjectsMD = (SAPbobsCOM.UserObjectsMD)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectsMD);

            result = UDO.addUserTable( tableName, description, SAPbobsCOM.BoUTBTableType.bott_NoObject, out errorText);
            
            if (result != 0)
            { }

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "Version");
            fieldskeysMap.Add("TableName", "BDOSAVRS");
            fieldskeysMap.Add("Description", "ID");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            //ProgressBarForm.Value++;
        }

        public static void addUpdateRecord_BDO_INTB()
        {
            string wsdl;
            string mode;
            string program;
            int returnCode;

            string url;
            int port;

            SAPbobsCOM.UserTable oUserTable = Program.oCompany.UserTables.Item("BDO_INTB");
            SAPbobsCOM.ValidValues oValidValues = oUserTable.UserFields.Fields.Item("U_program").ValidValues;

            try
            {
                for (int i = 0; i < oValidValues.Count; i++)
                {
                    program = oValidValues.Item(i).Value;
                    if (program == "TBC")
                    {
                        wsdl = "https://test.tbconline.ge/dbi/dbiService"; //"test.tbconline.ge";
                        mode = "test";

                        oUserTable.UserFields.Fields.Item("U_program").Value = program;
                        oUserTable.UserFields.Fields.Item("U_mode").Value = mode;
                        oUserTable.UserFields.Fields.Item("U_WSDL").Value = wsdl;

                        //oUserTable.Code = "TBC";
                        //oUserTable.Name = "TBC";
                        returnCode = oUserTable.Add();

                        if (returnCode != 0)
                        {
                            int errCode;
                            string errMsg;

                            Program.oCompany.GetLastError(out errCode, out errMsg);
                            string errorText = "Error description : " + errMsg + "! Code : " + errCode;
                        }
                    }
                    else if (program == "BOG")
                    {
                        wsdl = "https://cib2-web-dev.bog.ge"; //91.209.131.231
                        mode = "test";
                        url = "https://cib2-web-dev.bog.ge"; //91.209.131.231
                        port = 8090;

                        oUserTable.UserFields.Fields.Item("U_program").Value = program;
                        oUserTable.UserFields.Fields.Item("U_mode").Value = mode;
                        oUserTable.UserFields.Fields.Item("U_WSDL").Value = wsdl;
                        //oUserTable.UserFields.Fields.Item("U_ID").Value = id;
                        oUserTable.UserFields.Fields.Item("U_URL").Value = url;
                        oUserTable.UserFields.Fields.Item("U_port").Value = port;
                        //oUserTable.Code = "BOG";
                        //oUserTable.Name = "BOG";
                        returnCode = oUserTable.Add();

                        if (returnCode != 0)
                        {
                            int errCode;
                            string errMsg;

                            Program.oCompany.GetLastError(out errCode, out errMsg);
                            string errorText = "Error description : " + errMsg + "! Code : " + errCode;
                        }
                    }
                }
            }
            catch
            {

            }
            finally
            {
                Marshal.ReleaseComObject(oUserTable);
                oUserTable = null;
                Marshal.ReleaseComObject(oValidValues);
                oValidValues = null;
            }
        }

        public static int addUserTable( string tableName, string description, SAPbobsCOM.BoUTBTableType type, out string errorText)
        {
            errorText = null;

            /*if (userTableExist( tableName) == true)*/
            if (UserDefinedTableExists(tableName))
            {
                return 0;
            }
            
            SAPbobsCOM.UserTablesMD oUserTablesMD = ((SAPbobsCOM.UserTablesMD)(Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)));
            oUserTablesMD.TableName = tableName;
            oUserTablesMD.TableDescription = description;
            oUserTablesMD.TableType = type;

            int errCode;
            string errMsg;

            try
            {
                int returnCode = oUserTablesMD.Add();

                if (returnCode != 0)
                {
                    Program.oCompany.GetLastError(out errCode, out errMsg);
                    errorText = BDOSResources.getTranslate("ErrorOfTableAdd") + BDOSResources.getTranslate("ErrorDescription") + ":" + errMsg + "! " + BDOSResources.getTranslate("Code") + " : " + errCode + "!" + BDOSResources.getTranslate("Table") + ": " + "\"" + tableName + "\"";
                    return returnCode;
                }
            }
            catch (Exception ex)
            {
                Program.oCompany.GetLastError(out errCode, out errMsg);
                errorText = BDOSResources.getTranslate("ErrorOfTableAdd") + BDOSResources.getTranslate("ErrorDescription") + ":" + errMsg + "! " + BDOSResources.getTranslate("Code") + " : " + errCode + "!" + BDOSResources.getTranslate("Table") + ": " + "\"" + tableName + "\"" + BDOSResources.getTranslate("OtherInfo") + ": " + ex.Message;
                return -1;
            }
            finally
            {
                Marshal.ReleaseComObject(oUserTablesMD);
                GC.Collect();
            }

            return 0;
        }

        private static bool userTableExist( string tableName)
        {
            SAPbobsCOM.UserTablesMD oUserTablesMD = null;
            bool boolIdent = false;
            oUserTablesMD = ((SAPbobsCOM.UserTablesMD)(Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)));
            boolIdent = oUserTablesMD.GetByKey(tableName);
            Marshal.ReleaseComObject(oUserTablesMD);
            GC.Collect();

            return (boolIdent);
        }

        private static bool userTableFieldsExist( string tableName, int fieldID)
        {
            SAPbobsCOM.UserFieldsMD oUserFieldsMD = null;
            bool boolIdent = false;
            oUserFieldsMD = ((SAPbobsCOM.UserFieldsMD)(Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)));
            boolIdent = oUserFieldsMD.GetByKey(tableName, fieldID);
            Marshal.ReleaseComObject(oUserFieldsMD);
            GC.Collect();

            return (boolIdent);
        }

        public static void addUserTableFields( Dictionary<string, object> fieldskeysMap, out string errorText)
        {
            errorText = null;
           
            object propertyValue = null;

            string name = "";
            string tableName = "";

            if (fieldskeysMap.TryGetValue("Name", out propertyValue) == true) //8 characters
            {
                name = propertyValue.ToString().Trim();
            }
            if (fieldskeysMap.TryGetValue("TableName", out propertyValue) == true)
            {
                tableName = propertyValue.ToString().Trim();
            }

            if (UserDefinedFieldExists(tableName, name))
            {
                return;
            }

            SAPbobsCOM.UserFieldsMD oUserFieldsMD = null;
            oUserFieldsMD = ((SAPbobsCOM.UserFieldsMD)(Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)));

            SAPbobsCOM.ValidValuesMD validValuesCOM = null;

            oUserFieldsMD.Name = name;
            oUserFieldsMD.TableName = tableName;

            if (fieldskeysMap.TryGetValue("DefaultValue", out propertyValue) == true)
            {
                oUserFieldsMD.DefaultValue = propertyValue.ToString();
            }
            if (fieldskeysMap.TryGetValue("Description", out propertyValue) == true) //30 characters
            {
                oUserFieldsMD.Description = propertyValue.ToString();
            }
            if (fieldskeysMap.TryGetValue("EditSize", out propertyValue) == true)
            {
                oUserFieldsMD.EditSize = Convert.ToInt32(propertyValue);
            }
            if (fieldskeysMap.TryGetValue("LinkedSystemObject", out propertyValue) == true)
            {
                //oUserFieldsMD.LinkedSystemObject = (SAPbobsCOM.BoObjectTypes)propertyValue;
            }
            if (fieldskeysMap.TryGetValue("LinkedTable", out propertyValue) == true)
            {
                oUserFieldsMD.LinkedTable = propertyValue.ToString();
            }
            if (fieldskeysMap.TryGetValue("LinkedUDO", out propertyValue) == true)
            {
                oUserFieldsMD.LinkedUDO = propertyValue.ToString();
            }
            if (fieldskeysMap.TryGetValue("Mandatory", out propertyValue) == true)
            {
                oUserFieldsMD.Mandatory = (SAPbobsCOM.BoYesNoEnum)propertyValue;
            }
            if (fieldskeysMap.TryGetValue("Size", out propertyValue) == true)
            {
                oUserFieldsMD.Size = Convert.ToInt32(propertyValue);
            }
            if (fieldskeysMap.TryGetValue("SubType", out propertyValue) == true)
            {
                oUserFieldsMD.SubType = (SAPbobsCOM.BoFldSubTypes)propertyValue;
            }
            if (fieldskeysMap.TryGetValue("Type", out propertyValue) == true)
            {
                oUserFieldsMD.Type = (SAPbobsCOM.BoFieldTypes)propertyValue;
            }
            if (fieldskeysMap.TryGetValue("ValidValues", out propertyValue) == true)
            {
                if (propertyValue.GetType() == typeof(Dictionary<string, string>))
                {
                    validValuesCOM = oUserFieldsMD.ValidValues;
                    Dictionary<string, string> listValidValues = (Dictionary<string, string>)propertyValue;

                    foreach (KeyValuePair<string, string> keyValue in listValidValues)
                    {
                        validValuesCOM.Description = keyValue.Value;
                        validValuesCOM.Value = keyValue.Key;
                        validValuesCOM.Add();
                    }
                    listValidValues = null;
                }
                else
                {
                    List<string> listValidValues = (List<string>)propertyValue;
                    validValuesCOM = oUserFieldsMD.ValidValues;
                    for (int i = 0; i < listValidValues.Count(); i++)
                    {
                        validValuesCOM.Description = listValidValues[i];
                        validValuesCOM.Value = i == 0 & listValidValues[i] == "" ? "-1" : i.ToString();
                        validValuesCOM.Add();
                    }
                    listValidValues = null;
                }
            }

            int errCode;
            string errMsg;

            try
            {
                int returnCode = oUserFieldsMD.Add();
                if (returnCode != 0)
                {
                Program.oCompany.GetLastError(out errCode, out errMsg);
                errorText = BDOSResources.getTranslate("ErrorOfFieldAdd") + BDOSResources.getTranslate("ErrorDescription") + ":" + errMsg + "! " + BDOSResources.getTranslate("Code") + " : " + errCode + "!" + BDOSResources.getTranslate("Table") + ": " + "\"" + tableName + "\", " + BDOSResources.getTranslate("Field") + ": " + "\"" + name + "\"";
                    InsertLogRow(tableName, "", name, errCode, errMsg);
                }
                else
                {
                    InsertLogRow(tableName, "", name, 0, "");
                }
            }
            catch (Exception ex)
            {
                Program.oCompany.GetLastError(out errCode, out errMsg);
                errorText = BDOSResources.getTranslate("ErrorOfFieldAdd") + BDOSResources.getTranslate("ErrorDescription") + ":" + errMsg + "! " + BDOSResources.getTranslate("Code") + " : " + errCode + "!" + BDOSResources.getTranslate("Table") + ": " + "\"" + tableName + "\", " + BDOSResources.getTranslate("Field") + ": " + "\"" + name + "\"" + BDOSResources.getTranslate("OtherInfo") + ": " + ex.Message;
                InsertLogRow(tableName, "", name, errCode, errMsg);
            }

            finally
            {
                Marshal.ReleaseComObject(oUserFieldsMD);
                oUserFieldsMD = null;

                if (validValuesCOM != null)
                {
                    Marshal.ReleaseComObject(validValuesCOM);
                    validValuesCOM = null;
                }

                GC.WaitForPendingFinalizers();
                GC.Collect();
                propertyValue = null;
            }
        }

        public static void InsertLogRow(string tableName, string udoName, string name, int statusCode, string description)
        {
            if (tableName != "BDOSLOGS")
            {
                SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                var dtNow = DateTime.UtcNow;
                int lengthUserName = Math.Min(Environment.UserName.Length, 10);
                string Code = dtNow.ToString("yyyyMMddmmss") + "-" + Environment.UserName.Substring(0, lengthUserName) + "-" + tableName + "-" + name.Trim();

                string query = @"INSERT INTO ""@BDOSLOGS"" (""Code"",""Name"",""U_BDOSExID"",""U_BDOSTbNm"",""U_BDOSFdNm"",
                                                                   ""U_BDOSUDO"",""U_BDOSStts"",""U_BDOSPcUsr"",""U_BDOSB1Usr"",""U_BDOSDt"",""U_BDOSStCd"",""U_BDOSDesc"")" +
                $" VALUES('{Code}','{Code}','{Program.ExecutionDateISO + "-" + Environment.UserName}','{tableName}','{name.Trim()}','{udoName.Trim()}','{(statusCode == 0 ? 'Y' : 'N')}','{Environment.UserName}','{Program.oCompany.UserName}','{dtNow.ToString("yyyy-MM-ddTHH:mm:ss")}',{statusCode},'{description.Replace("'","-").Replace("\"","-")}')";

                oRecordSet.DoQuery(query);

                Marshal.ReleaseComObject(oRecordSet);
                oRecordSet = null;
                GC.Collect();
            }
        }

        public static bool UserDefinedTableExists(string tableName)
        {
            return Program.UserDefinedTablesCurrentCompany.AsEnumerable().Where(x => (string)x["TableName"] == tableName).Any();
        }

        public static bool UserDefinedFieldExists(string tableName, string fieldName)
        {        
            return Program.UserDefinedFieldsCurrentCompany.AsEnumerable().Where(
                        x => ((string)x["TableName"] == "@" + tableName.ToUpperInvariant() && (string)x["FieldName"] == fieldName && ((string)x["TableName"]).Substring(0,1) == "@")
                        || ((string)x["TableName"] == tableName.ToUpperInvariant() && (string)x["FieldName"] == fieldName && ((string)x["TableName"]).Substring(0, 1) != "@")).Any();
        }

        public static DataTable UserDefinedFieldsCurrentCompany()
        {
            
            var oDataTable = new DataTable();
            oDataTable.Columns.Add("TableName");
            oDataTable.Columns.Add("FieldName");

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string query = @"SELECT ""TableID"",""AliasID"",* FROM ""CUFD"" ";                 

            oRecordSet.DoQuery(query);

            while (!oRecordSet.EoF)
            {
                var dtRow = oDataTable.NewRow();
                dtRow["TableName"] = oRecordSet.Fields.Item("TableID").Value;
                dtRow["FieldName"] = oRecordSet.Fields.Item("AliasID").Value;

                oDataTable.Rows.Add(dtRow);
                oRecordSet.MoveNext();
            }

            Marshal.ReleaseComObject(oRecordSet);
            oRecordSet = null;
            GC.Collect();

            return oDataTable;
        }

        public static DataTable UserDefinedTablesCurrentCompany()
        {
            
            var oDataTable = new DataTable();
            oDataTable.Columns.Add("TableName");

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string query = @"SELECT ""TableName"",* FROM ""OUTB"" ";

            oRecordSet.DoQuery(query);

            while (!oRecordSet.EoF)
            {
                var dtRow = oDataTable.NewRow();
                dtRow["TableName"] = oRecordSet.Fields.Item("TableName").Value;

                oDataTable.Rows.Add(dtRow);
                oRecordSet.MoveNext();
            }

            Marshal.ReleaseComObject(oRecordSet);
            oRecordSet = null;
            GC.Collect();

            return oDataTable;
        }

        public static bool addUpdateRecordsUserTable( object tableName, string code, string name, object fieldName, string value, out string errorText)
        {
            errorText = null;
            SAPbobsCOM.UserTable oUserTable = null;

            oUserTable = Program.oCompany.UserTables.Item(tableName);

            oUserTable.Code = code;
            oUserTable.Name = name;
            oUserTable.UserFields.Fields.Item(fieldName).Value = value;

            try
            {
                int returnCode;

                if (oUserTable.GetByKey(code) == false)
                {
                    returnCode = oUserTable.Add();
                }
                else
                {
                    returnCode = oUserTable.Update();
                }

                int errCode;
                string errMsg;

                if (returnCode != 0)
                {
                    Program.oCompany.GetLastError(out errCode, out errMsg);
                    errorText = BDOSResources.getTranslate("ErrorDescription") + " : " + errMsg + "! " + BDOSResources.getTranslate("Code") + " : " + errCode + "!";
                    return false;
                }
            }
            catch
            {
                int errCode;
                string errMsg;

                Program.oCompany.GetLastError(out errCode, out errMsg);
                errorText = BDOSResources.getTranslate("ErrorDescription") + " : " + errMsg + "! " + BDOSResources.getTranslate("Code") + " : " + errCode + "! ";

                return false;
            }

            return true;
        }

        public static bool removeColumnUserTable( object tableName, object fieldName, out string errorText)
        {
            errorText = null;
            SAPbobsCOM.UserFieldsMD oUserFieldsMD = null;

            oUserFieldsMD = ((SAPbobsCOM.UserFieldsMD)(Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)));
            oUserFieldsMD.Name = fieldName.ToString();
            oUserFieldsMD.TableName = tableName.ToString();

            try
            {
                int returnCode;

                returnCode = oUserFieldsMD.Remove();

                int errCode;
                string errMsg;

                if (returnCode != 0)
                {
                    Program.oCompany.GetLastError(out errCode, out errMsg);
                    errorText = BDOSResources.getTranslate("ErrorDescription") + " : " + errMsg + "! " + BDOSResources.getTranslate("Code") + " : " + errCode + "!";
                    return false;
                }
            }
            catch (Exception ex)
            {
                int errCode;
                string errMsg;

                Program.oCompany.GetLastError(out errCode, out errMsg);
                errorText = BDOSResources.getTranslate("ErrorDescription") + " : " + errMsg + "! " + BDOSResources.getTranslate("Code") + " : " + errCode + "! " + BDOSResources.getTranslate("OtherInfo") + " : " + ex.Message;

                return false;
            }

            finally
            {
                Marshal.ReleaseComObject(oUserFieldsMD);
                GC.Collect();
            }

            return true;
        }

        /// <summary>ჩამონათვალის ტიპის ველებში ახალი მნიშვნელობის დამატება</summary>
        /// <param name="Program.oCompany"></param>
        /// <param name="tableID"></param>
        /// <param name="aliasID"></param>
        /// <param name="value"></param>
        /// <param name="description"></param>
        /// <param name="errorText"></param>
        /// <returns></returns>
        public static bool addNewValidValuesUserFieldsMD( string tableID, string aliasID, string value, string description, out string errorText)
        {
            errorText = null;

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            oRecordSet = null;
            GC.WaitForPendingFinalizers();
            GC.Collect();

            SAPbobsCOM.UserFieldsMD oUserFieldsMD = (SAPbobsCOM.UserFieldsMD)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD);
            oUserFieldsMD = null;
            GC.WaitForPendingFinalizers();
            GC.Collect();

            oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRecordSet.DoQuery("SELECT \"FieldID\" FROM \"CUFD\" Where \"AliasID\" = '" + aliasID + "' AND \"TableID\" = '" + tableID + "'");

            oUserFieldsMD = (SAPbobsCOM.UserFieldsMD)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);

            int errCode = 0;
            string errMsg = "";

            try
            {
                if (oRecordSet.RecordCount == 1)
                {
                    //oUserFieldsMD = (SAPbobsCOM.UserFieldsMD)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);

                    if (oUserFieldsMD.GetByKey(tableID, Convert.ToInt32(oRecordSet.Fields.Item("FieldID").Value.ToString())))
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                        oRecordSet = null;
                        GC.WaitForPendingFinalizers();

                        int lineNum = oUserFieldsMD.ValidValues.Count - 1;
                        oUserFieldsMD.ValidValues.Add();
                        oUserFieldsMD.ValidValues.Value = value;
                        oUserFieldsMD.ValidValues.Description = description;

                        if (oUserFieldsMD.Update() != 0)
                        {
                            Program.oCompany.GetLastError(out errCode, out errMsg);
                            errorText = BDOSResources.getTranslate("ErrorOfValueAdd") + BDOSResources.getTranslate("ErrorDescription") + ":" + errMsg + "! " + BDOSResources.getTranslate("Code") + " : " + errCode + "!" + BDOSResources.getTranslate("Table") + ": " + "\"" + tableID + "\", " + BDOSResources.getTranslate("Value") + ": " + "\"" + description + "\"";
                            return false;
                        }
                    }
                }
                else
                    return false;
            }
            catch (Exception ex)
            {
                Program.oCompany.GetLastError(out errCode, out errMsg);
                errorText = BDOSResources.getTranslate("ErrorOfValueAdd") + BDOSResources.getTranslate("ErrorDescription") + ":" + errMsg + "! " + BDOSResources.getTranslate("Code") + " : " + errCode + "!" + BDOSResources.getTranslate("Table") + ": " + "\"" + tableID + "\", " + BDOSResources.getTranslate("Value") + ": " + "\"" + description + "\"" + BDOSResources.getTranslate("OtherInfo") + ": " + ex.Message;
                return false;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD);
                oUserFieldsMD = null;
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
            return true;
        }

        private static bool userFormExist( string code)
        {
            SAPbobsCOM.UserObjectsMD oUserObjectsMD = null;
            bool boolIdent = false;
            oUserObjectsMD = ((SAPbobsCOM.UserObjectsMD)(Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)));
            boolIdent = oUserObjectsMD.GetByKey(code);
            Marshal.ReleaseComObject(oUserObjectsMD);
            GC.Collect();

            return (boolIdent);
        }

        public static int registerUDO( string code, Dictionary<string, object> formProperties, out string errorText)
        {
            errorText = null;

            SAPbobsCOM.UserObjectsMD oUserObjectsMD = null;
            oUserObjectsMD = (SAPbobsCOM.UserObjectsMD)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);
            bool udoFormExist = false;

            if (userFormExist( code) == true)
            {
                udoFormExist = true;
                oUserObjectsMD.GetByKey(code);
            }
            else
            {
                oUserObjectsMD.Code = code;
            }

            object propertyValue = null;
            object keyValue = null;

            if (formProperties.TryGetValue("CanApprove", out propertyValue) == true)
            {
                oUserObjectsMD.CanApprove = (SAPbobsCOM.BoYesNoEnum)propertyValue;
            }
            if (formProperties.TryGetValue("CanArchive", out propertyValue) == true)
            {
                oUserObjectsMD.CanArchive = (SAPbobsCOM.BoYesNoEnum)propertyValue;
            }
            if (formProperties.TryGetValue("CanCancel", out propertyValue) == true)
            {
                oUserObjectsMD.CanCancel = (SAPbobsCOM.BoYesNoEnum)propertyValue;
            }
            if (formProperties.TryGetValue("CanClose", out propertyValue) == true)
            {
                oUserObjectsMD.CanClose = (SAPbobsCOM.BoYesNoEnum)propertyValue;
            }
            if (formProperties.TryGetValue("CanCreateDefaultForm", out propertyValue) == true)
            {
                oUserObjectsMD.CanCreateDefaultForm = (SAPbobsCOM.BoYesNoEnum)propertyValue;
            }
            if (formProperties.TryGetValue("CanDelete", out propertyValue) == true)
            {
                oUserObjectsMD.CanDelete = (SAPbobsCOM.BoYesNoEnum)propertyValue;
            }
            if (formProperties.TryGetValue("CanFind", out propertyValue) == true)
            {
                oUserObjectsMD.CanFind = (SAPbobsCOM.BoYesNoEnum)propertyValue;
            }
            if (formProperties.TryGetValue("CanLog", out propertyValue) == true)
            {
                oUserObjectsMD.CanLog = (SAPbobsCOM.BoYesNoEnum)propertyValue;
            }
            if (formProperties.TryGetValue("CanYearTransfer", out propertyValue) == true)
            {
                oUserObjectsMD.CanYearTransfer = (SAPbobsCOM.BoYesNoEnum)propertyValue;
            }
            if (formProperties.TryGetValue("ChildTables", out propertyValue) == true)
            {
                List<Dictionary<string, object>> listChildTables = (List<Dictionary<string, object>>)propertyValue;

                for (int i = 0; i < listChildTables.Count(); i++)
                {
                    bool alreadyCont = false;
                    string strPropertykey = "";
                    if (listChildTables[i].TryGetValue("TableName", out keyValue) == true)
                    {
                        strPropertykey = keyValue.ToString();
                    }

                    if (strPropertykey == "") continue;

                    for (int n = 0; n < oUserObjectsMD.ChildTables.Count; n++)
                    {
                        oUserObjectsMD.ChildTables.SetCurrentLine(n);
                        if (oUserObjectsMD.ChildTables.TableName.Equals(strPropertykey))
                        {
                            alreadyCont = true;
                            break;
                    }
                    }

                    if (alreadyCont == true) continue;

                    oUserObjectsMD.ChildTables.Add();
                    oUserObjectsMD.ChildTables.SetCurrentLine(oUserObjectsMD.ChildTables.Count - 1);
                    oUserObjectsMD.ChildTables.TableName = strPropertykey;
                    //if (listChildTables[i].TryGetValue("TableName", out keyValue) == true)
                    //{
                    //    oUserObjectsMD.ChildTables.TableName = keyValue.ToString();
                    //}
                    if (listChildTables[i].TryGetValue("ObjectName", out keyValue) == true)
                    {
                        oUserObjectsMD.ChildTables.ObjectName = keyValue.ToString();
                    }
                    if (listChildTables[i].TryGetValue("LogTableName", out keyValue) == true)
                    {
                        oUserObjectsMD.ChildTables.LogTableName = keyValue.ToString();
                    }
                    //oUserObjectsMD.ChildTables.Add();
                }
                listChildTables = null;
            }
            if (formProperties.TryGetValue("EnableEnhancedForm", out propertyValue) == true)
            {
                oUserObjectsMD.EnableEnhancedForm = (SAPbobsCOM.BoYesNoEnum)propertyValue;
            }

            if (formProperties.TryGetValue("EnhancedFormColumns", out propertyValue) == true) //UDO4
            {
                List<Dictionary<string, object>> listEnhancedFormColumns = (List<Dictionary<string, object>>)propertyValue;

                for (int i = 0; i < listEnhancedFormColumns.Count(); i++)
                {
                    bool alreadyCont = false;
                    string strPropertykey = "";
                    if (listEnhancedFormColumns[i].TryGetValue("ColumnAlias", out keyValue) == true)
                    {
                        strPropertykey = keyValue.ToString();
                    }

                    if (strPropertykey == "") continue;

                    for (int n = 0; n < oUserObjectsMD.EnhancedFormColumns.Count; n++)
                    {
                        oUserObjectsMD.EnhancedFormColumns.SetCurrentLine(n);
                        if (oUserObjectsMD.EnhancedFormColumns.ColumnAlias.Equals(strPropertykey))
                    {
                            alreadyCont = true;
                            break;
                        }
                    }

                    if (alreadyCont == true) continue;

                    oUserObjectsMD.EnhancedFormColumns.Add();
                    oUserObjectsMD.EnhancedFormColumns.SetCurrentLine(oUserObjectsMD.EnhancedFormColumns.Count - 1);
                    oUserObjectsMD.EnhancedFormColumns.ColumnAlias = strPropertykey;

                    if (listEnhancedFormColumns[i].TryGetValue("ChildNumber", out keyValue) == true)
                    {
                        oUserObjectsMD.EnhancedFormColumns.ChildNumber = Convert.ToInt32(keyValue);
                    }
                    //if (listEnhancedFormColumns[i].TryGetValue("ColumnAlias", out keyValue) == true)
                    //{
                    //    oUserObjectsMD.EnhancedFormColumns.ColumnAlias = keyValue.ToString();
                    //}
                    if (listEnhancedFormColumns[i].TryGetValue("ColumnDescription", out keyValue) == true)
                    {
                        oUserObjectsMD.EnhancedFormColumns.ColumnDescription = keyValue.ToString();
                    }
                    if (listEnhancedFormColumns[i].TryGetValue("ColumnIsUsed", out keyValue) == true)
                    {
                        oUserObjectsMD.EnhancedFormColumns.ColumnIsUsed = (SAPbobsCOM.BoYesNoEnum)keyValue;
                    }
                    if (listEnhancedFormColumns[i].TryGetValue("ColumnNumber", out keyValue) == true)
                    {
                        oUserObjectsMD.EnhancedFormColumns.ColumnNumber = Convert.ToInt32(keyValue);
                    }
                    if (listEnhancedFormColumns[i].TryGetValue("Editable", out keyValue) == true)
                    {
                        oUserObjectsMD.EnhancedFormColumns.Editable = (SAPbobsCOM.BoYesNoEnum)keyValue;
                    }
                    //oUserObjectsMD.EnhancedFormColumns.Add();
                }
                listEnhancedFormColumns = null;
            }

            if (formProperties.TryGetValue("ExtensionName", out propertyValue) == true)
            {
                oUserObjectsMD.ExtensionName = propertyValue.ToString();
            }
            if (formProperties.TryGetValue("FatherMenuID", out propertyValue) == true)
            {
                oUserObjectsMD.FatherMenuID = Convert.ToInt32(propertyValue);
            }

            if (formProperties.TryGetValue("FindColumns", out propertyValue) == true) //UDO2
            {
                List<Dictionary<string, object>> listFindColumns = (List<Dictionary<string, object>>)propertyValue;

                for (int i = 0; i < listFindColumns.Count(); i++)
                {
                    bool alreadyCont = false;
                    string strPropertykey = "";
                    if (listFindColumns[i].TryGetValue("ColumnAlias", out keyValue) == true)
                    {
                        strPropertykey = keyValue.ToString();
                    }

                    if (strPropertykey == "") continue;

                    for (int n = 0; n < oUserObjectsMD.FindColumns.Count; n++)
                    {
                        oUserObjectsMD.FindColumns.SetCurrentLine(n);
                        if (oUserObjectsMD.FindColumns.ColumnAlias.Equals(strPropertykey))
                        {
                            alreadyCont = true;
                            break;
                    }
                    }

                    if (alreadyCont == true) continue;

                    oUserObjectsMD.FindColumns.Add();
                    oUserObjectsMD.FindColumns.SetCurrentLine(oUserObjectsMD.FindColumns.Count - 1);
                    oUserObjectsMD.FindColumns.ColumnAlias = strPropertykey;

                    //if (listFindColumns[i].TryGetValue("ColumnAlias", out keyValue) == true)
                    //{
                    //    oUserObjectsMD.FindColumns.ColumnAlias = keyValue.ToString();
                    //}
                    if (listFindColumns[i].TryGetValue("ColumnDescription", out keyValue) == true)
                    {
                        oUserObjectsMD.FindColumns.ColumnDescription = keyValue.ToString();
                    }
                    //oUserObjectsMD.FindColumns.Add();
                }
                listFindColumns = null;
            }

            if (formProperties.TryGetValue("FormColumns", out propertyValue) == true) //UDO3
            {
                List<Dictionary<string, object>> listFormColumns = (List<Dictionary<string, object>>)propertyValue;

                for (int i = 0; i < listFormColumns.Count(); i++)
                {
                    bool alreadyCont = false;
                    string strPropertykey = "";
                    if (listFormColumns[i].TryGetValue("FormColumnAlias", out keyValue) == true)
                    {
                        strPropertykey = keyValue.ToString();
                    }

                    if (strPropertykey == "") continue;

                    for (int n = 0; n < oUserObjectsMD.FormColumns.Count; n++)
                    {
                        oUserObjectsMD.FormColumns.SetCurrentLine(n);
                        if (oUserObjectsMD.FormColumns.FormColumnAlias.Equals(strPropertykey))
                    {
                            alreadyCont = true;
                            break;
                        }
                    }

                    if (alreadyCont == true) continue;

                    oUserObjectsMD.FormColumns.Add();
                    oUserObjectsMD.FormColumns.SetCurrentLine(oUserObjectsMD.FormColumns.Count - 1);
                    oUserObjectsMD.FormColumns.FormColumnAlias = strPropertykey;

                    if (listFormColumns[i].TryGetValue("Editable", out keyValue) == true)
                    {
                        oUserObjectsMD.FormColumns.Editable = (SAPbobsCOM.BoYesNoEnum)keyValue;
                    }
                    //if (listFormColumns[i].TryGetValue("FormColumnAlias", out keyValue) == true)
                    //{
                    //    oUserObjectsMD.FormColumns.FormColumnAlias = keyValue.ToString();
                    //}
                    if (listFormColumns[i].TryGetValue("FormColumnDescription", out keyValue) == true)
                    {
                        oUserObjectsMD.FormColumns.FormColumnDescription = keyValue.ToString();
                    }
                    if (listFormColumns[i].TryGetValue("SonNumber", out keyValue) == true)
                    {
                        oUserObjectsMD.FormColumns.SonNumber = Convert.ToInt32(keyValue);
                    }
                    //oUserObjectsMD.FormColumns.Add();
                }
                listFormColumns = null;
            }

            if (formProperties.TryGetValue("FormSRF", out propertyValue) == true)
            {
                oUserObjectsMD.FormSRF = propertyValue.ToString();
            }
            if (formProperties.TryGetValue("LogTableName", out propertyValue) == true)
            {
                oUserObjectsMD.LogTableName = propertyValue.ToString();
            }
            if (formProperties.TryGetValue("ManageSeries", out propertyValue) == true)
            {
                oUserObjectsMD.ManageSeries = (SAPbobsCOM.BoYesNoEnum)propertyValue;
            }
            if (formProperties.TryGetValue("MenuCaption", out propertyValue) == true)
            {
                oUserObjectsMD.MenuCaption = propertyValue.ToString();
            }
            if (formProperties.TryGetValue("MenuItem", out propertyValue) == true)
            {
                oUserObjectsMD.MenuItem = (SAPbobsCOM.BoYesNoEnum)propertyValue;
            }
            if (formProperties.TryGetValue("MenuUID", out propertyValue) == true)
            {
                oUserObjectsMD.MenuUID = propertyValue.ToString();
            }
            if (formProperties.TryGetValue("Name", out propertyValue) == true)
            {
                oUserObjectsMD.Name = propertyValue.ToString();
            }
            if (formProperties.TryGetValue("ObjectType", out propertyValue) == true)
            {
                oUserObjectsMD.ObjectType = (SAPbobsCOM.BoUDOObjType)propertyValue;
            }
            if (formProperties.TryGetValue("OverwriteDllfile", out propertyValue) == true)
            {
                oUserObjectsMD.OverwriteDllfile = (SAPbobsCOM.BoYesNoEnum)propertyValue;
            }
            if (formProperties.TryGetValue("Position", out propertyValue) == true)
            {
                oUserObjectsMD.Position = Convert.ToInt32(propertyValue);
            }
            if (formProperties.TryGetValue("RebuildEnhancedForm", out propertyValue) == true)
            {
                oUserObjectsMD.RebuildEnhancedForm = (SAPbobsCOM.BoYesNoEnum)propertyValue;
            }
            if (formProperties.TryGetValue("TableName", out propertyValue) == true)
            {
                oUserObjectsMD.TableName = propertyValue.ToString();
            }
            if (formProperties.TryGetValue("TemplateID", out propertyValue) == true)
            {
                oUserObjectsMD.TemplateID = propertyValue.ToString();
            }
            if (formProperties.TryGetValue("UseUniqueFormType", out propertyValue) == true)
            {
                oUserObjectsMD.UseUniqueFormType = (SAPbobsCOM.BoYesNoEnum)propertyValue;
            }

            int errCode;
            string errMsg;
            string updateAddTxt = udoFormExist == false ? BDOSResources.getTranslate("ErrorOfFormAdd") : BDOSResources.getTranslate("ErrorOfFormUpdate");

            try
            {
                int returnCode;

                if (udoFormExist == false)
                {
                    returnCode = oUserObjectsMD.Add();
                }
                else
                {
                    returnCode = oUserObjectsMD.Update();
                }

                if (returnCode != 0)
                {
                    Program.oCompany.GetLastError(out errCode, out errMsg);
                    errorText = updateAddTxt + BDOSResources.getTranslate("ErrorDescription") + ":" + errMsg + "! " + BDOSResources.getTranslate("Code") + " : " + errCode + "!" + BDOSResources.getTranslate("Form") + ": " + "\"" + code + "\"";

                    InsertLogRow(oUserObjectsMD.TableName, code, "", errCode, errMsg);

                    return returnCode;
                }
                else
                {
                    InsertLogRow(oUserObjectsMD.TableName, code, "", 0, "");
                }
            }
            catch (Exception ex)
            {
                Program.oCompany.GetLastError(out errCode, out errMsg);
                errorText = updateAddTxt + BDOSResources.getTranslate("ErrorDescription") + ":" + errMsg + "! " + BDOSResources.getTranslate("Code") + " : " + errCode + "!" + BDOSResources.getTranslate("Form") + ": " + "\"" + code + "\"" + BDOSResources.getTranslate("OtherInfo") + ": " + ex.Message;

                InsertLogRow(oUserObjectsMD.TableName, code, "", errCode, errMsg);

                return -1;
            }
            finally
            {
                Marshal.ReleaseComObject(oUserObjectsMD);
                GC.Collect();
            }

            return 0;
        }

        public static void DeleteUDF( string tableID, int fieldID, out string errorText)
        {
            errorText = null;

            SAPbobsCOM.UserFieldsMD sboField = (SAPbobsCOM.UserFieldsMD)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);

            try
            {
                if (sboField.GetByKey(tableID, fieldID))
                {
                    if (sboField.Remove() != 0)
                    {
                        errorText = Program.oCompany.GetLastErrorDescription();
                    }
                }
            }
            finally
            {
                Marshal.ReleaseComObject(sboField);
                sboField = null;
                GC.Collect();
            }
        }

        public static dynamic GetUDOFieldValueByParam(string UDO, string propertyFilterParam, dynamic valueFilterParam, string property)
        {     

            if (String.IsNullOrEmpty(propertyFilterParam) || String.IsNullOrEmpty(valueFilterParam))
            {
                return null;
            }
            else
            {
            var oGeneralService = Program.oCompany.GetCompanyService().GetGeneralService(UDO);
            var oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
            oGeneralParams.SetProperty(propertyFilterParam, valueFilterParam);
            SAPbobsCOM.GeneralData oGeneralData = oGeneralService.GetByParams(oGeneralParams);

            return oGeneralData.GetProperty(property);
        }
        }

        public static int GetFieldID( string sTableID, string sAliasID)
        {
            int iRetVal = 0;
            SAPbobsCOM.Recordset sboRec = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                sboRec.DoQuery(@"select ""FieldID"" from ""CUFD"" where ""TableID"" = '" + sTableID + @"' and ""AliasID"" = '" + sAliasID + "'");
                if (!sboRec.EoF) iRetVal = Convert.ToInt32(sboRec.Fields.Item("FieldID").Value.ToString());
            }
            finally
            {
                Marshal.ReleaseComObject(sboRec);
                sboRec = null;
                GC.Collect();
            }
            return iRetVal;
        }

        public static void AddUserKey( string tableName, string keyName, List<string> oColumnAlias, out string errorText)
        {
            errorText = null;

            SAPbobsCOM.UserKeysMD oUserKeysMD;
            oUserKeysMD = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserKeys);
            try
            {
                oUserKeysMD.TableName = tableName;
                oUserKeysMD.KeyName = keyName;

                for (int i = 0; i < oColumnAlias.Count; i++)
                {
                    if (i != 0)
                    {
                        oUserKeysMD.Elements.Add();
                    }
                    oUserKeysMD.Elements.ColumnAlias = oColumnAlias[i];
                }
   
                oUserKeysMD.Unique = SAPbobsCOM.BoYesNoEnum.tYES;
                
                int returnCode = oUserKeysMD.Add();
                if (returnCode != 0)
                {
                    int errCode;
                    string errMsg;

                    Program.oCompany.GetLastError(out errCode, out errMsg);
                    errorText = BDOSResources.getTranslate("ErrorDescription") + " : " + errMsg + "! " + BDOSResources.getTranslate("Code") + " : " + errCode;
                    return;
                }
            }
            catch(Exception ex)
            {
                errorText = ex.Message;
            }
            finally
            {
                Marshal.ReleaseComObject(oUserKeysMD);
                oUserKeysMD = null;
                GC.Collect();
            }
        }
    }
}
