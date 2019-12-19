using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace BDO_Localisation_AddOn
{
    public static class BDOSWarehouseAddresses
    {
        public static void CreateMasterDataUDO(out string errorText)
        {
            errorText = null;

            string tableName = "BDOSWRHADR";
            string description = "Warehouse Addresses";

            int result = UDO.addUserTable(tableName, description, SAPbobsCOM.BoUTBTableType.bott_NoObjectAutoIncrement, out errorText);

            if (result != 0)
            {
                return;
            }

            Dictionary<string, object> fieldskeysMap;

            //Warehouse Code
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "WhsCode");
            fieldskeysMap.Add("TableName", tableName);
            fieldskeysMap.Add("Description", "Warehouse Code");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);
            fieldskeysMap.Add("Mandatory", SAPbobsCOM.BoYesNoEnum.tYES);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            //End Address
            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "EndAddress");
            fieldskeysMap.Add("TableName", tableName);
            fieldskeysMap.Add("Description", "Warehouse End Address");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 200);
            fieldskeysMap.Add("Mandatory", SAPbobsCOM.BoYesNoEnum.tYES);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            //Make EndAddress unique
            SAPbobsCOM.UserKeysMD oUserKeysMD = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserKeys);

            oUserKeysMD.TableName = tableName;
            oUserKeysMD.KeyName = "EndAddress";
            oUserKeysMD.Elements.ColumnAlias = "EndAddress";
            oUserKeysMD.Unique = SAPbobsCOM.BoYesNoEnum.tYES;

            int returnCode = oUserKeysMD.Add();

            Marshal.ReleaseComObject(oUserKeysMD);

            if (returnCode != 0)
            {
                int errCode;
                string errMsg;

                Program.oCompany.GetLastError(out errCode, out errMsg);
                errorText = BDOSResources.getTranslate("ErrorDescription") + " : " + errMsg + "! " + BDOSResources.getTranslate("Code") + " : " + errCode;

            }
        }

        public static void UiApp_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD && pVal.BeforeAction)
                {
                    Program.FORM_LOAD_FOR_ACTIVATE = true;
                    ChangeFormItems(oForm);
                    oForm.Title = BDOSResources.getTranslate("WarehouseAddresses");
                }

                else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                {
                    SAPbouiCOM.IChooseFromListEvent oCFLEvento = (SAPbouiCOM.IChooseFromListEvent)pVal;
                    chooseFromList(oForm, pVal, oCFLEvento);
                }
            }
        }

        public static void ChangeFormItems(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("3").Specific));
            int lastRow = oMatrix.RowCount;

            SAPbouiCOM.Columns oColumns = oMatrix.Columns;


            SAPbouiCOM.Column oColumn = oColumns.Item("Code");
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Code");

            FormsB1.addChooseFromList(oForm, false, "64", "WarehouseCFL");
            oColumn = oColumns.Item("U_WhsCode");
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("Warehouse");
            oColumn.ChooseFromListUID = "WarehouseCFL";
            oColumn.ChooseFromListAlias = "WhsCode";
            oColumn.Cells.Item(lastRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);

            oColumn = oColumns.Item("Name");
            oColumn.Visible = false;

            oColumn = oColumns.Item("U_EndAddress");
            oColumn.TitleObject.Caption = BDOSResources.getTranslate("EndAddress");

        }

        public static void chooseFromList(SAPbouiCOM.Form oForm, SAPbouiCOM.ItemEvent pVal, SAPbouiCOM.IChooseFromListEvent oCFLEvento)
        {
            try
            {
                oForm.Freeze(true);
                if (!pVal.BeforeAction)
                {
                    SAPbouiCOM.DataTable oDataTable = oCFLEvento.SelectedObjects;

                    SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("3").Specific));

                    if (oDataTable != null)
                    {
                        if (oCFLEvento.ChooseFromListUID == "WarehouseCFL")
                        {
                            string whsCode = oDataTable.GetValue("WhsCode", 0);

                            LanguageUtils.IgnoreErrors<string>(() => oMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value = whsCode);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                oForm.Freeze(false);
            }

        }

        public static string GetWhsByAddress(string address)
        {
            string whsCode = null;

            SAPbobsCOM.Recordset oRecordSetWhsAddress = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            StringBuilder queryWhsAddress = new StringBuilder();
            queryWhsAddress.Append("select \"U_WhsCode\" \n");
            queryWhsAddress.Append("from \"@BDOSWRHADR\" \n");
            queryWhsAddress.Append("where \"U_EndAddress\" = '" + address + "'");

            oRecordSetWhsAddress.DoQuery(queryWhsAddress.ToString());

            if (!oRecordSetWhsAddress.EoF)
            {
                whsCode = oRecordSetWhsAddress.Fields.Item("U_WhsCode").Value;
            }

            Marshal.ReleaseComObject(oRecordSetWhsAddress);
            return whsCode;
        }

        public static void AddWhsByAddress(string address, string whsCode)
        {
            string whsCodeInTable = GetWhsByAddress(address);

            SAPbobsCOM.Recordset oRecordSetWhsAddress = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            StringBuilder queryWhsAddress = new StringBuilder();

            queryWhsAddress.Append("select max(\"Code\") as \"Code\" \n");
            queryWhsAddress.Append("from \"@BDOSWRHADR\"");
            oRecordSetWhsAddress.DoQuery(queryWhsAddress.ToString());
            int code = oRecordSetWhsAddress.Fields.Item("Code").Value+1;

            try
            {
                if (string.IsNullOrEmpty(whsCodeInTable))
                {
                    queryWhsAddress.Clear();
                    queryWhsAddress.Append("INSERT INTO \"@BDOSWRHADR\" (\"Code\",\"U_WhsCode\", \"U_EndAddress\") \n");
                    queryWhsAddress.Append("VALUES ('" + code + "','" + whsCode + "', '" + address + "')");

                    oRecordSetWhsAddress.DoQuery(queryWhsAddress.ToString());
                }
                else
                {
                    if (whsCode != whsCodeInTable)
                    {
                        queryWhsAddress.Clear();
                        queryWhsAddress.Append("UPDATE \"@BDOSWRHADR\" \n");
                        queryWhsAddress.Append("SET \"U_WhsCode\" = '" + whsCode + "' \n");
                        queryWhsAddress.Append("WHERE \"U_EndAddress\" = '" + address + "'");
                    }
                }
            }

            catch (Exception ex)
            {
                Program.uiApp.SetStatusBarMessage(ex.Message);
            }
            finally
            {
                Marshal.ReleaseComObject(oRecordSetWhsAddress);
            }
        }
    }
}
