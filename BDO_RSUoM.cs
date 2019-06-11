using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace BDO_Localisation_AddOn
{
    static partial class BDO_RSUoM
    {
        public static void downloadUnits( SAPbouiCOM.Form oForm, out string errorText)
        {

            errorText = null;

            Dictionary<string, string> rsSettings = CompanyDetails.getRSSettings( out errorText);
            if (errorText != null)
            {
                return;
            }

            string su = rsSettings["SU"];
            string sp = rsSettings["SP"];
            WayBill oWayBill = new WayBill(su, sp, rsSettings["ProtocolType"]);

            bool chek_service_user = oWayBill.chek_service_user(su, sp, out errorText);
            if (chek_service_user == false)
            {
                errorText = BDOSResources.getTranslate("ServiceUserPasswordNotCorrect");
                return;
            }

            Dictionary<string,string> RSUnits =  oWayBill.get_waybill_units(out errorText);

            string unitName;
            string unitID;

            oForm.Freeze(true);

            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("UoMMatrix").Specific;
            for (int i = 1; i <= oMatrix.RowCount; i++)
            {
                unitName = oMatrix.GetCellSpecific("UomName", i).Value;           
                KeyValuePair<string, string> temp = RSUnits.Where(x => x.Value.Equals(unitName)).FirstOrDefault(); //Contains
                unitID = temp.Key;
                if (unitID != null)
                {
                    SAPbouiCOM.EditText oEditText = oMatrix.Columns.Item("BDO_RSCode").Cells.Item(i).Specific;
                    oEditText.Value = unitID;
                }
                else
                {
                    SAPbouiCOM.EditText oEditText = oMatrix.Columns.Item("BDO_RSCode").Cells.Item(i).Specific;
                    oEditText.Value = "99";
                }
            }

            oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
            oForm.Freeze(false);
          
        }

        public static SAPbobsCOM.Recordset getUomByRSCode( string ItemCode, string RSCode, out string errorText)
        {
            errorText = null;
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                string query =
                        "SELECT" +
                        "\"OITM\".\"UgpEntry\",\"OUOM\".\"UomName\",\"UGP1\".\"UomEntry\",\"@BDO_RSUOM\".\"U_RSCode\",\"OUOM\".\"UomCode\"" +
                        "FROM \"OITM\"" +
                        "INNER JOIN \"UGP1\" ON \"UGP1\".\"UgpEntry\" = \"OITM\".\"UgpEntry\"" +
                        "INNER JOIN \"@BDO_RSUOM\" ON \"UGP1\".\"UomEntry\" = \"@BDO_RSUOM\".\"U_UomEntry\"" +
                        "INNER JOIN \"OUOM\" ON \"UGP1\".\"UomEntry\" = \"OUOM\".\"UomEntry\"" +
                       " WHERE (\"ItemCode\" = N'" + ItemCode + "') AND (\"@BDO_RSUOM\".\"U_RSCode\" = '" + RSCode + "')";
                oRecordSet.DoQuery(query);

                if (!oRecordSet.EoF)
                {
                    return oRecordSet;
                }
            }
            catch (Exception ex)
            {
                int errCode;
                string errMsg;

                Program.oCompany.GetLastError(out errCode, out errMsg);
                errorText = BDOSResources.getTranslate("ErrorDescription")+" : " + errMsg + "! "+BDOSResources.getTranslate("Code") +" : " + errCode + "! "+ BDOSResources.getTranslate("OtherInfo")+" : " + ex.Message;
                return null;
            }
            finally
            {
                //Marshal.FinalReleaseComObject(oRecordSet);
                GC.Collect();
            }
            return null;
        }

        public static void createUserFields( out string errorText)
        {
            Dictionary<string, object> fieldskeysMap;

            string tableName = "BDO_RSUom";
            string description = "UomRSCodes";
            int result = UDO.addUserTable( tableName, description, SAPbobsCOM.BoUTBTableType.bott_NoObject, out errorText);

            if (result != 0)
            {
                return;
            }

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "UomEntry");
            fieldskeysMap.Add("TableName", "BDO_RSUom");
            fieldskeysMap.Add("Description", "Uom Entry");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Numeric);
            fieldskeysMap.Add("EditSize", 11);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "RSCode");
            fieldskeysMap.Add("TableName", "BDO_RSUom");
            fieldskeysMap.Add("Description", "Uom RsCode");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            GC.Collect();
        }

        public static void updateCodes(  SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            oForm.Freeze(true);
            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("UoMMatrix").Specific));
            try
            {
                oMatrix.FlushToDataSource();
            }
            catch
            {
                int errCode;
                string errMsg;

                Program.oCompany.GetLastError(out errCode, out errMsg);
                errorText = BDOSResources.getTranslate("ErrorDescription")+" : " + errMsg + "! "+BDOSResources.getTranslate("Code") +" : " + errCode + "!";
                return;
            }

            SAPbouiCOM.DataTable oDataTableSource = oForm.DataSources.DataTables.Item("UomDataTable");

            bool successUpdating = true;

            for (int i = 0; i < oDataTableSource.Rows.Count; i++)
            {
                int UomEntry = Convert.ToInt32(oDataTableSource.GetValue(1, i));
                string UomName = oDataTableSource.GetValue(2, i).Trim();
                string UomCode = oDataTableSource.GetValue(0, i).Trim();
                string RSCode = oDataTableSource.GetValue(3, i).Trim();

                try
                {
                    SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    string query = @"SELECT * FROM ""@BDO_RSUOM"" WHERE ""U_UomEntry"" = '" + UomEntry.ToString() + "'";
                    oRecordSet.DoQuery(query);

                    SAPbobsCOM.Recordset oRecordSetUpdate = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    if (!oRecordSet.EoF)
                    {
                        string updatequery = @"UPDATE ""@BDO_RSUOM""
                                                SET ""U_RSCode"" = '" + RSCode + @"' WHERE ""U_UomEntry"" = " + UomEntry.ToString();

                        oRecordSetUpdate.DoQuery(updatequery);
                    }
                    else
                    {
                        string insertquery = @"INSERT INTO ""@BDO_RSUOM""
                            (""Code"",""Name"",""U_UomEntry"",""U_RSCode"") VALUES ("
                            + UomEntry.ToString() + "," + UomEntry.ToString() + "," + UomEntry.ToString() + ",'" + RSCode + "') ";

                        oRecordSetUpdate.DoQuery(insertquery);
                    }
                }
                catch
                {
                    successUpdating = false;

                    int errCode;
                    string errMsg;

                    Program.oCompany.GetLastError(out errCode, out errMsg);
                    errorText = BDOSResources.getTranslate("ErrorDescription")+" : " + errMsg + "! "+BDOSResources.getTranslate("Code") +" : " + errCode + "!";
                    CommonFunctions.EndTransaction( SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }
                finally
                {
                    CommonFunctions.EndTransaction( SAPbobsCOM.BoWfTransOpt.wf_Commit);

                    if (errorText != null)
                    {
                        Program.uiApp.StatusBar.SetText(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    }
                }
            }

            if (successUpdating)
            {
                Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("UoMRsCodesUpdateSuccess") + " ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }

            oMatrix.LoadFromDataSource();
            oMatrix.AutoResizeColumns();
            oForm.Freeze(false);
        }

        public static void createForm(  out string errorText)
        {
            errorText = null;
            Dictionary<string, object> formItems;
            string itemName;
            SAPbouiCOM.Columns oColumns;
            SAPbouiCOM.Column oColumn;

            //ფორმის აუცილებელი თვისებები
            Dictionary<string, object> formProperties = new Dictionary<string, object>();
            formProperties.Add("UniqueID", "BDO_RSUoMForm");
            formProperties.Add("BorderStyle", SAPbouiCOM.BoFormBorderStyle.fbs_Sizable);
            formProperties.Add("Title", BDOSResources.getTranslate("UomRsCodes"));
            formProperties.Add("Left", 558);
            formProperties.Add("ClientWidth", 400);
            formProperties.Add("Top", 335);
            formProperties.Add("ClientHeight", 250);

            SAPbouiCOM.Form oForm;
            bool newForm;
            bool formExist = FormsB1.createForm( formProperties, out oForm, out newForm, out errorText);

            if (formExist == true)
            {
                if (newForm)
                {
                    itemName = "1";
                    formItems = new Dictionary<string, object>();
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", 5);
                    formItems.Add("Width", 80);
                    formItems.Add("Top", 220);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    //formItems.Add("Caption", "RS კოდების მინიჭება");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    itemName = "DwnldUnts";
                    formItems = new Dictionary<string, object>();
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", 5 + 80 + 2);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", 220);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("RSDownload"));

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    oForm.Freeze(true);

                    SAPbouiCOM.DataTable oDataTable;
                    if (oForm.DataSources.DataTables.Count == 1)
                    {
                        oDataTable = oForm.DataSources.DataTables.Item("UomDataTable");
                    }
                    else
                    {
                        oDataTable = oForm.DataSources.DataTables.Add("UomDataTable");
                    }

                    string queryStr = "SELECT " +
                    "\"OUOM\".\"UomCode\" as UomCode," +
                    "\"OUOM\".\"UomEntry\" as UomEntry," +
                    "\"OUOM\".\"UomName\" as UomName," +
                    "\"@BDO_RSUOM\".\"U_RSCode\" as RSCode" +
                    " FROM  \"OUOM\"" +
                    "LEFT JOIN \"@BDO_RSUOM\" ON \"OUOM\".\"UomEntry\" = \"@BDO_RSUOM\".\"U_UomEntry\"";

                    oDataTable.ExecuteQuery(queryStr);

                    //ზედნადებების ცხრილი
                    itemName = "UoMMatrix";
                    formItems = new Dictionary<string, object>();
                    formItems.Add("isDataSource", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
                    //formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_GRID);
                    formItems.Add("Left", 5);
                    formItems.Add("Width", 400);
                    formItems.Add("Top", 10);
                    formItems.Add("Height", 200);
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("UoMMatrix").Specific;
                    oColumns = oMatrix.Columns;

                    oColumn = oColumns.Add("UomEntry", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("UomEntry");
                    oColumn.Width = 40;
                    oColumn.Editable = false;
                    oColumn.Visible = true;
                    oColumn.DataBind.Bind("UomDataTable", "UomEntry");

                    oColumn = oColumns.Add("UomCode", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("UomCode");
                    oColumn.Width = 40;
                    oColumn.Editable = false;
                    oColumn.Visible = true;
                    oColumn.DataBind.Bind("UomDataTable", "UomCode");

                    oColumn = oColumns.Add("UomName", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("UomName");
                    oColumn.Width = 40;
                    oColumn.Editable = false;
                    oColumn.Visible = true;
                    oColumn.DataBind.Bind("UomDataTable", "UomName");

                    oColumn = oColumns.Add("BDO_RSCode", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("UomRsCode");
                    oColumn.Width = 40;
                    oColumn.Editable = true;
                    oColumn.Visible = true;
                    oColumn.DataBind.Bind("UomDataTable", "RSCode");

                    oMatrix.Clear();
                    oMatrix.LoadFromDataSource();
                    oMatrix.AutoResizeColumns();
                    oForm.Freeze(false);
                }
                oForm.Visible = true;
                oForm.Select();
            }

            GC.Collect();
        }

        public static void uiApp_ItemEvent(  string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED & pVal.ItemUID == "1")
                {
                    SAPbouiCOM.Button OKButton = (SAPbouiCOM.Button)oForm.Items.Item("1").Specific;

                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                    {
                        BDO_RSUoM.updateCodes( oForm, out errorText);
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                        BubbleEvent = !BubbleEvent;
                    }
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED & pVal.ItemUID == "DwnldUnts")
                {
                    BDO_RSUoM.downloadUnits( oForm, out errorText);
                }
            }
        }

        public static void uiApp_ItemEvent_Setup(  string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED & pVal.ItemUID == "10000001" & pVal.BeforeAction == true)
                {
                    SAPbouiCOM.Button OKButton = (SAPbouiCOM.Button)oForm.Items.Item("10000001").Specific;

                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                    {
                        bool isError = false;
                        SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("10000003").Specific;

                        for (int i = 1; i <= oMatrix.RowCount; i++)
                        {
                            string UoMName1 = oMatrix.Columns.Item("10000003").Cells.Item(i).Specific.Value;

                            for (int j = i+1; j <= oMatrix.RowCount; j++)
                            {
                                string UoMName2 = oMatrix.Columns.Item("10000003").Cells.Item(j).Specific.Value;
                                if(UoMName1==UoMName2)
                                {
                                    Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("UoMNameDuplicated") + ": " + UoMName1, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                    isError = true;
                                }

                            }
                        }

                        if (isError)
                        {
                            BubbleEvent = !BubbleEvent;
                        }
                    }
                }

            }
        }
    }
}
