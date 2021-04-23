using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;
using System.Reflection;
using System.Resources;
using System.Runtime.InteropServices;
using System.Threading;
using SAPbobsCOM;
using static BDO_Localisation_AddOn.Program;

namespace BDO_Localisation_AddOn
{
    class StockRevaluation
    {
        public static void createFormItems(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = "";
            Dictionary<string, object> formItems = null;
            string itemName = "";
            double left = oForm.Items.Item("1002").Left;
            double width = oForm.Items.Item("1002").Width;
            double top = oForm.Items.Item("1002").Top + 20;
            double left_e = oForm.Items.Item("1003").Left;

            formItems = new Dictionary<string, object>();
            itemName = "LandCostS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left);
            formItems.Add("Width", width);
            formItems.Add("Top", top);
            formItems.Add("Height", 14);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("LandedCost"));

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "LandCostE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "UserDataSource");
            formItems.Add("TableName", "OMRV");
            formItems.Add("Alias", "DocNum");
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width);
            formItems.Add("Top", top);
            formItems.Add("Height", 13);
            formItems.Add("UID", itemName);
            formItems.Add("Enabled", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "LandCostLB"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            formItems.Add("Left", left_e - 20);
            formItems.Add("Top", top);
            formItems.Add("Height", 13);
            formItems.Add("UID", itemName);
            formItems.Add("LinkTo", "LandCostE");
            formItems.Add("LinkedObjectType", "69");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }
        }

        public static void uiApp_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (!Program.FORM_LOAD_FOR_ACTIVATE && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE && !pVal.BeforeAction)
                {
                    oForm.Items.Item("LandCostE").Specific.Value = "";
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD & pVal.BeforeAction)
                {
                    createFormItems(oForm, out errorText);
                    formDataLoad(oForm);
                }
            }
        }
        
        public static void createUserFields(out string errorText)
        {
            Dictionary<string, object> fieldskeysMap;

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BsDocEntry");
            fieldskeysMap.Add("TableName", "OMRV");
            fieldskeysMap.Add("Description", "Base Doc Entry");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Numeric);
            fieldskeysMap.Add("EditSize", 11);

            UDO.addUserTableFields(fieldskeysMap, out errorText);
        }

        public static void fillStockRevaluation(string docEntLC, out string docEntry)
        {
            SAPbobsCOM.Recordset oRecordSetCostVal = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            docEntry = getDocEntry(docEntLC);
            Program.uiApp.ActivateMenuItem("3086");
            
            SAPbouiCOM.Form oForm = Program.uiApp.Forms.ActiveForm;
            oForm.Freeze(true);

            SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("41").Specific;

            try
            {
                string itemCode = "";
                string lastItemCode = "";
                double allCostValByDocEnt = 0;

                string queryStockDoesNotExist = "select \"ItemCode\", \"WhsCode\", \"Quantity\" from IPF1 " + "\n"
                  + "where \"DocEntry\" = (select \"DocEntry\" from OIPF where \"DocEntry\" = '" + docEntLC + "')";

                oRecordSet.DoQuery(queryStockDoesNotExist);
                
                SAPbouiCOM.ComboBox ocomboBox = (SAPbouiCOM.ComboBox)oForm.Items.Item("1003").Specific;
                ocomboBox.Select("M");

                int row = 1;

                while (!oRecordSet.EoF)
                {
                    string str = oForm.Items.Item("1003").Specific.Value;
                    itemCode = oRecordSet.Fields.Item("ItemCode").Value; 
                    oMatrix.Columns.Item("6").Cells.Item(row).Specific.Value = itemCode;
                    string whs = CommonFunctions.getOADM("U_StockRevWh").ToString();
                    if (whs != "") oMatrix.Columns.Item("4").Cells.Item(row).Specific.Value = whs;
                    else oMatrix.Columns.Item("4").Cells.Item(row).Specific.Value = oRecordSet.Fields.Item("WhsCode").Value;
                    oMatrix.Columns.Item("1").Cells.Item(row).Specific.Value = oRecordSet.Fields.Item("Quantity").Value;
                    
                    SAPbouiCOM.Form oFormLC = Program.uiApp.Forms.GetForm("992", 1);
                    SAPbouiCOM.Matrix oMatrixLC = oFormLC.Items.Item("51").Specific;

                    int rowCount = oMatrixLC.RowCount;
                    string itemName = "";
                    string debit = "";
                    string credit = "";
                    for (int rowMatrix = 1; rowMatrix <= rowCount; rowMatrix++)
                    {
                        itemName = oMatrixLC.Columns.Item("1").Cells.Item(rowMatrix).Specific.Value;
                        LandedCosts.TtlCostLCFromJrnEntry(docEntLC, itemName, out debit, out credit);
                    }
                    double allocCostVal = 0;
                    if (debit != "")
                    {
                        allocCostVal = Convert.ToDouble(debit);
                    }

                    string query = "select \"TtlCostLC\", \"ItemCode\" " + "\n"
                     + "from IPF1 " + "\n"
                     + "where \"DocEntry\" = " + "\n"
                     + "(select \"DocEntry\" from OIPF " + "\n"
                     + "where \"DocEntry\" = '" + docEntLC + "')";

                    oRecordSetCostVal.DoQuery(query);
                    while (!oRecordSetCostVal.EoF)
                    {
                        allCostValByDocEnt = oRecordSetCostVal.Fields.Item("TtlCostLC").Value;
                        lastItemCode = oRecordSetCostVal.Fields.Item("ItemCode").Value;
                        if(itemCode == lastItemCode) oMatrix.Columns.Item("7").Cells.Item(row).Specific.Value = allCostValByDocEnt - lastAllCostVal(itemCode, docEntLC);
                        oRecordSetCostVal.MoveNext();
                    }
                    
                    row += 1;
                    oRecordSet.MoveNext();
                }


            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSetCostVal);
                Marshal.FinalReleaseComObject(oRecordSet);
                oForm.Freeze(false);
            }
            
        }

        public static string getDocEntry(string docEntry)
        {
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                string query = "select \"DocEntry\" from \"OMRV\" " + "\n"
                + "where \"U_BsDocEntry\" = '" + docEntry + "'";

                oRecordSet.DoQuery(query);
                if (!oRecordSet.EoF)
                {
                    return oRecordSet.Fields.Item("DocEntry").Value.ToString();
                }
                return string.Empty;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
            }
        }

        public static void formDataLoad(SAPbouiCOM.Form oForm)
        {
            oForm.Items.Item("LandCostE").Enabled = false;

            SAPbouiCOM.DBDataSource DocDBSource = oForm.DataSources.DBDataSources.Item(0);
            string DocNum = DocDBSource.GetValue("DocNum", 0);

            SAPbouiCOM.EditText oEdit = oForm.Items.Item("LandCostE").Specific;

            string docEntry = getBaseDocEntry(DocNum);
            if (docEntry != "")
                oEdit.Value = docEntry;
            else
                oEdit.Value = "";
            oForm.Update();
        }

        public static string getBaseDocEntry(string DocNum)
        {
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                if (DocNum == "") return string.Empty;
                
                string query ="select \"DocEntry\" from OIPF " + "\n"
                + "where \"DocEntry\" = ( " + "\n"
                + "select \"U_BsDocEntry\" from OMRV " + "\n"
                + "where \"DocEntry\" = ( " + "\n"
                + "select \"DocEntry\" from OMRV " + "\n"
                + "where \"DocNum\" = '" + DocNum + "'))";
                
                //string query = "select \"DocEntry\" from OMRV " + "\n"
                //+ "where \"DocNum\" = '" + DocNum + "'";

                oRecordSet.DoQuery(query);
                if (!oRecordSet.EoF)
                {
                    return oRecordSet.Fields.Item("DocEntry").Value.ToString();
                }
                return string.Empty;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
            }
        }

        public static void uiApp_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;

            SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(BusinessObjectInfo.FormTypeEx, Program.currentFormCount);

            if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD & !BusinessObjectInfo.BeforeAction)
            {
                formDataLoad(oForm);
            }
            if (!BusinessObjectInfo.BeforeAction && BusinessObjectInfo.ActionSuccess && BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD)
            {
                SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("41").Specific;
                createStockRevaluation(oForm, oMatrix);
            }
        }

        public static double lastAllCostVal(string itemCode, string docEntLC)
        {
            SAPbobsCOM.Recordset oRecordSetDocEnts = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset oRecordSetCost = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string docEntry = "";
            try
            {
                string queryDocEnts = "select \"CreatedBy\" from \"JDT1\" " + "\n"
                + "where \"TransType\" = '69' and \"Account\" = " + "\n"
                + "(select \"BalInvntAc\" from \"OITB\" " + "\n"
                + "where \"ItmsGrpCod\" = " + "\n"
                + "(select \"ItmsGrpCod\" from \"OITM\" " + "\n"
                + "where \"ItemCode\" = '" + itemCode + "')) " + "\n"
                + "order by \"TransId\" desc";

                oRecordSetDocEnts.DoQuery(queryDocEnts);

                while (!oRecordSetDocEnts.EoF)
                {
                    docEntry = oRecordSetDocEnts.Fields.Item("CreatedBy").Value.ToString();
                    string query = "select \"ItemCode\" from IPF1 " + "\n"
                    + "where \"DocEntry\" = ( " + "\n"
                    + "select \"DocEntry\" from OIPF " + "\n"
                    + "where \"DocEntry\" = '" + docEntry + "' ) and \"ItemCode\" = '" + itemCode + "'";

                    oRecordSet.DoQuery(query);
                    if (docEntry != docEntLC)
                    {
                        if (!oRecordSet.EoF) break;
                    }
                    oRecordSetDocEnts.MoveNext();
                }

                string queryCost = "select \"TtlCostLC\" " + "\n"
                + "from IPF1 " + "\n"
                + "where \"ItemCode\" = '" + itemCode + "' and \"DocEntry\" = " + "\n"
                + "(select \"DocEntry\" from OIPF " + "\n"
                + "where \"DocEntry\" = '" + docEntry + "')";

                oRecordSetCost.DoQuery(queryCost);

                if (!oRecordSetCost.EoF) return oRecordSetCost.Fields.Item("TtlCostLC").Value;

                return 0;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                Marshal.FinalReleaseComObject(oRecordSetCost);
                Marshal.FinalReleaseComObject(oRecordSetDocEnts);
            }
        }

        public static void createStockRevaluation(SAPbouiCOM.Form oForm, SAPbouiCOM.Matrix oMatrix)
        {

            SAPbouiCOM.Form oFormLC = Program.uiApp.Forms.GetForm("992", 1);
            SAPbouiCOM.DBDataSource DocDBSource = oFormLC.DataSources.DBDataSources.Item(0);
            string docEntLC = DocDBSource.GetValue("DocEntry", 0);

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string query = "UPDATE OMRV " + "\n"
            + "SET \"U_BsDocEntry\" = '" + docEntLC + "' " + "\n"
            + "where \"DocEntry\" = " + "\n"
            + "(select top 1 \"DocEntry\" from OMRV " + "\n"
            + "order by \"DocEntry\" desc)";

            oRecordSet.DoQuery(query);
            
            LandedCosts.formDataLoad(oFormLC);
        }
    }
}