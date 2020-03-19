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
            fieldskeysMap.Add("Name", "BaseDocNum");
            fieldskeysMap.Add("TableName", "OMRV");
            fieldskeysMap.Add("Description", "Base Doc Num");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Numeric);
            fieldskeysMap.Add("EditSize", 11);

            UDO.addUserTableFields(fieldskeysMap, out errorText);
        }

        public static void fillStockRevaluation(string docNum, out string docEntry)
        {
            SAPbobsCOM.Recordset oRecordSetCostVal = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.MaterialRevaluation m_MaterialRev = (MaterialRevaluation)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oMaterialRevaluation);
            SAPbobsCOM.MaterialRevaluation_lines m_MaterialRevLines = m_MaterialRev.Lines;

            try
            {
                string itemCode = "";
                string lastItemCode = "";
                double allCostValByDocNum = 0;

                string queryStockDoesNotExist = "select \"ItemCode\", \"WhsCode\", \"Quantity\" from IPF1 " + "\n"
                  + "where \"DocEntry\" = (select \"DocEntry\" from OIPF where \"DocNum\" = '" + docNum + "')";

                oRecordSet.DoQuery(queryStockDoesNotExist);

                m_MaterialRev.DocDate = DateTime.Now;
                m_MaterialRev.RevalType = "M";
                m_MaterialRev.UserFields.Fields.Item("U_BaseDocNum").Value = docNum;
                int row = 0;

                while (!oRecordSet.EoF)
                {
                    m_MaterialRevLines.SetCurrentLine(row);
                    itemCode = oRecordSet.Fields.Item("ItemCode").Value;
                    m_MaterialRevLines.Quantity = oRecordSet.Fields.Item("Quantity").Value;
                    m_MaterialRevLines.WarehouseCode = oRecordSet.Fields.Item("WhsCode").Value;
                    m_MaterialRevLines.ItemCode = itemCode;


                    double allocCostVal = lastAllCostVal(itemCode, docNum);
                    string query = "select \"TtlExpndLC\", \"ItemCode\" " + "\n"
                     + "from IPF1 " + "\n"
                     + "where \"DocEntry\" = " + "\n"
                     + "(select \"DocEntry\" from OIPF " + "\n"
                     + "where \"DocNum\" = '" + docNum + "')";

                    oRecordSetCostVal.DoQuery(query);
                    while (!oRecordSetCostVal.EoF)
                    {
                        allCostValByDocNum = oRecordSetCostVal.Fields.Item("TtlExpndLC").Value;
                        lastItemCode = oRecordSetCostVal.Fields.Item("ItemCode").Value;
                        if(itemCode == lastItemCode)
                        {
                            m_MaterialRevLines.DebitCredit = allCostValByDocNum - allocCostVal;
                        }
                        oRecordSetCostVal.MoveNext();
                    }
                    m_MaterialRevLines.Add();

                    row += 1;
                    oRecordSet.MoveNext();
                }

                int res = m_MaterialRev.Add();
                int errCode = oCompany.GetLastErrorCode();
                string errMsg;
                Program.oCompany.GetLastError(out errCode, out errMsg);
                if (errMsg == "At least one amount is required in document rows ") errMsg = "Stock revaluation document can't be created because item type is - Batch/Serial number";
                if (errMsg != "") uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate(errMsg), SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);

                docEntry = getDocEntry(docNum);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                Marshal.FinalReleaseComObject(m_MaterialRev);
                Marshal.FinalReleaseComObject(m_MaterialRevLines);
                Marshal.FinalReleaseComObject(oRecordSetCostVal);
                Marshal.FinalReleaseComObject(oRecordSet);
            }
        }

        public static string getDocEntry(string docNum)
        {
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                string query = "select \"DocEntry\" from \"OMRV\" " + "\n"
            + "where \"U_BaseDocNum\" = '" + docNum + "'";

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

            oEdit.Value = docEntry;
            oForm.Update();
        }

        public static string getBaseDocEntry(string docNum)
        {
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                string query = "select \"DocEntry\" from \"OIPF\" " + "\n"
                + "where \"DocNum\" = (select \"U_BaseDocNum\" from OMRV " + "\n"
                + "where \"DocNum\" = '" + docNum + "')";

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
        }

        public static double lastAllCostVal(string itemCode, string docNumLC)
        {
            SAPbobsCOM.Recordset oRecordSetDocNums = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset oRecordSetCost = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string docNum = "";

            string queryDocNums = "select \"BaseRef\" from \"JDT1\" " + "\n"
            + "where \"TransType\" = '69' and \"Account\" = " + "\n"
            + "(select \"BalInvntAc\" from \"OITB\" " + "\n"
            + "where \"ItmsGrpCod\" = " + "\n"
            + "(select \"ItmsGrpCod\" from \"OITM\" " + "\n"
            + "where \"ItemCode\" = '" + itemCode + "')) " + "\n"
            + "order by \"TransId\" desc";

            oRecordSetDocNums.DoQuery(queryDocNums);
            
            while (!oRecordSetDocNums.EoF)
            {
                docNum = oRecordSetDocNums.Fields.Item("BaseRef").Value.ToString();
                string query =  "select \"ItemCode\" from IPF1 " + "\n"
                + "where \"DocEntry\" = ( " + "\n"
                + "select \"DocEntry\" from OIPF " + "\n"
                + "where \"DocNum\" = '" + docNum +"' ) and \"ItemCode\" = '" + itemCode + "'";

                oRecordSet.DoQuery(query);
                if (docNum != docNumLC)
                {
                    if (!oRecordSet.EoF) break;
                }
                oRecordSetDocNums.MoveNext();
            }

            string queryCost = "select \"TtlExpndLC\" " + "\n"
            + "from IPF1 " + "\n"
            + "where \"ItemCode\" = '" + itemCode + "' and \"DocEntry\" = " + "\n"
            + "(select \"DocEntry\" from OIPF " + "\n"
            + "where \"DocNum\" = '" + docNum + "')";

            oRecordSetCost.DoQuery(queryCost);

            if (!oRecordSetCost.EoF) return oRecordSetCost.Fields.Item("TtlExpndLC").Value;

            return 0;
        }
    }
}