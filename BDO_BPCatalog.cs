using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace BDO_Localisation_AddOn
{
    static partial class BDO_BPCatalog
    {
        /// <summary> ბიზნესპარტნიორების კატოლოგისთვის UserField - ის შექმნა </summary>
        /// <param name="Program.oCompany"></param>
        /// <param name="errorText"></param>
        public static void createUserFields( out string errorText)
        {
            errorText = null;
            Dictionary<string, object> fieldskeysMap;

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDO_SubDsc");
            fieldskeysMap.Add("TableName", "OSCN");
            fieldskeysMap.Add("Description", "BP Catalogue Description");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 254);
            fieldskeysMap.Add("Size", 20);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDO_UoMCod");
            fieldskeysMap.Add("TableName", "OSCN");
            fieldskeysMap.Add("Description", "Unit Of Measurement");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);
            fieldskeysMap.Add("Size", 20);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            GC.Collect();
        }

        public static void updateFields()
        {
            SAPbobsCOM.UserFieldsMD oUserFieldsMD = ((SAPbobsCOM.UserFieldsMD)(Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)));

            string query = @"select * from ""CUFD"" WHERE ""TableID"" = 'OSCN' AND ""AliasID"" = 'BDO_SubDsc' ";

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            oRecordSet.DoQuery(query);

            if (!oRecordSet.EoF)
            {
                oUserFieldsMD.GetByKey("OSCN", oRecordSet.Fields.Item("FieldID").Value);

                if (oUserFieldsMD.EditSize != 254)
                {
                    oUserFieldsMD.EditSize = 254;
                    var res = oUserFieldsMD.Update();
                }

                string error = Program.oCompany.GetLastErrorDescription();

                Marshal.ReleaseComObject(oRecordSet);
                oRecordSet = null;
                Marshal.ReleaseComObject(oUserFieldsMD);
                oUserFieldsMD = null;
                GC.WaitForPendingFinalizers();
            }
            else
            {
                Marshal.ReleaseComObject(oRecordSet);
                oRecordSet = null;
                Marshal.ReleaseComObject(oUserFieldsMD);
                oUserFieldsMD = null;
                GC.WaitForPendingFinalizers();
            }
        }

        public static void createFormItems( SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            try
            {
                bool multiSelection = false;
                string objectType = "10000199";
                string uniqueID_BaseDocCFL = "CFLUoMCdB";
                FormsB1.addChooseFromList( oForm, multiSelection, objectType, uniqueID_BaseDocCFL);

                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("17").Specific;
                oMatrix.Columns.Item("U_BDO_UoMCod").ChooseFromListUID = "CFLUoMCdB";
                oMatrix.Columns.Item("U_BDO_UoMCod").ChooseFromListAlias = "UoMCode";

                multiSelection = false;
                objectType = "10000199";
                uniqueID_BaseDocCFL = "CFLUoMCdI";
                FormsB1.addChooseFromList( oForm, multiSelection, objectType, uniqueID_BaseDocCFL);

                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("28").Specific;
                oMatrix.Columns.Item("U_BDO_UoMCod").ChooseFromListUID = "CFLUoMCdI";
                oMatrix.Columns.Item("U_BDO_UoMCod").ChooseFromListAlias = "UoMCode";
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
            finally
            {
                GC.Collect();
            }
        }

        public static void chooseFromList(SAPbouiCOM.Form oForm, SAPbouiCOM.ChooseFromListEvent oCFLEvento, out string errorText)
        {
            errorText = null;
            try
            {
                if (oCFLEvento.ChooseFromListUID == "CFLUoMCdB" || oCFLEvento.ChooseFromListUID == "CFLUoMCdI")
                {

                    SAPbouiCOM.DataTable oDataTableSelectedObjects = oCFLEvento.SelectedObjects;
                    string UoMCode = oDataTableSelectedObjects.GetValue("UomCode", 0);

                    SAPbouiCOM.Matrix oMatrix = null;

                    if (oCFLEvento.ChooseFromListUID == "CFLUoMCdB")
                    {
                        oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("17").Specific));
                    }
                    else
                    {
                        oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("28").Specific));
                    }

                    oMatrix.Columns.Item("U_BDO_UoMCod").Cells.Item(oCFLEvento.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    oForm.Freeze(false);
                    SAPbouiCOM.EditText UoMCodeEdit = oMatrix.Columns.Item("U_BDO_UoMCod").Cells.Item(oCFLEvento.Row).Specific;

                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                    {
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                    }
                       
                    UoMCodeEdit.Value = UoMCode;              

                  
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
            finally
            {
                GC.Collect();
            }
        }

        public static SAPbobsCOM.Recordset getCatalogEntryByBPBarcode(string CardCode, string ItemName, string Barcode, out string errorText)
        {
            errorText = null;

            SAPbobsCOM.BusinessPartners oBP;
            oBP = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);
            oBP.GetByKey(CardCode);

            string searchingParam = oBP.UserFields.Fields.Item("U_BDO_ItmPrm").Value;

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string query;
            try
            {
                CardCode = CardCode.Trim();

                if (searchingParam == "1") //დასახელებით
                {
                    if (ItemName.Length > 254)
                    {
                        ItemName = ItemName.Substring(0, 254);
                    }

                    query = @"SELECT * FROM ""OSCN""  WHERE ""U_BDO_SubDsc"" = N'" + ItemName.Replace("'", "''") + @"' AND ""CardCode""  = N'" + CardCode + "'";
                }
                else //კოდით
                {
                    query = @"SELECT * FROM ""OSCN"" WHERE ""Substitute"" = N'" + Barcode + @"' AND ""CardCode""  = N'" + CardCode + "'";
                }

                oRecordSet.DoQuery(query);

                if (!oRecordSet.EoF)
                {
                    return oRecordSet;
                }
                else
                {
                    return null;
                }
            }
            catch (Exception ex)
            {
                int errCode;
                string errMsg;

                Program.oCompany.GetLastError(out errCode, out errMsg);
                errorText = BDOSResources.getTranslate("ErrorDescription") + " : " + errMsg + "! " + BDOSResources.getTranslate("Code") + " : " + errCode + "! " + BDOSResources.getTranslate("OtherInfo") + " : " + ex.Message;
                return null;
            }
            finally
            {
                //Marshal.FinalReleaseComObject(oRecordSet);
                GC.Collect();
            }
        }

        public static SAPbobsCOM.Recordset getCatalogEntryByBPItmCode(string CardCode, string ItemName, string ItemCode,  out string errorText)
        {
            errorText = null;

            SAPbobsCOM.BusinessPartners oBP;
            oBP = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);
            oBP.GetByKey(CardCode);

            string searchingParam = oBP.UserFields.Fields.Item("U_BDO_ItmPrm").Value;

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string query = "";

            try
            {
                ItemCode = ItemCode.Replace("'", "");

                query = @"SELECT * FROM ""OSCN"" WHERE ""ItemCode"" = N'" + ItemCode + @"' AND ""CardCode"" = N'" + CardCode + "'";

                oRecordSet.DoQuery(query);

                if (!oRecordSet.EoF)
                {
                    return oRecordSet;
                }
                else
                {
                    return null;
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
        }

        public static void uiApp_ItemEvent(  string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
                {
                    //item column
                    BDO_BPCatalog.createFormItems( oForm, out errorText);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                {
                    if (pVal.BeforeAction == false)
                    {
                        SAPbouiCOM.ChooseFromListEvent oCFLEvento = null;
                        oCFLEvento = ((SAPbouiCOM.ChooseFromListEvent)(pVal));

                        BDO_BPCatalog.chooseFromList(oForm, oCFLEvento, out errorText);
                    }
                }
            }
        }
    }
}
