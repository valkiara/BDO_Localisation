using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BDO_Localisation_AddOn
{
    static partial class VatGroup
    {
        public static void MatrixLink(  SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            bool multiSelection = false;
            string objectType = "1"; //Accounting
            FormsB1.addChooseFromList( oForm, multiSelection, objectType, "Acc_CFL");
            FormsB1.addChooseFromList( oForm, multiSelection, objectType, "AccCVt_CFL");

            SAPbouiCOM.ChooseFromList oCFL;
            SAPbouiCOM.Conditions oCons;
                SAPbouiCOM.Condition oCon;

            oCFL = oForm.ChooseFromLists.Item("Acc_CFL");
            oCons = oCFL.GetConditions();
            oCon = oCons.Add();
            oCon.Alias = "Postable";
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCon.CondVal = "Y";
            oCFL.SetConditions(oCons);

            oCFL = oForm.ChooseFromLists.Item("AccCVt_CFL");
            oCons = oCFL.GetConditions();
            oCon = oCons.Add();
            oCon.Alias = "Postable";
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCon.CondVal = "Y";
            oCFL.SetConditions(oCons);



            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("3").Specific));

            SAPbouiCOM.Column oColumn;

            oColumn = oMatrix.Columns.Item("U_BDOSAccF");
            oColumn.ChooseFromListUID = "Acc_CFL";
            oColumn.ChooseFromListAlias = "AcctCode";

            oColumn = oMatrix.Columns.Item("U_BDOSAccCVt");
            oColumn.ChooseFromListUID = "AccCVt_CFL";
            oColumn.ChooseFromListAlias = "AcctCode";

        }

        public static void chooseFromList(SAPbouiCOM.Form oForm, SAPbouiCOM.ChooseFromListEvent oCFLEvento, out string errorText)
        {
            errorText = null;
            try
            {
                if (oCFLEvento.ChooseFromListUID == "Acc_CFL")
                {
                    SAPbouiCOM.DataTable oDataTableSelectedObjects = oCFLEvento.SelectedObjects;
                    string AcctCode = oDataTableSelectedObjects.GetValue("AcctCode", 0);

                    SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("3").Specific));

                    oMatrix.Columns.Item("U_BDOSAccF").Cells.Item(oCFLEvento.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    SAPbouiCOM.EditText AcctCodeEdit = oMatrix.Columns.Item("U_BDOSAccF").Cells.Item(oCFLEvento.Row).Specific;
                    AcctCodeEdit.Value = AcctCode;

                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                    {
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                    }
                }
                if (oCFLEvento.ChooseFromListUID == "AccCVt_CFL")
                {
                    SAPbouiCOM.DataTable oDataTableSelectedObjects = oCFLEvento.SelectedObjects;
                    string AcctCode = oDataTableSelectedObjects.GetValue("AcctCode", 0);

                    SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("3").Specific));

                    oMatrix.Columns.Item("U_BDOSAccCVt").Cells.Item(oCFLEvento.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    SAPbouiCOM.EditText AcctCodeEdit = oMatrix.Columns.Item("U_BDOSAccCVt").Cells.Item(oCFLEvento.Row).Specific;
                    AcctCodeEdit.Value = AcctCode;

                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                    {
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                    }
                }
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                }
            }
            finally
            {
                GC.Collect();
            }
        }

        public static void createUserFields( out string errorText)
        {
            errorText = null;
            Dictionary<string, object> fieldskeysMap;

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSAccF");
            fieldskeysMap.Add("TableName", "OVTG");
            fieldskeysMap.Add("Description", "Tax Account For Budget");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSAccCVt");
            fieldskeysMap.Add("TableName", "OVTG");
            fieldskeysMap.Add("Description", BDOSResources.getTranslate("CustomVATPaidAccount"));
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields( fieldskeysMap, out errorText);

            GC.Collect();
        }

        public static void uiApp_ItemEvent(  string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD & pVal.BeforeAction == false)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
                {
                    VatGroup.MatrixLink( oForm, out errorText);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                {
                    if (pVal.BeforeAction == false)
                    {
                        SAPbouiCOM.ChooseFromListEvent oCFLEvento = null;
                        oCFLEvento = ((SAPbouiCOM.ChooseFromListEvent)(pVal));

                        VatGroup.chooseFromList(oForm, oCFLEvento, out errorText);
                    }
                }
            }
        }
    }
}
