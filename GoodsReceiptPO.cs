using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Data;

namespace BDO_Localisation_AddOn
{
    class GoodsReceiptPO
    {
        public static void createUserFields(out string errorText)
        {
            errorText = null;
            //BDO_WBReceivedDocs.createUserFields( "OPDN", out errorText);
        }

        public static void createFormItems(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            BDO_WBReceivedDocs.createFormItems(oForm, "OPDN", out errorText);

            Dictionary<string, object> formItems;
            string itemName;

            SAPbouiCOM.Item oItem = oForm.Items.Item("70");
            int height = oItem.Height;
            int top = oItem.Top + height * 2 + 1;
            int left_s = oItem.Left;
            int width_s = oItem.Width;
            oItem = oForm.Items.Item("4");
            int left_e = oItem.Left;
            int width_e = oItem.Width;

            formItems = new Dictionary<string, object>();
            itemName = "BDOSACNumS"; //10 characters
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left_s);
            formItems.Add("Width", width_s);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("ACNumber"));
            formItems.Add("LinkTo", "BDOSACNumE");

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            formItems = new Dictionary<string, object>();
            itemName = "BDOSACNumE"; //10 characters
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", "OPDN");
            formItems.Add("Alias", "U_BDOSACNum");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left_e);
            formItems.Add("Width", width_e);
            formItems.Add("Top", top);
            formItems.Add("Height", height);
            formItems.Add("UID", itemName);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }
        }

        public static void attachWBToDoc(SAPbouiCOM.Form oForm, SAPbouiCOM.Form oIncWaybDocForm, out string errorText)
        {
            BDO_WBReceivedDocs.attachWBToDoc(oForm, oIncWaybDocForm, out errorText);
        }

        public static void uiApp_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool bubbleEvent)
        {
            bubbleEvent = true;
            string errorText;

            SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(BusinessObjectInfo.FormTypeEx, Program.currentFormCount);

            if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD)
            {
                if (BusinessObjectInfo.BeforeAction)
                {
                    SAPbouiCOM.DBDataSource DocDBSource = oForm.DataSources.DBDataSources.Item(0);
                    if (DocDBSource.GetValue("CANCELED", 0) == "N")
                    {
                        //დღგს თარიღის შევსება
                        //oForm.Freeze(true);
                        //int panelLevel = oForm.PaneLevel;
                        //string sdocDate = oForm.Items.Item("10").Specific.Value;
                        //oForm.PaneLevel = 7;
                        //oForm.Items.Item("1000").Specific.Value = sdocDate;
                        //oForm.PaneLevel = panelLevel;
                        //oForm.Freeze(false);


                        //დოკუმენტი არ დაემატოს ზედნადების გარეშე, თუ მომწოდებელს ჩართული აქვს
                        string CardCode = DocDBSource.GetValue("CardCode", 0);

                        SAPbobsCOM.BusinessPartners oBP;
                        oBP = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);
                        oBP.GetByKey(CardCode);

                        string RSControlType = oBP.UserFields.Fields.Item("U_BDO_MapCnt").Value;
                        string NeedWB = oBP.UserFields.Fields.Item("U_BDO_NeedWB").Value;
                        RSControlType = RSControlType.Trim();
                        NeedWB = NeedWB.Trim();

                        string DocType = DocDBSource.GetValue("DocType", 0);

                        if ((RSControlType == "2" || RSControlType == "3") && (DocType == "I"))
                        {
                            SAPbouiCOM.EditText oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("BDO_WBID").Specific;
                            string WBID = oEditText.Value;

                            if (WBID == "" && NeedWB == "Y")
                            {
                                bool isStock = false;

                                SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("38").Specific;

                                for (int row = 1; row <= oMatrix.RowCount; row++)
                                {
                                    // SAPbouiCOM.EditText Edtfieldtxt = oMatrix.Columns.Item("ItemCode").Cells.Item(row).Specific;
                                    string formItemCode = oMatrix.GetCellSpecific("1", row).Value;

                                    if (Items.isStockItem(formItemCode))
                                    {
                                        isStock = true;
                                        break;
                                    }
                                }

                                if (isStock)
                                {
                                    Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("BPControledOnRSLinkWaybillDocument"));
                                    bubbleEvent = !(bubbleEvent);
                                }
                            }
                            else
                            {
                                string Doctype = "";

                                if (BusinessObjectInfo.Type == "18")
                                {
                                    Doctype = "APInvoice";
                                }
                                else if (BusinessObjectInfo.Type == "19")
                                {
                                    Doctype = "CredMemo";
                                }
                                else if (BusinessObjectInfo.Type == "20")
                                {
                                    Doctype = "GoodsReceiptPO";
                                }
                                try
                                {
                                    bool continuePosting = BDO_WBReceivedDocs.waybillsCompare(WBID, oForm, RSControlType, Doctype, out errorText);

                                    if (continuePosting == false)
                                    {
                                        Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("GoodsTableNotMatchedESTable"));
                                        Program.uiApp.MessageBox(BDOSResources.getTranslate("OperationUnsuccesfullSeeLog"));
                                        bubbleEvent = !(bubbleEvent);
                                    }
                                }
                                catch { }
                            }
                        }
                    }
                }
            }

            //A/C Number Update
            if ((BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD || BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE)
                && BusinessObjectInfo.ActionSuccess && !BusinessObjectInfo.BeforeAction)
            {
                CommonFunctions.StartTransaction();

                SAPbouiCOM.DBDataSource DocDBSource = oForm.DataSources.DBDataSources.Item(0);
                string DocEntry = DocDBSource.GetValue("DocEntry", 0);
                string ObjType = DocDBSource.GetValue("ObjType", 0);
                string ACNumber = DocDBSource.GetValue("U_BDOSACNum", 0);

                JournalEntry.UpdateJournalEntryACNumber(DocEntry, ObjType, ACNumber, out errorText);
                if (string.IsNullOrEmpty(errorText))
                {
                    CommonFunctions.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                }
                else
                {
                    Program.uiApp.MessageBox(errorText);
                    CommonFunctions.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }
            }

            if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD && !BusinessObjectInfo.BeforeAction)
            {
                BDO_WBReceivedDocs.setwaybillText(oForm);
            }
        }

        public static void uiApp_ItemEvent(ref SAPbouiCOM.ItemEvent pVal, out bool bubbleEvent)
        {
            bubbleEvent = true;
            string errorText;

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
                {
                    if (pVal.BeforeAction)
                    {
                        createFormItems(oForm, out errorText);
                        FormsB1.WB_TAX_AuthorizationsItems(oForm);
                    }
                }
                else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_COMBO_SELECT)
                {
                    if (!pVal.BeforeAction)
                    {
                        if (pVal.ItemUID == "WBOper")
                        {
                            Program.oIncWaybDocFormGdsRecpPO = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                            oForm.Freeze(true);
                            BDO_WBReceivedDocs.comboSelect(oForm, Program.oIncWaybDocFormGdsRecpPO, "GoodsReceiptPO", out errorText);
                            oForm.Freeze(false);
                        }
                    }
                }               
            }
        }
    }
}
