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
        public static void createUserFields( out string errorText)
        {
            errorText = null;
            //BDO_WBReceivedDocs.createUserFields( "OPDN", out errorText);
        }

        public static void createFormItems(  SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            BDO_WBReceivedDocs.createFormItems( oForm, "OPDN", out errorText);
            
            Dictionary<string, object> formItems = null;
            string itemName = "";

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
            formItems.Add("DisplayDesc", true);
            formItems.Add("Enabled", true);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            GC.Collect();
        }

        public static void attachWBToDoc(  SAPbouiCOM.Form oForm, SAPbouiCOM.Form oIncWaybDocForm, out string errorText)
        {
            errorText = null;
            BDO_WBReceivedDocs.attachWBToDoc( oForm, oIncWaybDocForm, out errorText);
        }

        public static void formDataLoad( SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            oForm.Freeze(true);
            try
            {
                setVisibleFormItems(oForm, out errorText);

                //-------------------------------------------სასაქონლო ზედნადები----------------------------------->              
                BDO_WBReceivedDocs.setwaybillText( oForm, out errorText);
                //<-------------------------------------------სასაქონლო ზედნადები-----------------------------------

                int docEntry = Convert.ToInt32(oForm.DataSources.DBDataSources.Item("OPDN").GetValue("DocEntry", 0));
                string cardCode = oForm.DataSources.DBDataSources.Item("OPDN").GetValue("CardCode", 0).Trim();
                
            }
            catch (Exception ex)
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE && oForm.DataSources.DBDataSources.Item("OPDN").GetValue("U_BDO_WBID", 0).Trim() == "")
                {

                }
                else
                {
                    SAPbouiCOM.EditText oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("BDO_WBNo").Specific;
                    oEditText.Value = " ";

                    oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("BDO_WBID").Specific;
                    oEditText.Value = " ";

                    oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("actDate").Specific;
                    oEditText.Value = "00010101";

                    SAPbouiCOM.ComboBox oCombobox = (SAPbouiCOM.ComboBox)oForm.Items.Item("BDO_WBSt").Specific;
                    oCombobox.Select("-1");

                    SAPbouiCOM.CheckBox oCheckBox = (SAPbouiCOM.CheckBox)oForm.Items.Item("WBrec").Specific;
                    oCheckBox.Checked = false;

                }

                SAPbouiCOM.StaticText oStaticText = (SAPbouiCOM.StaticText)oForm.Items.Item("WBInfoST").Specific;
                oStaticText.Caption = BDOSResources.getTranslate("NotLinked");

                errorText = ex.Message;

            }
            finally
            {
                oForm.Freeze(false);
                GC.Collect();
            }
        }

        public static void setVisibleFormItems(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;
            //oForm.Freeze(true);

            //try
            //{
            //    string docEntrySTR = oForm.DataSources.DBDataSources.Item("OPDN").GetValue("DocEntry", 0).Trim();
                
            //}
            //catch (Exception ex)
            //{
            //    errorText = ex.Message;
            //}
            //finally
            //{
            //    oForm.Freeze(false);
            //    oForm.Update();
            //    GC.Collect();
            //}
        }

        public static void formDataAddUpdate( SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            //try
            //{

            //}
            //catch (Exception ex)
            //{
            //    errorText = ex.Message;
            //}
            //finally
            //{
            //    GC.Collect();
            //}
        }

        public static void uiApp_FormDataEvent(  ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(BusinessObjectInfo.FormTypeEx, Program.currentFormCount);

            if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD)
            {
                if (BusinessObjectInfo.BeforeAction == true)
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

                            if (WBID == "" & NeedWB == "Y")
                            {
                                bool isStock = false;

                                SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("38").Specific;

                                for (int row = 1; row <= oMatrix.RowCount; row++)
                                {
                                    // SAPbouiCOM.EditText Edtfieldtxt = oMatrix.Columns.Item("ItemCode").Cells.Item(row).Specific;
                                    string formItemCode = oMatrix.GetCellSpecific("1", row).Value;

                                    if (Items.isStockItem( formItemCode))
                                    {
                                        isStock = true;
                                        break;
                                    }
                                }

                                if (isStock)
                                {
                                    Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("BPControledOnRSLinkWaybillDocument"));
                                    BubbleEvent = !(BubbleEvent);
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
                                    bool continuePosting = BDO_WBReceivedDocs.waybillsCompare( WBID, oForm, RSControlType, Doctype, out errorText);

                                    if (continuePosting == false)
                                    {
                                        Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("GoodsTableNotMatchedESTable"));
                                        Program.uiApp.MessageBox(BDOSResources.getTranslate("OperationUnsuccesfullSeeLog"));
                                        BubbleEvent = !(BubbleEvent);
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
                && BusinessObjectInfo.ActionSuccess == true && BusinessObjectInfo.BeforeAction == false)
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

                AddWblIDAndNumberInJrnEntry(oForm, out errorText);
            }
                        
            if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE)
            {
                if (BusinessObjectInfo.BeforeAction == true)
                {
                    formDataAddUpdate( oForm, out errorText);
                    if (string.IsNullOrEmpty(errorText) == false)
                    {
                        Program.uiApp.MessageBox(errorText);
                        BubbleEvent = false;
                    }
                }
                else if (BusinessObjectInfo.BeforeAction == false && BusinessObjectInfo.ActionSuccess == true)
                {
                    setVisibleFormItems(oForm, out errorText);
                }
            }

            if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD & BusinessObjectInfo.BeforeAction == false)
            {
                formDataLoad( oForm, out errorText);
                setVisibleFormItems(oForm, out errorText);
            }            
        }

        public static void uiApp_ItemEvent(  string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            {
                SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD & pVal.BeforeAction == true)
                {
                    createFormItems( oForm, out errorText);
                    formDataLoad( oForm, out errorText);
                    setVisibleFormItems(oForm, out errorText);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_COMBO_SELECT & pVal.ItemUID == "WBOper" & pVal.BeforeAction == false)
                {
                    Program.oIncWaybDocFormGdsRecpPO = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                    oForm.Freeze(true);
                    BDO_WBReceivedDocs.comboSelect(  oForm, Program.oIncWaybDocFormGdsRecpPO, pVal.ItemUID, "GoodsReceiptPO", out errorText);
                    oForm.Freeze(false);
                }

            }
        }

        public static void AddWblIDAndNumberInJrnEntry(SAPbouiCOM.Form oForm, out string errorText)
        {
            CommonFunctions.StartTransaction();

            SAPbouiCOM.DBDataSource DocDBSource = oForm.DataSources.DBDataSources.Item(0);
            string DocEntry = DocDBSource.GetValue("DocEntry", 0);
            string ObjType = DocDBSource.GetValue("ObjType", 0);

            string WblId = DocDBSource.GetValue("U_BDO_WBID", 0);
            string WblNum = DocDBSource.GetValue("U_BDO_WBNo", 0);

            JournalEntry.UpdateJournalEntryWblIdAndNumber(DocEntry, ObjType, WblId, WblNum, out errorText);

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


    }
}
