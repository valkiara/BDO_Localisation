using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace BDO_Localisation_AddOn
{
    static partial class APDownPaymentInvoice
    {
        public static void uiApp_FormDataEvent(SAPbouiCOM.Application uiApp, SAPbobsCOM.Company oCompany, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;

            SAPbouiCOM.Form oForm = uiApp.Forms.GetForm(BusinessObjectInfo.FormTypeEx, Program.currentFormCount);

            if (oForm.TypeEx == "65301")
            {              


                if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD & BusinessObjectInfo.BeforeAction == true)
                {
                    oForm.Freeze(true);
                    int panelLevel = oForm.PaneLevel;
                    string sdocDate = oForm.Items.Item("10").Specific.Value;
                    oForm.PaneLevel = 7;
                    oForm.Items.Item("1000").Specific.Value = sdocDate;
                    oForm.PaneLevel = panelLevel;
                    oForm.Freeze(false);

                    
                    SAPbouiCOM.DBDataSource DocDBSourceOCRD = oForm.DataSources.DBDataSources.Item(0);

                    string VatStatus = oForm.DataSources.DBDataSources.Item("OCRD").GetValue("VatStatus", 0);
                    string errorText = "";

                    if (VatStatus.Trim() == "E")
                    {
                        WithholdingTax.JrnEntryAPInvoiceCredidNoteCheck(oCompany, oForm, BusinessObjectInfo.Type, out errorText);

                        if (errorText != null)
                        {
                            uiApp.MessageBox(errorText);
                            BubbleEvent = false;
                            return;
                        }
                    }

                }

                //გატარება
                if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD & BusinessObjectInfo.ActionSuccess & BusinessObjectInfo.BeforeAction == false)
                {
                    string errorText = "";
                    
                    SAPbouiCOM.DBDataSource DocDBSourceOCRD = oForm.DataSources.DBDataSources.Item(0);

                    string VatStatus = oForm.DataSources.DBDataSources.Item("OCRD").GetValue("VatStatus", 0);

                    if (VatStatus.Trim() == "E")
                    {
                        string DocEntry = DocDBSourceOCRD.GetValue("DocEntry", 0);
                        string DocNum = DocDBSourceOCRD.GetValue("DocNum", 0);
                        //DateTime DocDate = Convert.ToDateTime(DocDBSourceOCRD.GetValue("DocDate", 0));
                        DateTime DocDate = DateTime.ParseExact(DocDBSourceOCRD.GetValue("DocDate", 0), "yyyyMMdd", CultureInfo.InvariantCulture);
                        WithholdingTax.JrnEntryAPInvoiceCredidNote(oCompany, oForm, BusinessObjectInfo.Type, DocEntry, DocNum, DocDate, out errorText);
                        if (errorText != null)
                        {
                            uiApp.MessageBox(errorText);
                            BubbleEvent = false;
                        }
                    }
                }
            }
        }
    }
}
