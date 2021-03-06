using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Globalization;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using System.Data;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;
using SAPbobsCOM;
using SAPbouiCOM;

namespace BDO_Localisation_AddOn
{
    static partial class BDO_WaybillsJournalReceived
    {

        public static Dictionary<string, string[][]> wbTempLines = new Dictionary<string, string[][]>();
        public static int WBGdMatrixRow = 0;
        public static decimal WBGdMatrixMaxQty = 0;
        public static decimal WBGdMatrixNewQty = 0;
        public static string itemCodeOld;

        public static void setUomCodeBtRSCode(SAPbouiCOM.Form oForm, int Row, out string errorText)
        {
            errorText = null;

            string WBGUntCode = "";
            string ItemCode;
            string WBUntCdRS;

            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("WBGdMatrix").Specific;

            ItemCode = oMatrix.GetCellSpecific("ItemCode", Row).Value;
            WBUntCdRS = oMatrix.GetCellSpecific("WBUntCdRS", Row).Value;
            WBGUntCode = oMatrix.GetCellSpecific("WBUntCode", Row).Value;

            SAPbobsCOM.Recordset oRecordsetbyRSCODE = BDO_RSUoM.getUomByRSCode(ItemCode, WBUntCdRS, out errorText);
            SAPbouiCOM.EditText oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item("WBUntCode").Cells.Item(Row).Specific;

            if (oRecordsetbyRSCODE != null)
            {
                WBGUntCode = oRecordsetbyRSCODE.Fields.Item("UomCode").Value;

                try
                {
                    oEditText.Value = WBGUntCode;
                }
                catch
                {
                }


                //if (true)
                //{
                //    string WBUntName = oRecordsetbyRSCODE.Fields.Item("UomName").Value;
                //    SAPbouiCOM.EditText oEditText = (SAPbouiCOM.EditText)oMatrix.Columns.Item("WBUntName").Cells.Item(Row).Specific;

                //    if ((oEditText.Value) != WBUntName)
                //        try
                //        {
                //            oEditText.Value = WBUntName;
                //        }
                //        catch
                //        {
                //        }
                //}
            }
            else
            {
                oEditText.Value = "";
            }
        }
        public static string DetectVATByRSCode(string RSVatCode, out string errorText)
        {
            errorText = null;

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string query = "SELECT \"U_BDO_RSVAT" + RSVatCode + "\" FROM \"OADM\" ";
            oRecordSet.DoQuery(query);

            if (!oRecordSet.EoF)
            {
                return oRecordSet.Fields.Item("U_BDO_RSVAT" + RSVatCode).Value;
            }

            return null;
        }
        public static double detectVATRate(string VatCode, out string errorText)
        {
            errorText = null;

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string query = "SELECT \"Rate\" FROM \"OVTG\" WHERE \"Code\"='" + VatCode + "'"; ;
            oRecordSet.DoQuery(query);

            if (!oRecordSet.EoF)
            {
                return oRecordSet.Fields.Item("Rate").Value;
            }

            return 0;
        }
        public static void createWaybillIncDocs(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            SAPbobsCOM.BusinessPartners oBP;
            SAPbobsCOM.PriceLists oPL;
            int pricelistnum;

            oForm.Freeze(true);

            SAPbouiCOM.Matrix oMatrix = oForm.Items.Item("WBMatrix").Specific;
            oMatrix.FlushToDataSource();

            SAPbouiCOM.Matrix oMatrixGoods = oForm.Items.Item("WBGdMatrix").Specific;
            oMatrixGoods.FlushToDataSource();

            string oGdsRcpt = oForm.DataSources.UserDataSources.Item("GdsRcpt").ValueEx;

            var asDraft = oForm.DataSources.UserDataSources.Item("AsDraft").ValueEx == "Y";
            var draftType = "APInvoiceDraft";

            for (int row = 1; row <= oMatrix.RowCount; row++)
            {
                SAPbouiCOM.CheckBox Edtfield = oMatrix.Columns.Item("WBCheckbox").Cells.Item(row).Specific;
                bool checkedLine = (Edtfield.Checked);

                bool saveWhs = oMatrix.Columns.Item("WBCheckb").Cells.Item(row).Specific.Checked;

                if (checkedLine)
                {
                    SAPbouiCOM.ComboBox statusOfInvoice = (SAPbouiCOM.ComboBox)oMatrix.Columns.Item("WBStat").Cells.Item(row).Specific;
                    string statusOFINVOICE = statusOfInvoice.Value;
                    if (statusOFINVOICE != "5")
                    {


                        CommonFunctions.StartTransaction();
                        SAPbobsCOM.Documents APInv = null;
                        bool NotToCreate = false;
                        SAPbouiCOM.ComboBox ComboboxStatus = (SAPbouiCOM.ComboBox)oMatrix.Columns.Item("TYPE").Cells.Item(row).Specific;
                        string TYPE = ComboboxStatus.Value;
                        if (TYPE == "Procurement") //2
                            if (oGdsRcpt == "1")
                            {
                                if (!asDraft)
                                {
                                    APInv = Program.oCompany.GetBusinessObject(BoObjectTypes
                                        .oPurchaseDeliveryNotes);
                                }

                                else
                                {
                                    APInv = Program.oCompany.GetBusinessObject(BoObjectTypes.oDrafts);
                                    draftType = "GdsRcptDraft";
                                }
                            }
                            else
                            {
                                if (!asDraft)
                                {
                                    APInv = Program.oCompany.GetBusinessObject(BoObjectTypes.oPurchaseInvoices);
                                }
                                else
                                {
                                    APInv = Program.oCompany.GetBusinessObject(BoObjectTypes.oDrafts);
                                    draftType = "APInvoiceDraft";
                                }
                            }
                        else if (TYPE == "Return") //1
                            APInv = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes);

                        SAPbouiCOM.EditText Edtfieldtxt = oMatrix.Columns.Item("WBID").Cells.Item(row).Specific;
                        string WBID = Edtfieldtxt.Value;

                        Edtfieldtxt = oForm.Items.Item("Whs").Specific;
                        string whs = Edtfieldtxt.Value;

                        Edtfieldtxt = oForm.Items.Item("PrjCode").Specific;
                        string PrjCode = Edtfieldtxt.Value;

                        Edtfieldtxt = oMatrix.Columns.Item("WBNo").Cells.Item(row).Specific;
                        string WBNo = Edtfieldtxt.Value;

                        Edtfieldtxt = oMatrix.Columns.Item("WBSupTIN").Cells.Item(row).Specific;
                        string TIN = Edtfieldtxt.Value;

                        SAPbouiCOM.ComboBox Combobox = oMatrix.Columns.Item("WBStat").Cells.Item(row).Specific;
                        string WBStat = Combobox.Value;

                        string cardName;
                        string CardCode = BusinessPartners.GetCardCodeByTin(TIN, "S", out cardName);

                        if (CardCode == null)
                        {
                            oForm.Freeze(false);
                            errorText = BDOSResources.getTranslate("BPNotFound") + BDOSResources.getTranslate("BPTin") + " : " + TIN;
                            return;
                        }
                        //Edtfieldtxt = oMatrix.Columns.Item("WBActDate").Cells.Item(row).Specific;
                        DateTime WBActDate = oForm.DataSources.DataTables.Item("WBTable").GetValue("WBActDate", row - 1);
                        DateTime WBStartDat = oForm.DataSources.DataTables.Item("WBTable").GetValue("WBStartDat", row - 1);
                        string WBBlankAgr = oForm.DataSources.DataTables.Item("WBTable").GetValue("WBBlankAgr", row - 1);
                        string comment = oForm.DataSources.DataTables.Item("WBTable").GetValue("WBCOMMENT", row - 1);
                        string wBWhs = oForm.DataSources.DataTables.Item("WBTable").GetValue("WBWhs", row - 1);
                        string wBProject = oForm.DataSources.DataTables.Item("WBTable").GetValue("WBProject", row - 1);
                        string wBEndAdr = oForm.DataSources.DataTables.Item("WBTable").GetValue("WBEndAdd", row - 1);

                        if (!string.IsNullOrEmpty(wBWhs))
                        {
                            whs = wBWhs;
                        }
                        else if(!string.IsNullOrEmpty(wBProject))
                        {
                            oForm.Freeze(false);
                            errorText = BDOSResources.getTranslate("PleaseFillWarehouse")+". "+BDOSResources.getTranslate("Row") +": " + row;
                            Program.uiApp.StatusBar.SetSystemMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            return;
                        }

                        if (saveWhs) BDOSWarehouseAddresses.AddWhsByAddress(wBEndAdr, whs);

                        //SAPbobsCOM.CompanyService oCompanyService;
                        //SAPbobsCOM.BlanketAgreementsService oAcuerdoServicio;
                        //SAPbobsCOM.BlanketAgreement oAcuerdo;
                        //SAPbobsCOM.BlanketAgreementParams oParams;
                        //// Initialize it
                        //oCompanyService = Program.oCompany.GetCompanyService();
                        //oAcuerdoServicio = oCompanyService.GetBusinessService(SAPbobsCOM.ServiceTypes.BlanketAgreementsService);
                        //oParams = oAcuerdoServicio.GetDataInterface(SAPbobsCOM.BlanketAgreementsServiceDataInterfaces.basBlanketAgreementParams);

                        //oParams.AgreementNo = Convert.ToInt32(WBBlankAgr);
                        //oAcuerdo = oAcuerdoServicio.GetBlanketAgreement(oParams);

                        //int PaymentGroupCode = oAcuerdo.PaymentTerms;


                        SAPbobsCOM.Recordset oRecordSetWH = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        string queryPr = @"SELECT ""U_BDOSPrjCod"" FROM ""OWHS"" WHERE ""WhsCode"" = '" + whs + "'";

                        oRecordSetWH.DoQuery(queryPr);
                        if (!string.IsNullOrEmpty(oRecordSetWH.Fields.Item("U_BDOSPrjCod").Value))
                        {
                            APInv.Project = oRecordSetWH.Fields.Item("U_BDOSPrjCod").Value;
                        }

                        if (PrjCode != "")
                        {
                            APInv.Project = PrjCode;
                        }

                        var blanketProject = "";
                        if (WBBlankAgr != "")
                        {
                            APInv.BlanketAgreementNumber = Convert.ToInt32(WBBlankAgr);
                            SAPbobsCOM.Recordset oRecordSetBA = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            string query = @"SELECT ""Project"" FROM ""OOAT"" WHERE ""AbsID"" = '" + Convert.ToInt32(WBBlankAgr) + "'";

                            oRecordSetBA.DoQuery(query);


                            if (!oRecordSetBA.EoF)
                            {
                                blanketProject = oRecordSetBA.Fields.Item("Project").Value;
                                APInv.Project = blanketProject;
                            }

                        }

                        if (wBProject != "")
                        {
                            APInv.Project = wBProject;
                        }
                        else
                        {
                            oForm.Freeze(false);
                            errorText = BDOSResources.getTranslate("PleaseFillTheProject") + ". " + BDOSResources.getTranslate("Row") + ": " + row;
                            Program.uiApp.StatusBar.SetSystemMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            return;
                        }

                        //APInv.PaymentGroupCode = PaymentGroupCode;
                        APInv.CardCode = CardCode;
                        APInv.DocDate = WBStartDat;
                        APInv.VatDate = WBStartDat;
                        APInv.TaxDate = WBStartDat;
                        APInv.Comments = comment;
                        //APInv.DocCurrency = Program.LocalCurrency;

                        APInv.UserFields.Fields.Item("U_BDO_WBNo").Value = WBNo;
                        APInv.UserFields.Fields.Item("U_BDO_WBSt").Value = WBStat;
                        APInv.UserFields.Fields.Item("U_BDO_WBID").Value = WBID;

                        APInv.UserFields.Fields.Item("U_actDate").Value = WBActDate;


                        Dictionary<string, string> rsSettings = CompanyDetails.getRSSettings(out errorText);
                        if (errorText != null)
                        {
                            oForm.Freeze(false);
                            return;
                        }

                        string su = rsSettings["SU"];
                        string sp = rsSettings["SP"];
                        WayBill oWayBill = new WayBill(rsSettings["ProtocolType"]);

                        bool chek_service_user = oWayBill.chek_service_user(su, sp, out errorText);
                        if (chek_service_user == false)
                        {
                            oForm.Freeze(false);
                            errorText = BDOSResources.getTranslate("ServiceUserPasswordNotCorrect");
                            return;
                        }

                        string[] array_HEADER;
                        string[][] array_GOODS, array_SUB_WAYBILLS;
                        int returnCode = oWayBill.get_waybill(Convert.ToInt32(WBID), out array_HEADER, out array_GOODS, out array_SUB_WAYBILLS, out errorText);

                        string[][] wbTempTable = null;

                        if (wbTempLines.TryGetValue(WBNo, out wbTempTable))
                        {
                            array_GOODS = wbTempTable;
                        }

                        int rowCounter = 1;
                        //int rowIndex = 0;

                        string BPID = TIN;

                        int index = 0;
                        bool createApInvoice = true;

                        foreach (string[] goodsRow in array_GOODS)
                        {
                            string WBBarcode = goodsRow[6] == null ? "" : Regex.Replace(goodsRow[6], @"\t|\n|\r|'", "").Trim();
                            string WBItmName = goodsRow[1];
                            string WBGUntName = "";
                            string WBGUntCode = "";
                            string WBUntCdRS = goodsRow[2];
                            string Cardcode = BusinessPartners.GetCardCodeByTin(BPID, "S", out cardName);
                            if (CardCode == null)
                            {
                                oForm.Freeze(false);
                                errorText = BDOSResources.getTranslate("BPNotFound") + BDOSResources.getTranslate("BPTin") + " : " + BPID;
                                return;
                            }
                            string ItemCode = "";
                            string ItemName = "";
                            ItemCode = findItemByNameOITM(WBItmName, WBBarcode, Cardcode, out ItemName);
                            if (ItemName == null) ItemName = "";

                            SAPbobsCOM.Recordset CatalogEntry = BDO_BPCatalog.getCatalogEntryByBPBarcode(Cardcode, WBItmName, WBBarcode);

                            if (CatalogEntry != null)
                            {
                                ItemCode = CatalogEntry.Fields.Item("ItemCode").Value;
                                WBGUntCode = CatalogEntry.Fields.Item("U_BDO_UoMCod").Value;
                            } else
                            {
                                Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("PleaseChangeSearchParameter:ItIsNotPossibleToFindProductInBPCatalogUsingThisSearchingParameter"), SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                createApInvoice = false;
                                break;
                            }

                            SAPbobsCOM.Recordset oRecordsetbyRSCODE = BDO_RSUoM.getUomByRSCode(ItemCode, WBUntCdRS, out errorText);

                            if (oRecordsetbyRSCODE != null)
                            {
                                if (string.IsNullOrEmpty(WBGUntCode))
                                {
                                    WBGUntCode = oRecordsetbyRSCODE.Fields.Item("UomCode").Value;
                                }
                            }

                            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            string query = @"SELECT ""UomName"", ""UomEntry"" FROM ""OUOM"" WHERE ""UomCode"" = N'" + WBGUntCode + "'";

                            oRecordSet.DoQuery(query);

                            int UomEntry = -1;
                            if (!oRecordSet.EoF)
                            {
                                WBGUntName = oRecordSet.Fields.Item("UomName").Value;
                                UomEntry = oRecordSet.Fields.Item("UomEntry").Value;
                            }

                            //double WBQty = Convert.ToDouble(goodsRow[3], CultureInfo.InvariantCulture);
                            ////double WBPrice = Convert.ToDouble(goodsRow[4], CultureInfo.InvariantCulture);
                            //double WBSum = Convert.ToDouble(goodsRow[5], CultureInfo.InvariantCulture);
                            //------------------

                            decimal WBQty = FormsB1.cleanStringOfNonDigits(goodsRow[3]);
                            //double WBPrice = Convert.ToDouble(goodsRow[4], CultureInfo.InvariantCulture);
                            decimal WBSum = FormsB1.cleanStringOfNonDigits(goodsRow[5]);
                            decimal price = CommonFunctions.roundAmountByGeneralSettings(WBSum / WBQty, "Price");

                            SAPbobsCOM.Recordset oRecordSetIt = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            string queryIt = @"SELECT ""ManBtchNum"" FROM ""OITM"" WHERE ""ItemCode"" = '" + ItemCode + "'";

                            oRecordSetIt.DoQuery(queryIt);

                            string ManBtchNum = "N";
                            if (!oRecordSetIt.EoF)
                            {
                                ManBtchNum = oRecordSetIt.Fields.Item("ManBtchNum").Value;
                            }

                            if (ManBtchNum == "Y")
                            {
                                string BatchNumber = oMatrixGoods.GetCellSpecific("DistNumber", index + 1).Value;
                                if (BatchNumber == "")
                                {
                                    string BatchNumberFinal = Items.creatBatchNumbers(ItemCode, index, out errorText);
                                    if (errorText != null)
                                    {
                                        Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("Error") + ", " + BDOSResources.getTranslate("WaybillNumber") + ": " + WBNo + " ID:" + WBID + " " + errorText);
                                        CommonFunctions.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                                        NotToCreate = true;
                                        break;
                                    }
                                    APInv.Lines.BatchNumbers.Add();
                                    APInv.Lines.BatchNumbers.BatchNumber = BatchNumberFinal;
                                    APInv.Lines.BatchNumbers.Quantity = Convert.ToDouble(WBQty, CultureInfo.InvariantCulture);

                                    oMatrixGoods.GetCellSpecific("DistNumber", index + 1).Value = BatchNumberFinal;
                                }
                                else
                                {
                                    APInv.Lines.BatchNumbers.Add();
                                    APInv.Lines.BatchNumbers.BatchNumber = BatchNumber;
                                    APInv.Lines.BatchNumbers.Quantity = Convert.ToDouble(WBQty, CultureInfo.InvariantCulture);

                                }
                            }

                            //--------------------------------------------------
                            APInv.Lines.ItemCode = ItemCode;
                            //APInv.Lines.ItemDescription = WBItmName;

                            //Uom Entry ირკვევა UOMCODe-ის მიხედვით ცხრილში OUOM
                            APInv.Lines.UoMEntry = UomEntry;

                            APInv.Lines.WarehouseCode = whs;

                            if (!string.IsNullOrEmpty(oRecordSetWH.Fields.Item("U_BDOSPrjCod").Value))
                            {
                                APInv.Lines.ProjectCode = oRecordSetWH.Fields.Item("U_BDOSPrjCod").Value;
                            }

                            string WBPrjCode = "";
                            WBPrjCode = array_GOODS[0][12];
                            //WBPrjCode = oMatrixGoods.GetCellSpecific("WBPrjCode", index + 1).Value;

                            goodsRow[12] = WBPrjCode;

                            if (PrjCode != "")
                            {
                                APInv.Lines.ProjectCode = PrjCode;
                            }

                            if (blanketProject != "")
                            {
                                APInv.Lines.ProjectCode = blanketProject;
                            }

                            if (WBPrjCode != "")
                            {
                                APInv.Lines.ProjectCode = WBPrjCode;
                            }

                            if (WBBlankAgr != "")
                            {
                                APInv.Lines.AgreementNo = Convert.ToInt32(WBBlankAgr);
                            }

                            SAPbobsCOM.Recordset oRecordSetVat = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            string queryVat = @"SELECT ""VatStatus"",""ECVatGroup"" FROM ""OCRD"" WHERE ""OCRD"".""CardCode"" ='" + Cardcode + "'";
                            oRecordSetVat.DoQuery(queryVat);
                            string status = "";
                            string VatCode = "";
                            if (!oRecordSetVat.EoF)
                            {
                                status = oRecordSetVat.Fields.Item("VatStatus").Value;
                                VatCode = oRecordSetVat.Fields.Item("ECVatGroup").Value;
                            }
                            if (status == "Y")
                            {
                                string RSVatCode = goodsRow[8];
                                APInv.Lines.VatGroup = DetectVATByRSCode(RSVatCode, out errorText);
                            }
                            else if (status == "N")
                            {
                                APInv.Lines.VatGroup = oRecordSetVat.Fields.Item("ECVatGroup").Value;
                            }

                            if (APInv.Lines.VatGroup == null || APInv.Lines.VatGroup == "")
                            {
                                Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("TableRow") + " " + rowCounter + " " + BDOSResources.getTranslate("CannotFindVATCodeDocumentNotCreated"));
                            }

                            APInv.Lines.Quantity = Convert.ToDouble(WBQty, CultureInfo.InvariantCulture);
                            //APInv.Lines.LineTotal = WBSum;
                            Dictionary<string, string> activeDimensionsList = CommonFunctions.getActiveDimensionsList(out errorText);
                            int activeDimensions=activeDimensionsList.Count;

                            if(activeDimensions>=1) APInv.Lines.CostingCode = goodsRow[14] == null ? "" : goodsRow[14];
                            if (activeDimensions >= 2) APInv.Lines.CostingCode2 = goodsRow[15] == null ? "" : goodsRow[15];
                            if (activeDimensions >= 3) APInv.Lines.CostingCode3 = goodsRow[16] == null ? "" : goodsRow[16];
                            if (activeDimensions >= 4) APInv.Lines.CostingCode4 = goodsRow[17] == null ? "" : goodsRow[17];
                            if (activeDimensions >= 5) APInv.Lines.CostingCode5 = goodsRow[18] == null ? "" : goodsRow[18];
                            oBP = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);
                            oBP.GetByKey(CardCode);

                            pricelistnum = oBP.PriceListNum;

                            oPL = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPriceLists);
                            oPL.GetByKey(pricelistnum.ToString());

                            //if (oPL.IsGrossPrice == SAPbobsCOM.BoYesNoEnum.tYES)
                            //{
                            //    APInv.Lines.UnitPrice = WBSum / WBQty;
                            //}
                            //else
                            //{
                            //    double percent = detectVATRate( APInv.Lines.VatGroup, out errorText);
                            //    double coefficient = 1 + percent / 100;
                            //    if (RSVatCode == "0") APInv.Lines.UnitPrice = (WBSum / WBQty) / coefficient;
                            //    else if (RSVatCode == "1") APInv.Lines.UnitPrice = WBSum / WBQty;
                            //    else if (RSVatCode == "2") APInv.Lines.UnitPrice = WBSum / WBQty;
                            //}
                            APInv.Lines.Currency = Program.LocalCurrency;

                            APInv.Lines.PriceAfterVAT = Convert.ToDouble(price, CultureInfo.InvariantCulture);

                            APInv.Lines.Add();

                            index++;
                            rowCounter++;
                            //rowIndex++;
                        }

                        wbTempLines[oMatrix.GetCellSpecific("WBNo", 1).Value] = array_GOODS;

                        if (NotToCreate)
                        {
                            continue;
                        }

                        if (asDraft)
                        {
                            if (draftType == "APInvoiceDraft")
                            {
                                APInv.DocObjectCodeEx = "18";
                            }
                            else if (draftType == "GdsRcptDraft")
                            {
                                APInv.DocObjectCodeEx = "20";
                            }
                        }
                        if (createApInvoice)
                        {
                            int retvals = APInv.Add();

                            if (retvals == 0)
                            {
                                CommonFunctions.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                                string LinkedDocType = "";
                                int LinkedDocEnrty = 0;

                                if (TYPE == "Procurement")//2
                                {
                                    if (oGdsRcpt == "1")
                                    {
                                        Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("CreatedDocumentBasedOnWaybill") + " " + BDOSResources.getTranslate("GoodsRcptPO") + ", " + BDOSResources.getTranslate("WaybillNumber") + ": " + WBNo + " ID:" + WBID, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                                        if (!asDraft)
                                        {
                                            BDO_WBReceivedDocs.getGoodsReceipePOByWB(WBID, out LinkedDocType, out LinkedDocEnrty, out var linkedWhsGoodsReceipePO, out var linkedProjectGoodsReceipePO, out errorText);
                                        }
                                        else
                                        {
                                            BDO_WBReceivedDocs.GetDraftByWB(WBID, out LinkedDocType, out LinkedDocEnrty, out var linkedWhsDraft, out var linkedProjectDraft, out errorText);
                                        }
                                        oMatrix.Columns.Item("GdsRcpPO").Cells.Item(row).Specific.Value = LinkedDocEnrty;
                                    }
                                    else
                                    {
                                        Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("CreatedDocumentBasedOnWaybill") + " " + BDOSResources.getTranslate("Purchase") + ", " + BDOSResources.getTranslate("WaybillNumber") + ": " + WBNo + " ID:" + WBID, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                                        if (!asDraft)
                                        {
                                            BDO_WBReceivedDocs.getInvoiceByWB(WBID, out LinkedDocType, out LinkedDocEnrty, out var linkedWhsInvoice, out var linkedProjectInvoice, out errorText);
                                        }
                                        else
                                        {
                                            BDO_WBReceivedDocs.GetDraftByWB(WBID, out LinkedDocType, out LinkedDocEnrty, out var linkedWhsDraft, out var linkedProjectDraft, out errorText);
                                        }
                                        oMatrix.Columns.Item("APInvoice").Cells.Item(row).Specific.Value = LinkedDocEnrty;
                                    }
                                }

                                if (TYPE == "Return")//1
                                {
                                    Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("CreatedDocumentBasedOnWaybill") + " " + BDOSResources.getTranslate("Return") + ", " + BDOSResources.getTranslate("WaybillNumber") + ": " + WBNo + " ID:" + WBID, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                                    BDO_WBReceivedDocs.getMemoByWB(WBID, out LinkedDocType, out LinkedDocEnrty, out var linkedWhsMemo, out var linkedProjectMemo, out errorText);
                                    oMatrix.Columns.Item("CredMemo").Cells.Item(row).Specific.Value = LinkedDocEnrty;
                                }

                                oMatrix.Columns.Item("WBCheckbox").Cells.Item(row).Specific.Checked = false;

                            }
                            else
                            {
                                CommonFunctions.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                                int errCode;
                                string errMSG;

                                Program.oCompany.GetLastError(out errCode, out errMSG);
                                Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("Error") + ", " + BDOSResources.getTranslate("WaybillNumber") + ": " + WBNo + " ID:" + WBID + " " + errMSG);

                                int ind = 0;
                                foreach (string[] goodsRow in array_GOODS)
                                {
                                    oMatrixGoods.GetCellSpecific("DistNumber", ind + 1).Value = "";
                                    ind++;
                                }
                            }
                        }
                    }
                    else
                    {
                        string errMSG = "UnableToCreateDocumentOnCanceledInvoice";
                        Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("Error") + " : " + BDOSResources.getTranslate(errMSG));

                    }
                }
            }

            //updateForm( oForm, out errorText);
            oForm.Update();
            oForm.Freeze(false);
        }
        public static void updateForm(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("WBMatrix").Specific));

            SAPbouiCOM.DataTable oDataTable;
            oDataTable = oForm.DataSources.DataTables.Item("WBTable");
            oDataTable.Rows.Clear();

            int rowCounter = 1;
            int rowIndex = 0;

            WayBill oWayBill;
            Dictionary<string, Dictionary<string, string>> waybills_map = getDataFromRS(oForm, out oWayBill, out errorText);

            if (errorText != null)
            {
                Program.uiApp.StatusBar.SetSystemMessage(errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

            foreach (var map_record in waybills_map)
            {
                string WBID = map_record.Key;

                Dictionary<string, string> Waybill_Header = map_record.Value;
                string WBNo = Waybill_Header["WAYBILL_NUMBER"];
                string WBStat = Waybill_Header["STATUS"];
                string WBSupName = Waybill_Header["SELLER_NAME"];
                string WBSupTIN = Waybill_Header["SELLER_TIN"];
                string WBActDate = Waybill_Header["ACTIVATE_DATE"].Replace("T", " ");
                string WBStartDat = Waybill_Header["BEGIN_DATE"].Replace("T", " ");
                string WBStartAdd = Waybill_Header["START_ADDRESS"];
                string WBEndAdd = Waybill_Header["END_ADDRESS"];
                string WBtype = Waybill_Header["TYPE"];
                string WBveh = Waybill_Header["CAR_NUMBER"];
                double WBSUM = Convert.ToDouble(Waybill_Header["FULL_AMOUNT"], CultureInfo.InvariantCulture);
                string TYPE = Waybill_Header["TYPE"];
                string WBCOM = Waybill_Header["WAYBILL_COMMENT"];
                string LinkedDocType = "";

                int LinkedDocEntryInvoice = 0;
                BDO_WBReceivedDocs.getInvoiceByWB(WBID, out LinkedDocType, out LinkedDocEntryInvoice, out var linkedWhsInvoice, out var linkedProjectInvoice, out errorText);

                int LinkedDocEntryGoodsReceipePO = 0;
                BDO_WBReceivedDocs.getGoodsReceipePOByWB(WBID, out LinkedDocType, out LinkedDocEntryGoodsReceipePO, out var linkedWhsGoodsReceipePO, out var linkedProjectGoodsReceipePO, out errorText);

                BDO_WBReceivedDocs.GetDraftByWB(WBID, out var linkedDocTypeDraft, out var linkedDocEntryDraft, out var linkedWhsDraft, out var linkedProjectDraft, out errorText);

                int LinkedDocEntryMemo = 0;
                BDO_WBReceivedDocs.getMemoByWB(WBID, out LinkedDocType, out LinkedDocEntryMemo, out var linkedWhsMemo, out var linkedProjectMemo, out errorText);

                string attachFilter = oForm.DataSources.UserDataSources.Item("Attach").Value;

                if (LinkedDocEntryInvoice != 0 || LinkedDocEntryMemo != 0 || LinkedDocEntryGoodsReceipePO != 0 || linkedDocEntryDraft != 0)
                {
                    if (attachFilter == "1") continue; //მიუბმელი გვინდა
                }
                else
                {
                    if (attachFilter == "2") continue; //მიბმულები გვინდა
                }

                DateTime ActDt = new DateTime(1, 1, 1);

                if (DateTime.TryParseExact(WBActDate, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture, DateTimeStyles.None, out ActDt) == false)
                {
                    ActDt = new DateTime(1, 1, 1);
                }


                DateTime StartDt = new DateTime(1, 1, 1);

                if (DateTime.TryParseExact(WBStartDat, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture, DateTimeStyles.None, out StartDt) == false)
                {
                    StartDt = new DateTime(1, 1, 1);
                }

                oDataTable.Rows.Add();
                oDataTable.SetValue(2, rowIndex, WBID);
                oDataTable.SetValue(0, rowIndex, rowCounter);
                oDataTable.SetValue(1, rowIndex, WBNo);
                oDataTable.SetValue(3, rowIndex, BDO_WBReceivedDocs.detectWBStatus(WBStat));
                oDataTable.SetValue(4, rowIndex, WBSupName);
                oDataTable.SetValue(5, rowIndex, ActDt);
                oDataTable.SetValue(6, rowIndex, WBStartAdd);
                oDataTable.SetValue(7, rowIndex, WBEndAdd);
                oDataTable.SetValue(8, rowIndex, WBSUM);
                oDataTable.SetValue(9, rowIndex, WBSupTIN);
                oDataTable.SetValue(10, rowIndex, "0");
                oDataTable.SetValue(14, rowIndex, TYPE);
                oDataTable.SetValue("WBStartDat", rowIndex, StartDt);



                if (LinkedDocEntryInvoice != 0)
                {
                    oDataTable.SetValue(11, rowIndex, LinkedDocEntryInvoice.ToString());
                    oDataTable.SetValue(18, rowIndex, linkedWhsInvoice);
                    oDataTable.SetValue(19, rowIndex, linkedProjectInvoice);
                }

                if (LinkedDocEntryGoodsReceipePO != 0)
                {
                    oDataTable.SetValue(12, rowIndex, LinkedDocEntryGoodsReceipePO.ToString());
                    oDataTable.SetValue(18, rowIndex, linkedWhsGoodsReceipePO);
                    oDataTable.SetValue(19, rowIndex, linkedProjectGoodsReceipePO);
                }

                if (linkedDocEntryDraft != 0)
                {
                    if (linkedDocTypeDraft == "APInvoiceDraft")
                    {
                        oDataTable.SetValue(11, rowIndex, linkedDocEntryDraft.ToString());
                    }
                    else if (linkedDocTypeDraft == "GdsRcptDraft")
                    {
                        oDataTable.SetValue(12, rowIndex, linkedDocEntryDraft.ToString());
                    }

                    oDataTable.SetValue(18, rowIndex, linkedWhsDraft);
                    oDataTable.SetValue(19, rowIndex, linkedProjectDraft);
                }

                if (LinkedDocEntryMemo != 0)
                {
                    oDataTable.SetValue(13, rowIndex, LinkedDocEntryMemo.ToString());
                    oDataTable.SetValue(18, rowIndex, linkedWhsMemo);
                    oDataTable.SetValue(19, rowIndex, linkedProjectMemo);
                }

                oDataTable.SetValue(16, rowIndex, WBCOM);

                if (!string.IsNullOrEmpty(WBEndAdd))
                {
                    string whsCode = BDOSWarehouseAddresses.GetWhsByAddress(WBEndAdd);
                    if (!string.IsNullOrEmpty(whsCode))
                    {
                        oDataTable.SetValue(18, rowIndex, whsCode);
                    }
                }


                rowCounter++;
                rowIndex++;

            }

            oForm.Freeze(true);
            oMatrix.Clear();
            oMatrix.LoadFromDataSource();
            oMatrix.AutoResizeColumns();
            oForm.Freeze(false);

            oForm.Freeze(true);
            fillWBGoods(oForm, 1, false, out errorText);
            oForm.Freeze(false);
        }
        public static void addRow(out string errorText)
        {
            errorText = null;

            SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm("60004", Program.currentFormCount);
            //oForm.Select();

            oForm.Freeze(true);

            SAPbouiCOM.Matrix oMatrixGoods = (SAPbouiCOM.Matrix)oForm.Items.Item("WBGdMatrix").Specific;

            if (WBGdMatrixRow > 0)
            {

                oMatrixGoods.AddRow();

                decimal price = FormsB1.cleanStringOfNonDigits(oMatrixGoods.GetCellSpecific("WBPrice", WBGdMatrixRow).Value);

                decimal oldQty = FormsB1.cleanStringOfNonDigits(oMatrixGoods.GetCellSpecific("WBQty", WBGdMatrixRow).Value);
                decimal oldSum = oldQty * price;

                decimal newQty = WBGdMatrixNewQty;
                decimal newSum = newQty * price;

                int oldRow = WBGdMatrixRow;

                //string newQtySt = oForm.Items.Item("newQty").Specific.Value;

                oMatrixGoods.Columns.Item("#").Cells.Item(oMatrixGoods.RowCount).Specific.Value = oMatrixGoods.RowCount;
                oMatrixGoods.Columns.Item("WBNo").Cells.Item(oMatrixGoods.RowCount).Specific.Value = oMatrixGoods.GetCellSpecific("WBNo", oldRow).Value;
                oMatrixGoods.Columns.Item("WBBarcode").Cells.Item(oMatrixGoods.RowCount).Specific.Value = oMatrixGoods.GetCellSpecific("WBBarcode", oldRow).Value;
                oMatrixGoods.Columns.Item("ItemCode").Cells.Item(oMatrixGoods.RowCount).Specific.Value = oMatrixGoods.GetCellSpecific("ItemCode", oldRow).Value;
                oMatrixGoods.Columns.Item("ItemName").Cells.Item(oMatrixGoods.RowCount).Specific.Value = oMatrixGoods.GetCellSpecific("ItemName", oldRow).Value;
                oMatrixGoods.Columns.Item("DistNumber").Cells.Item(oMatrixGoods.RowCount).Specific.Value = oMatrixGoods.GetCellSpecific("DistNumber", oldRow).Value;
                oMatrixGoods.Columns.Item("WBItmName").Cells.Item(oMatrixGoods.RowCount).Specific.Value = oMatrixGoods.GetCellSpecific("WBItmName", oldRow).Value;
                oMatrixGoods.Columns.Item("WBUntCode").Cells.Item(oMatrixGoods.RowCount).Specific.Value = oMatrixGoods.GetCellSpecific("WBUntCode", oldRow).Value;
                //oMatrixGoods.Columns.Item("WBUntName").Cells.Item(oMatrixGoods.RowCount).Specific.Value = oMatrixGoods.GetCellSpecific("WBUntName", oldRow).Value;
                oMatrixGoods.Columns.Item("WbUntNmRS").Cells.Item(oMatrixGoods.RowCount).Specific.Value = oMatrixGoods.GetCellSpecific("WbUntNmRS", oldRow).Value;

                oMatrixGoods.Columns.Item("WBQty").Cells.Item(oMatrixGoods.RowCount).Specific.Value = FormsB1.ConvertDecimalToString(newQty);
                //oMatrixGoods.GetCellSpecific("WBQty", WBGdMatrixRow).Value;
                oMatrixGoods.Columns.Item("WBPrice").Cells.Item(oMatrixGoods.RowCount).Specific.Value = oMatrixGoods.GetCellSpecific("WBPrice", oldRow).Value;
                oMatrixGoods.Columns.Item("WBSum").Cells.Item(oMatrixGoods.RowCount).Specific.Value = FormsB1.ConvertDecimalToString(newSum);
                //oMatrixGoods.GetCellSpecific("WBSum", WBGdMatrixRow).Value;

                oMatrixGoods.Columns.Item("WBUntCdRS").Cells.Item(oMatrixGoods.RowCount).Specific.Value = oMatrixGoods.GetCellSpecific("WBUntCdRS", oldRow).Value;
                oMatrixGoods.Columns.Item("RSVatCode").Cells.Item(oMatrixGoods.RowCount).Specific.Value = oMatrixGoods.GetCellSpecific("RSVatCode", oldRow).Value;
                oMatrixGoods.Columns.Item("WBPrjCode").Cells.Item(oMatrixGoods.RowCount).Specific.Value = oMatrixGoods.GetCellSpecific("WBPrjCode", oldRow).Value;

                //Gadasamowmebelia FormsB1.ConvertDecimalToString() da cleanStringOfNonDigits()
                //Roca formidan vigebt ricxvs an formashi vcert ricxvs unda itvaliswinebdes Saerto awyobebis Gamyofs da Atasebis gamyofs!!
                oMatrixGoods.Columns.Item("WBQty").Cells.Item(oldRow).Specific.Value = FormsB1.ConvertDecimalToString(oldQty - newQty);
                oMatrixGoods.Columns.Item("WBSum").Cells.Item(oldRow).Specific.Value = FormsB1.ConvertDecimalToString(oldSum - newSum);

                WBGdMatrixMaxQty = oldQty - newQty;

                if (WBGdMatrixRow != oMatrixGoods.RowCount)
                {

                    {
                        oForm.Freeze(false);
                        for (int i = 1; i <= oMatrixGoods.RowCount; i++)
                        {
                            oMatrixGoods.CommonSetting.SetRowBackColor(i, FormsB1.getLongIntRGB(231, 231, 231));
                        }

                        try
                        {
                            oMatrixGoods.CommonSetting.SetRowBackColor(oMatrixGoods.RowCount, FormsB1.getLongIntRGB(255, 255, 153));
                            WBGdMatrixMaxQty = FormsB1.cleanStringOfNonDigits(oMatrixGoods.Columns.Item("WBQty").Cells.Item(WBGdMatrixRow).Specific.Value);
                        }
                        catch
                        {
                        }
                        oForm.Freeze(true);
                    }
                }

            }

            oForm.Freeze(false);

        }
        public static Dictionary<string, Dictionary<string, string>> getDataFromRS(SAPbouiCOM.Form oForm, out WayBill oWayBill, out string errorText)
        {
            errorText = null;
            oWayBill = null;

            string startDateStr = oForm.DataSources.UserDataSources.Item("StartDate").ValueEx;
            DateTime startDate = FormsB1.DateFormats(startDateStr, "yyyyMMdd") == new DateTime() ? DateTime.Today : FormsB1.DateFormats(startDateStr, "yyyyMMdd");

            string endDateStr = oForm.DataSources.UserDataSources.Item("EndDate").ValueEx;
            DateTime endDate = FormsB1.DateFormats(endDateStr, "yyyyMMdd") == new DateTime() ? DateTime.Now : FormsB1.DateFormats(endDateStr, "yyyyMMdd").AddDays(1).AddSeconds(-1);

            Dictionary<string, Dictionary<string, string>> waybills_map = new Dictionary<string, Dictionary<string, string>>();

            //საქონლის ცხრილი
            Dictionary<string, string> rsSettings = CompanyDetails.getRSSettings(out errorText);
            if (errorText != null)
            {
                return waybills_map;
            }

            string su = rsSettings["SU"];
            string sp = rsSettings["SP"];
            oWayBill = new WayBill(rsSettings["ProtocolType"]);

            bool chek_service_user = oWayBill.chek_service_user(su, sp, out errorText);
            if (chek_service_user == false)
            {
                errorText = BDOSResources.getTranslate("ServiceUserPasswordNotCorrect");
                return waybills_map;
            }

            string typefilter = oForm.DataSources.UserDataSources.Item("WaybType").Value;

            string itypes = "2,3,4,5,6";

            if (typefilter == "1")
            {
                itypes = "5";
            }
            else if (typefilter == "2")
            {
                itypes = "2,3,4,6";
            }

            string statuses = ",1,2,";
            string car_number = oForm.DataSources.UserDataSources.Item("CarNo").Value;
            DateTime begin_date_s = startDate;
            DateTime begin_date_e = endDate;
            DateTime create_date_s = startDate;
            DateTime create_date_e = endDate;
            string driver_tin = null;
            DateTime delivery_date_s = startDate;
            DateTime delivery_date_e = endDate;
            decimal full_amount = 0;
            string waybill_number = oForm.DataSources.UserDataSources.Item("WBNo").Value;
            DateTime close_date_s = startDate;
            DateTime close_date_e = endDate;
            string s_user_id = "";
            string comment = null;
            string seller_id = oForm.DataSources.UserDataSources.Item("WBSuplNo").Value;
            string startAddress = oForm.DataSources.UserDataSources.Item("StartAdd").Value;
            string endAddress = oForm.DataSources.UserDataSources.Item("EndAdd").Value;

            DateTime startDateParam = new DateTime();
            DateTime endDateParam = new DateTime();
            startDateParam = startDate;

            while (startDateParam < endDate)
            {
                endDateParam = startDateParam.AddDays(2);

                if (endDateParam > endDate)
                {
                    endDateParam = endDate;
                }

                string stAdd = oForm.DataSources.UserDataSources.Item("stAddEm").ValueEx;
                string enAdd = oForm.DataSources.UserDataSources.Item("endAddEm").ValueEx;
                string startAd = oForm.DataSources.UserDataSources.Item("StartAdd").ValueEx;
                string endAd = oForm.DataSources.UserDataSources.Item("EndAdd").ValueEx;
                if (stAdd == "Y")
                {
                    startAd = "blank";
                }
                if (enAdd == "Y")
                {
                    endAd = "blank";
                }

                Dictionary<string, Dictionary<string, string>> waybills_map_part = oWayBill.get_buyer_waybills(itypes, seller_id, statuses, car_number, startDateParam, endDateParam, startDateParam, endDateParam, driver_tin, startDateParam, endDateParam, full_amount, waybill_number, startDateParam, endDateParam, s_user_id, comment, startAd, endAd, out errorText);
                foreach (KeyValuePair<string, Dictionary<string, string>> keyvalue in waybills_map_part)
                {
                    try
                    {
                        Dictionary<string, string> Waybill_Header = keyvalue.Value;

                        if (AddressesMatch(startAddress, endAddress, keyvalue))
                        {
                            string WBID = keyvalue.Key;
                            string WBStat = Waybill_Header["STATUS"];
                            string WBActDate = Waybill_Header["ACTIVATE_DATE"].Replace("T", " ");
                            string WBActDat = Waybill_Header["ACTIVATE_DATE"];
                            string WBTYPE = Waybill_Header["TYPE"];
                            if (WBTYPE == "5")
                            {
                                Waybill_Header["TYPE"] = "Return";
                            }
                            else
                            {
                                Waybill_Header["TYPE"] = "Procurement";
                            }
                            double WBSUM = Convert.ToDouble(Waybill_Header["FULL_AMOUNT"], CultureInfo.InvariantCulture);
                            SAPbouiCOM.EditText wBIDD = (SAPbouiCOM.EditText)(oForm.Items.Item("wayBID").Specific);
                            SAPbouiCOM.EditText actDate = (SAPbouiCOM.EditText)(oForm.Items.Item("ActDate").Specific);
                            SAPbouiCOM.EditText amouNT = (SAPbouiCOM.EditText)(oForm.Items.Item("AmountE").Specific);
                            SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)(oForm.Items.Item("Status").Specific);
                            string chosen = oCombo.Value;
                            string val = "";
                            if (chosen == "1")
                            {
                                val = "1";
                            }
                            else if (chosen == "2")
                            {

                                val = "2";
                            }
                            string toCut = WBActDate;
                            string actDateDict = toCut.Substring(0, 4) + toCut.Substring(5, 2) + toCut.Substring(8, 2);
                            if ((wBIDD.Value == "" || WBID == wBIDD.Value)
                               && ((val != "1" && val != "2") || val == WBStat) &&
                                ((amouNT.Value == "") || (WBSUM.ToString() == amouNT.Value))
                                && (actDate.Value == "" || actDate.Value == actDateDict))
                            {


                                waybills_map.Add(keyvalue.Key, keyvalue.Value);
                            }
                        }
                    }
                    catch
                    {
                    }
                }

                startDateParam = endDateParam;

            }

            return waybills_map;
        }
        public static bool AddressesMatch(string reqStartAddress, string reqEndAddress, KeyValuePair<string, Dictionary<string, string>> keyvalue)
        {
            reqStartAddress = reqStartAddress.Replace("*", "");
            reqEndAddress = reqEndAddress.Replace("*", "");

            keyvalue.Value.TryGetValue("START_ADDRESS", out string rsStartAddress);
            keyvalue.Value.TryGetValue("END_ADDRESS", out string rsEndAddress);

            return StringContainsOtherString(rsStartAddress, reqStartAddress) && StringContainsOtherString(rsEndAddress, reqEndAddress);
        }
        public static bool StringContainsOtherString(string s1, string s2)
        {
            bool match = true;

            if (String.IsNullOrEmpty(s2))
            {
                match = true;
            }
            else
            {
                if (String.IsNullOrEmpty(s2))
                {
                    match = false;
                }
                else
                {
                    match = s1.Contains(s2);
                }
            }

            return match;
        }
        public static void createFormNewRow(SAPbouiCOM.Form oDocForm, out string errorText)
        {
            errorText = null;

            int left = 558 + 500;
            int Top = 200 + 300;

            //ფორმის აუცილებელი თვისებები
            Dictionary<string, object> formProperties = new Dictionary<string, object>();
            formProperties.Add("UniqueID", "BDO_WaybillsReceivedNewRowForm");
            formProperties.Add("BorderStyle", SAPbouiCOM.BoFormBorderStyle.fbs_Sizable);
            formProperties.Add("Title", BDOSResources.getTranslate("NewRowQuantity"));
            formProperties.Add("Left", left);
            formProperties.Add("Width", 200);
            formProperties.Add("Top", Top);
            formProperties.Add("Height", 10);
            formProperties.Add("Modality", SAPbouiCOM.BoFormModality.fm_Modal);

            SAPbouiCOM.Form oForm;
            bool newForm;
            bool formExist = FormsB1.createForm(formProperties, out oForm, out newForm, out errorText);

            if (formExist == true)
            {
                if (newForm == true)
                {
                    //ფორმის ელემენტების თვისებები
                    Dictionary<string, object> formItems = null;

                    Top = 1;
                    left = 6;

                    //formItems = new Dictionary<string, object>();
                    //string itemName = "WBNoSt";
                    //formItems.Add("Size", 20);
                    //formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    //formItems.Add("Left", left);
                    // formItems.Add("Width", 120);
                    //formItems.Add("Top", Top);
                    //formItems.Add("Caption", BDOSResources.getTranslate("WaybillNumber"));
                    //formItems.Add("UID", itemName);

                    //FormsB1.createFormItem(oForm, formItems, out errorText);
                    //if (errorText != null)
                    //{
                    //    return;
                    //}

                    //left = left + 128+ 10;

                    formItems = new Dictionary<string, object>();
                    string itemName = "newQty";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_QUANTITY);
                    formItems.Add("Length", 30);
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    Top = Top + 19 + 5;
                    left = 6;

                    itemName = "1";
                    formItems = new Dictionary<string, object>();
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 120);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("Update"));

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }
                }
                oForm.Visible = true;
                //oForm.Select();
            }

            GC.Collect();
        }
        public static void createForm(SAPbouiCOM.Form oDocForm, out string errorText)
        {
            errorText = null;

            //ფორმის აუცილებელი თვისებები
            Dictionary<string, object> formProperties = new Dictionary<string, object>();
            formProperties.Add("UniqueID", "BDO_WaybillsReceivedForm");
            formProperties.Add("BorderStyle", SAPbouiCOM.BoFormBorderStyle.fbs_Sizable);
            formProperties.Add("Title", BDOSResources.getTranslate("WaybillReceived"));
            formProperties.Add("Left", 558);
            formProperties.Add("Width", 1300);
            formProperties.Add("Top", 200);
            formProperties.Add("Height", 800);

            SAPbouiCOM.Form oForm;
            bool newForm;
            bool formExist = FormsB1.createForm(formProperties, out oForm, out newForm, out errorText);

            if (formExist)
            {
                if (newForm)
                {
                    oForm.DataSources.UserDataSources.Add("DocEntry", SAPbouiCOM.BoDataType.dt_LONG_NUMBER);
                    oForm.DataSources.UserDataSources.Add("DocType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                    oForm.DataSources.UserDataSources.Add("CurrWBNo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                    oForm.DataSources.UserDataSources.Add("CurrWBID", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                    oForm.DataSources.UserDataSources.Add("CurrWBSt", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                    oForm.DataSources.UserDataSources.Add("CurrDate", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                    oForm.DataSources.UserDataSources.Add("CurrRow", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                    oForm.DataSources.UserDataSources.Add("CurrBP", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);

                    //ფორმის ელემენტების თვისებები
                    Dictionary<string, object> formItems = null;

                    string itemName = "";
                    int left = 6;
                    int Top = 5;

                    List<string> ValidValues = new List<string>();
                    //რიგი 1
                    //თარიღები
                    formItems = new Dictionary<string, object>();
                    itemName = "dateFrom";
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 120);
                    formItems.Add("Top", Top);
                    formItems.Add("Caption", BDOSResources.getTranslate("StartDate"));
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left = left + 128 + 10;

                    string startOfMonthStr = DateTime.Today.ToString("yyyyMMdd");
                    formItems = new Dictionary<string, object>();
                    itemName = "StartDate";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_DATE);
                    formItems.Add("Length", 1);
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("ValueEx", startOfMonthStr);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }


                    left = left + 100 + 10;

                    itemName = "dateTo";
                    formItems = new Dictionary<string, object>();
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 120);
                    formItems.Add("Top", Top);
                    formItems.Add("Caption", BDOSResources.getTranslate("EndDate"));
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left = left + 128 + 10;

                    string endOfMonthStr = DateTime.Today.ToString("yyyyMMdd");
                    formItems = new Dictionary<string, object>();
                    itemName = "EndDate";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_DATE);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("Length", 1);
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("ValueEx", endOfMonthStr);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left = left + 100 + 10;

                    formItems = new Dictionary<string, object>();
                    itemName = "StartAddSt";
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 130);
                    formItems.Add("Top", Top);
                    formItems.Add("Caption", BDOSResources.getTranslate("StartAddress"));
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left = left + 128 + 10;

                    formItems = new Dictionary<string, object>();
                    itemName = "StartAdd";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("Length", 100);
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }
                    int leftCheckBox = left + 100;
                    formItems = new Dictionary<string, object>();
                    itemName = "stAddEm";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("Length", 1);
                    formItems.Add("Left", leftCheckBox);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 14);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("OnlyBlank"));
                    formItems.Add("ValOff", "N");
                    formItems.Add("ValOn", "Y");
                    formItems.Add("LinkTo", "StartAdd");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left = left + 110 + 10;

                    formItems = new Dictionary<string, object>();
                    itemName = "CarNoSt";
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left + 90);
                    formItems.Add("Width", 160);
                    formItems.Add("Top", Top);
                    formItems.Add("Caption", BDOSResources.getTranslate("VehicleNumber"));
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left = left + 160 + 10;

                    formItems = new Dictionary<string, object>();
                    itemName = "CarNo";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("Length", 30);
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    //     UPDATE
                    itemName = "10";
                    formItems = new Dictionary<string, object>();
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left + 100 + 5);
                    formItems.Add("Width", 65);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("RSUpdate"));

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    //რიგი 2
                    Top = Top + 20;
                    left = 6;
                    formItems = new Dictionary<string, object>();
                    itemName = "actDateS";
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 120);
                    formItems.Add("Top", Top);
                    formItems.Add("Caption", BDOSResources.getTranslate("ActivateDate"));
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left = left + 128 + 10;

                    string ActOfMonthStr = DateTime.Today.ToString("yyyyMMdd");
                    formItems = new Dictionary<string, object>();
                    itemName = "ActDate";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_DATE);
                    formItems.Add("Length", 1);
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    //   formItems.Add("ValueEx", ActOfMonthStr);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left = left + 100 + 10;

                    formItems = new Dictionary<string, object>();
                    itemName = "WBSuplNoSt";
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 120);
                    formItems.Add("Top", Top);
                    formItems.Add("Caption", BDOSResources.getTranslate("BPTin"));
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left = left + 128 + 10;

                    formItems = new Dictionary<string, object>();
                    itemName = "WBSuplNo";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("Length", 30);
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }


                    left = left + 100 + 10;
                    formItems = new Dictionary<string, object>();
                    itemName = "EndAddSt";
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 120);
                    formItems.Add("Top", Top);
                    formItems.Add("Caption", BDOSResources.getTranslate("EndAddress"));
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left = left + 128 + 10;

                    formItems = new Dictionary<string, object>();
                    itemName = "EndAdd";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("Length", 100);
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }


                    formItems = new Dictionary<string, object>();
                    itemName = "endAddEm";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("Length", 1);
                    formItems.Add("Left", leftCheckBox);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 14);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("OnlyBlank"));
                    formItems.Add("ValOff", "N");
                    formItems.Add("ValOn", "Y");
                    formItems.Add("LinkTo", "EndAdd");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }




                    left = oForm.Items.Item("10").Left;

                    formItems = new Dictionary<string, object>();
                    itemName = "GdsRcpt";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_OPTION_BUTTON);
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("Length", 1);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 160);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 14);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("CreateGoodsReceiptPO"));
                    formItems.Add("ValOff", "N");
                    formItems.Add("ValOn", "Y");
                    formItems.Add("DisplayDesc", true);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "APInv";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_OPTION_BUTTON);
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("Length", 1);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 160);
                    formItems.Add("Top", Top + 14);
                    formItems.Add("Height", 14);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("CreateAPInvoice"));
                    formItems.Add("GroupWith", "GdsRcpt");
                    formItems.Add("ValOn", "Y");
                    formItems.Add("ValOff", "N");
                    formItems.Add("Selected", true);
                    formItems.Add("DisplayDesc", true);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "AsDraft";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("Length", 1);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 160);
                    formItems.Add("Top", Top + 28);
                    formItems.Add("Height", 14);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("CreateAsDraft"));
                    formItems.Add("ValOn", "Y");
                    formItems.Add("ValOff", "N");
                    formItems.Add("DisplayDesc", true);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    //რიგი3
                    Top = Top + 20;
                    left = 6;
                    formItems = new Dictionary<string, object>();
                    itemName = "WBNoSt";
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 120);
                    formItems.Add("Top", Top);
                    formItems.Add("Caption", BDOSResources.getTranslate("WaybillNumber"));
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left = left + 128 + 10;

                    formItems = new Dictionary<string, object>();
                    itemName = "WBNo";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("Length", 30);
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left = left + 100 + 10;

                    formItems = new Dictionary<string, object>();
                    itemName = "StatusS";
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 65);
                    formItems.Add("Top", Top);
                    formItems.Add("Caption", BDOSResources.getTranslate("Status"));
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left = left + 128 + 10;
                    ValidValues.Clear();
                    ValidValues.Add(BDOSResources.getTranslate("WithoutFilter"));
                    ValidValues.Add(BDOSResources.getTranslate("Active")); //0
                    ValidValues.Add(BDOSResources.getTranslate("Finished"));//1

                    formItems = new Dictionary<string, object>();
                    itemName = "Status";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("Length", 30);
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("ValidValues", ValidValues);
                    formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
                    formItems.Add("DisplayDesc", true);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }
                    left = left + 100 + 10;
                    formItems = new Dictionary<string, object>();
                    itemName = "AttachST";
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 120);
                    formItems.Add("Top", Top);
                    formItems.Add("Caption", BDOSResources.getTranslate("LinkToDocument"));
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left = left + 128 + 10;

                    ValidValues = new List<string>();
                    ValidValues.Add(BDOSResources.getTranslate("WithoutFilter"));
                    ValidValues.Add(BDOSResources.getTranslate("NotLinked"));
                    ValidValues.Add(BDOSResources.getTranslate("Linked"));

                    formItems = new Dictionary<string, object>();
                    itemName = "Attach";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("Length", 30);
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("ValidValues", ValidValues);
                    formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
                    formItems.Add("DisplayDesc", true);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    //left = left + 128+ 10;
                    //left = left + 128+ 10;


                    //რიგი4
                    Top = Top + 20;
                    left = 6;

                    formItems = new Dictionary<string, object>();
                    itemName = "wayBIDS";
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 120);
                    formItems.Add("Top", Top);
                    formItems.Add("Caption", BDOSResources.getTranslate("waybillID"));
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left = left + 128 + 10;

                    formItems = new Dictionary<string, object>();
                    itemName = "wayBID";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("Length", 30);
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left = left + 100 + 10;
                    formItems = new Dictionary<string, object>();
                    itemName = "WaybTypeS";
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", Top);
                    formItems.Add("Caption", BDOSResources.getTranslate("WaybillType"));
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left = left + 128 + 10;

                    ValidValues = new List<string>();
                    ValidValues.Add(BDOSResources.getTranslate("WithoutFilter")); //0
                    ValidValues.Add(BDOSResources.getTranslate("Return"));//1
                    ValidValues.Add(BDOSResources.getTranslate("Purchase"));//2

                    formItems = new Dictionary<string, object>();
                    itemName = "WaybType";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("Length", 30);
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("ValidValues", ValidValues);
                    formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
                    formItems.Add("DisplayDesc", true);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left = left + 100 + 10;

                    formItems = new Dictionary<string, object>();
                    itemName = "AmountS";
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 120);
                    formItems.Add("Top", Top);
                    formItems.Add("Caption", BDOSResources.getTranslate("Amount"));
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left = left + 128 + 10;

                    formItems = new Dictionary<string, object>();
                    itemName = "AmountE";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("Length", 30);
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left = oForm.Items.Item("10").Left;

                    //საწყობი

                    formItems = new Dictionary<string, object>();
                    itemName = "WhsST";
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 120);
                    formItems.Add("Top", Top + 2);
                    formItems.Add("Caption", BDOSResources.getTranslate("Warehouse"));
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left = left + 70 + 10;
                    var objectType = "64"; //Warehouse
                    const string uniqueID_lf_BusinessPartnerCFL = "Whs_CFL";
                    FormsB1.addChooseFromList(oForm, false, objectType, uniqueID_lf_BusinessPartnerCFL);

                    formItems = new Dictionary<string, object>();
                    itemName = "Whs";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("Length", 30);
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", Top + 2);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("ChooseFromListUID", uniqueID_lf_BusinessPartnerCFL);
                    formItems.Add("ChooseFromListAlias", "WhsCode");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    formItems = new Dictionary<string, object>();
                    itemName = "BDOSWhsLB"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    formItems.Add("Left", left - 20);
                    formItems.Add("Top", Top + 2);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("LinkTo", "Whs");
                    formItems.Add("LinkedObjectType", objectType);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }


                    //პროექტი

                    left = oForm.Items.Item("10").Left;

                    formItems = new Dictionary<string, object>();
                    itemName = "PrjCodeST";
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 120);
                    formItems.Add("Top", Top + 22);
                    formItems.Add("Caption", BDOSResources.getTranslate("Project"));
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    left = left + 70 + 10;
                    objectType = "63"; //Project
                    const string uniqueID_lf_Project = "Project_CFLA";
                    FormsB1.addChooseFromList(oForm, false, objectType, uniqueID_lf_Project);


                    formItems = new Dictionary<string, object>();
                    itemName = "PrjCode";
                    formItems.Add("isDataSource", true);
                    formItems.Add("DataSource", "UserDataSources");
                    formItems.Add("DataType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                    formItems.Add("Length", 30);
                    formItems.Add("Size", 20);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    formItems.Add("TableName", "");
                    formItems.Add("Alias", itemName);
                    formItems.Add("Bound", true);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 100);
                    formItems.Add("Top", Top + 22);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("ChooseFromListUID", uniqueID_lf_Project);
                    formItems.Add("ChooseFromListAlias", "PrjCode");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }
                    formItems = new Dictionary<string, object>();
                    itemName = "BDOSPrjLB"; //10 characters
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    formItems.Add("Left", left - 20);
                    formItems.Add("Top", Top + 22);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("LinkTo", "PrjCode");
                    formItems.Add("LinkedObjectType", objectType);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    //პროექტი



                    //რიგი5
                    Top = Top + 42;
                    left = 6;

                    //ზედნადებების ცხრილი
                    itemName = "WBMatrix";
                    formItems = new Dictionary<string, object>();
                    formItems.Add("isDataSource", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
                    //formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_GRID);
                    formItems.Add("Left", left);
                    formItems.Add("Width", oForm.Width - 20);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 40);
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    //რიგი6-0
                    Top = Top + 200;
                    left = 6;

                    itemName = "AddRow";
                    formItems = new Dictionary<string, object>();
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", left);
                    formItems.Add("Width", 85);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("AddRow"));

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    //რიგი6
                    Top = Top + 25;
                    left = 6;

                    //საქონლის ცხრილი
                    itemName = "WBGdMatrix";
                    formItems = new Dictionary<string, object>();
                    formItems.Add("isDataSource", true);
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
                    formItems.Add("Left", left);
                    formItems.Add("Width", oForm.Width - 20);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 100);
                    formItems.Add("UID", itemName);

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    Top = Top + 105;

                    itemName = "3";
                    formItems = new Dictionary<string, object>();
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", 5);
                    formItems.Add("Width", 65);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", "OK");

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    itemName = "2";
                    formItems = new Dictionary<string, object>();
                    formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    formItems.Add("Left", 75);
                    formItems.Add("Width", 65);
                    formItems.Add("Top", Top);
                    formItems.Add("Height", 19);
                    formItems.Add("UID", itemName);
                    formItems.Add("Caption", BDOSResources.getTranslate("Close"));

                    FormsB1.createFormItem(oForm, formItems, out errorText);
                    if (errorText != null)
                    {
                        return;
                    }

                    if (oDocForm == null)
                    {
                        itemName = "CreateDocs";
                        formItems = new Dictionary<string, object>();
                        formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                        formItems.Add("Left", 145);
                        formItems.Add("Width", 150);
                        formItems.Add("Top", Top);
                        formItems.Add("Height", 19);
                        formItems.Add("UID", itemName);
                        formItems.Add("Caption", BDOSResources.getTranslate("CreateDocument"));

                        FormsB1.createFormItem(oForm, formItems, out errorText);
                        if (errorText != null)
                        {
                            return;
                        }
                    }

                    //SAPbouiCOM.Grid OGrid = ((SAPbouiCOM.Grid)(oForm.Items.Item("WBMatrix").Specific));
                    //ზედნადებების ცხრილი
                    SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("WBMatrix").Specific));
                    SAPbouiCOM.Columns oColumns = oMatrix.Columns;
                    SAPbouiCOM.Column oColumn;

                    string WbNo = "";
                    string WBTIN = "";
                    int DocEntry = 0;
                    string DocType = "";
                    string docDate = "";
                    string oCNTp = "";

                    if (oDocForm != null)
                    {
                        SAPbouiCOM.DBDataSource DocDBSourceOCRD = oDocForm.DataSources.DBDataSources.Item(1);
                        string CardCode = DocDBSourceOCRD.GetValue("CardCode", 0);

                        SAPbobsCOM.BusinessPartners oBP;
                        oBP = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);
                        oBP.GetByKey(CardCode);

                        WBTIN = oBP.UserFields.Fields.Item("LicTradNum").Value;

                        SAPbouiCOM.DBDataSource oDBDataSource = oDocForm.DataSources.DBDataSources.Item(0);
                        WbNo = oDBDataSource.GetValue("U_BDO_WBNo", 0);
                        docDate = oDBDataSource.GetValue("DocDate", 0);
                        //DateTime docDate = DateTime.ParseExact(DocDBSourceOCRD.GetValue("DocDate", 0), "yyyyMMdd", CultureInfo.InvariantCulture);

                        try
                        {
                            DocEntry = Convert.ToInt32(oDBDataSource.GetValue("DocEntry", 0));
                        }
                        catch
                        {
                            DocEntry = 0;
                        }

                        if (oDBDataSource.TableName == "OPCH")
                        {
                            DocType = "1"; // A/P Invoice
                        }
                        else if (oDBDataSource.TableName == "ORPC")
                        {
                            DocType = "2"; // Credit Memo
                            oCNTp = oDBDataSource.GetValue("U_BDO_CNTp", 0);
                        }
                        else if (oDBDataSource.TableName == "OPDN")
                        {
                            DocType = "3"; // Goods Receipt PO
                        }
                        else if (oDBDataSource.TableName == "OCPI")
                        {
                            DocType = "4"; // Credit Memo
                            oCNTp = oDBDataSource.GetValue("U_BDO_CNTp", 0);
                        }
                    }

                    oForm.DataSources.UserDataSources.Item("WBNo").ValueEx = oCNTp == "1" ? "" : WbNo;
                    oForm.DataSources.UserDataSources.Item("WaybType").ValueEx = oCNTp == "1" ? "1" : "";
                    oForm.DataSources.UserDataSources.Item("WBSuplNo").ValueEx = WBTIN;
                    if (DocEntry > 0)
                    {
                        oForm.DataSources.UserDataSources.Item("DocEntry").ValueEx = DocEntry.ToString();
                    }
                    oForm.DataSources.UserDataSources.Item("DocType").ValueEx = DocType;
                    oForm.DataSources.UserDataSources.Item("StartDate").ValueEx = docDate;
                    oForm.DataSources.UserDataSources.Item("EndDate").ValueEx = docDate;

                    WayBill oWayBill;
                    Dictionary<string, Dictionary<string, string>> waybills_map = getDataFromRS(oForm, out oWayBill, out errorText);

                    oWayBill.get_error_codes("", "", out errorText);

                    if (errorText != null)
                    {
                        Program.uiApp.MessageBox(errorText);
                    }
                    //RS ცხრილის მიღება - დასასრული

                    SAPbouiCOM.DataTable oDataTable;
                    oDataTable = oForm.DataSources.DataTables.Add("WBTable");
                    oDataTable.Columns.Add("#", SAPbouiCOM.BoFieldsType.ft_Text, 20); // 0 - ინდექსი გვჭირდება SetValue-ს პირველ პარემტრად
                    oDataTable.Columns.Add("WBNo", SAPbouiCOM.BoFieldsType.ft_Text, 20); //1
                    oDataTable.Columns.Add("WBID", SAPbouiCOM.BoFieldsType.ft_Text, 20); //2
                    oDataTable.Columns.Add("WBStat", SAPbouiCOM.BoFieldsType.ft_Text, 20); //3                
                    oDataTable.Columns.Add("WBSupName", SAPbouiCOM.BoFieldsType.ft_Text, 20); //4
                    oDataTable.Columns.Add("WBActDate", SAPbouiCOM.BoFieldsType.ft_Date, 20); //5
                    oDataTable.Columns.Add("WBStartAdd", SAPbouiCOM.BoFieldsType.ft_Text, 20); //6
                    oDataTable.Columns.Add("WBEndAdd", SAPbouiCOM.BoFieldsType.ft_Text, 20); //7
                    oDataTable.Columns.Add("WBSUM", SAPbouiCOM.BoFieldsType.ft_Sum, 20); //8
                    oDataTable.Columns.Add("WBSupTIN", SAPbouiCOM.BoFieldsType.ft_Text, 20); //9
                    oDataTable.Columns.Add("WBCheckbox", SAPbouiCOM.BoFieldsType.ft_Text, 20); //10
                    oDataTable.Columns.Add("APInvoice", SAPbouiCOM.BoFieldsType.ft_Text, 20); //11
                    oDataTable.Columns.Add("GdsRcpPO", SAPbouiCOM.BoFieldsType.ft_Text, 20); //12
                    oDataTable.Columns.Add("CredMemo", SAPbouiCOM.BoFieldsType.ft_Text, 20); //13    
                    oDataTable.Columns.Add("TYPE", SAPbouiCOM.BoFieldsType.ft_Text, 20); //14
                    oDataTable.Columns.Add("WBBlankAgr", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20); //15
                    oDataTable.Columns.Add("WBCOMMENT", SAPbouiCOM.BoFieldsType.ft_Text, 20); //16
                    oDataTable.Columns.Add("WBCheckb", SAPbouiCOM.BoFieldsType.ft_Text, 20); //17
                    oDataTable.Columns.Add("WBWhs", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20); //18
                    oDataTable.Columns.Add("WBProject", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20); //19
                    oDataTable.Columns.Add("WBStartDat", SAPbouiCOM.BoFieldsType.ft_Date, 20); //5

                    int rowCounter = 1;
                    int rowIndex = 0;

                    foreach (var map_record in waybills_map)
                    {
                        string WBID = map_record.Key;

                        Dictionary<string, string> Waybill_Header = map_record.Value;

                        string WBNo = Waybill_Header["WAYBILL_NUMBER"];
                        string WBStat = Waybill_Header["STATUS"];
                        string WBSupName = Waybill_Header["SELLER_NAME"];
                        string WBActDate = Waybill_Header["ACTIVATE_DATE"].Replace("T", " ");
                        string WBStartDat = Waybill_Header["BEGIN_DATE"].Replace("T", " ");
                        string WBStartAdd = Waybill_Header["START_ADDRESS"];
                        string WBEndAdd = Waybill_Header["END_ADDRESS"];
                        string WBSupTIN = Waybill_Header["SELLER_TIN"];
                        double WBSUM = Convert.ToDouble(Waybill_Header["FULL_AMOUNT"], CultureInfo.InvariantCulture);
                        string TYPE = Waybill_Header["TYPE"];
                        string WBCOM = Waybill_Header["WAYBILL_COMMENT"];

                        DateTime ActDt = new DateTime(1, 1, 1);
                        DateTime StartDt = new DateTime(1, 1, 1);

                        if (DateTime.TryParseExact(WBActDate, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture, DateTimeStyles.None, out ActDt) == false)
                        {
                            ActDt = new DateTime(1, 1, 1);
                        }
                        if (DateTime.TryParseExact(WBStartDat, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture, DateTimeStyles.None, out StartDt) == false)
                        {
                            StartDt = new DateTime(1, 1, 1);
                        }
                        if (TYPE == "5")
                        {
                            TYPE = "Return";
                        }
                        else
                        {
                            TYPE = "Procurement";
                        }

                        oDataTable.Rows.Add();
                        oDataTable.SetValue(0, rowIndex, rowCounter);
                        oDataTable.SetValue(1, rowIndex, WBNo);
                        oDataTable.SetValue(2, rowIndex, WBID);
                        oDataTable.SetValue(3, rowIndex, BDO_WBReceivedDocs.detectWBStatus(WBStat));
                        oDataTable.SetValue(4, rowIndex, WBSupName);
                        oDataTable.SetValue(5, rowIndex, ActDt);
                        oDataTable.SetValue(6, rowIndex, WBStartAdd);
                        oDataTable.SetValue(7, rowIndex, WBEndAdd);
                        oDataTable.SetValue(8, rowIndex, WBSUM);
                        oDataTable.SetValue(9, rowIndex, WBSupTIN);
                        oDataTable.SetValue(10, rowIndex, "0");
                        oDataTable.SetValue(14, rowIndex, TYPE);
                        oDataTable.SetValue(16, rowIndex, WBCOM);
                        oDataTable.SetValue("WBStartDat", rowIndex, StartDt);



                        if (!string.IsNullOrEmpty(WBEndAdd))
                        {
                            string whsCode = BDOSWarehouseAddresses.GetWhsByAddress(WBEndAdd);
                            if (!string.IsNullOrEmpty(whsCode))
                            {
                                oDataTable.SetValue(18, rowIndex, whsCode);
                            }
                        }
                        string LinkedDocType = "";

                        int LinkedDocEntryInvoice = 0;
                        BDO_WBReceivedDocs.getInvoiceByWB(WBID, out LinkedDocType, out LinkedDocEntryInvoice, out var linkedWhsInvoice, out var linkedProjectInvoice, out errorText);

                        int LinkedDocEntryGoodsReceipePO = 0;
                        BDO_WBReceivedDocs.getGoodsReceipePOByWB(WBID, out LinkedDocType, out LinkedDocEntryGoodsReceipePO, out var linkedWhsGoodsReceipePO, out var linkedProjectGoodsReceipePO, out errorText);

                        BDO_WBReceivedDocs.GetDraftByWB(WBID, out var linkedDocTypeDraft, out var linkedDocEntryDraft, out var linkedWhsDraft, out var linkedProjectDraft, out errorText);

                        int LinkedDocEntryMemo = 0;
                        BDO_WBReceivedDocs.getMemoByWB(WBID, out LinkedDocType, out LinkedDocEntryMemo, out var linkedWhsMemo, out var linkedProjectMemo, out errorText);

                        if (LinkedDocEntryInvoice != 0)
                        {
                            oDataTable.SetValue(11, rowIndex, LinkedDocEntryInvoice.ToString());
                            oDataTable.SetValue(18, rowIndex, linkedWhsInvoice);
                            oDataTable.SetValue(19, rowIndex, linkedProjectInvoice);
                        }

                        if (LinkedDocEntryGoodsReceipePO != 0)
                        {
                            oDataTable.SetValue(12, rowIndex, LinkedDocEntryGoodsReceipePO.ToString());
                            oDataTable.SetValue(18, rowIndex, linkedWhsGoodsReceipePO);
                            oDataTable.SetValue(19, rowIndex, linkedProjectGoodsReceipePO);
                        }

                        if (linkedDocEntryDraft != 0)
                        {
                            if (linkedDocTypeDraft == "APInvoiceDraft")
                            {
                                oDataTable.SetValue(11, rowIndex, linkedDocEntryDraft.ToString());
                            }
                            else if (linkedDocTypeDraft == "GdsRcptDraft")
                            {
                                oDataTable.SetValue(12, rowIndex, linkedDocEntryDraft.ToString());
                            }

                            oDataTable.SetValue(18, rowIndex, linkedWhsDraft);
                            oDataTable.SetValue(19, rowIndex, linkedProjectDraft);
                        }

                        if (LinkedDocEntryMemo != 0)
                        {
                            oDataTable.SetValue(13, rowIndex, LinkedDocEntryMemo.ToString());
                            oDataTable.SetValue(18, rowIndex, linkedWhsMemo);
                            oDataTable.SetValue(19, rowIndex, linkedProjectMemo);
                        }

                        rowCounter++;
                        rowIndex++;
                    }

                    oColumn = oColumns.Add("#", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = "#";
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind("WBTable", "#");

                    if (oDocForm == null)
                    {
                        oColumn = oColumns.Add("WBCheckbox", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
                        oColumn.TitleObject.Caption = "";
                        oColumn.Width = 100;
                        oColumn.Editable = true;
                        oColumn.DataBind.Bind("WBTable", "WBCheckbox");
                    }

                    oColumn = oColumns.Add("WBNo", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("WaybillNumber");
                    oColumn.Width = 100;
                    oColumn.Editable = false;

                    oColumn.DataBind.Bind("WBTable", "WBNo");

                    oColumn = oColumns.Add("WBID", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("WaybillID");
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind("WBTable", "WBID");

                    oColumn = oColumns.Add("APInvoice", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("Purchase");
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind("WBTable", "APInvoice");
                    SAPbouiCOM.LinkedButton oLink = oColumn.ExtendedObject;
                    oLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_PurchaseInvoice;

                    oColumn = oColumns.Add("GdsRcpPO", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("GoodsRcptPO");
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind("WBTable", "GdsRcpPO");
                    oLink = oColumn.ExtendedObject;
                    oLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_GoodsReceiptPO;

                    oColumn = oColumns.Add("CredMemo", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("Correction");
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind("WBTable", "CredMemo");
                    oLink = oColumn.ExtendedObject;
                    oLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_PurchaseInvoiceCreditMemo;

                    oColumn = oColumns.Add("WBStat", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("Status");
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind("WBTable", "WBStat");

                    SAPbobsCOM.Documents oAPInv;
                    oAPInv = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices);

                    SAPbobsCOM.ValidValues BaseValues = oAPInv.UserFields.Fields.Item("U_BDO_WBSt").ValidValues;

                    oColumn.ValidValues.Add("-1", " ");
                    oColumn.ValidValues.Add("1", BDOSResources.getTranslate("Saved"));
                    oColumn.ValidValues.Add("2", BDOSResources.getTranslate("Active"));
                    oColumn.ValidValues.Add("3", BDOSResources.getTranslate("finished"));
                    oColumn.ValidValues.Add("4", BDOSResources.getTranslate("deleted"));
                    oColumn.ValidValues.Add("5", BDOSResources.getTranslate("cancelled"));
                    oColumn.ValidValues.Add("6", BDOSResources.getTranslate("SentToTransporter"));

                    oColumn.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
                    oColumn.DisplayDesc = true;

                    oColumn = oColumns.Add("WBSupName", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("BPName");
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind("WBTable", "WBSupName");

                    oColumn = oColumns.Add("WBSupTIN", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("BPTin");
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind("WBTable", "WBSupTIN");


                    //--------------------
                    SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams;
                    SAPbouiCOM.ChooseFromList oCFL;
                    SAPbouiCOM.ChooseFromListCollection oCFLs = oForm.ChooseFromLists;
                    //--------------------

                    oCFLCreationParams = Program.uiApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);

                    oCFLCreationParams.MultiSelection = false;
                    oCFLCreationParams.ObjectType = "1250000025";
                    oCFLCreationParams.UniqueID = "WBBlankAgr_CFLA";
                    oCFL = oCFLs.Add(oCFLCreationParams);

                    oColumn = oColumns.Add("WBBlankAgr", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("BlnkAgr");
                    oColumn.Width = 100;
                    oColumn.Editable = true;
                    oColumn.DataBind.Bind("WBTable", "WBBlankAgr");

                    oLink = oColumn.ExtendedObject;
                    oLink.LinkedObjectType = "1250000025"; //SAPbouiCOM.BoLinkedObject.

                    oColumn.ChooseFromListUID = "WBBlankAgr_CFLA";
                    oColumn.ChooseFromListAlias = "AbsID";

                    //--------------------

                    oColumn = oColumns.Add("WBSum", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("Amount");
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind("WBTable", "WBSum");

                    oColumn = oColumns.Add("WBActDate", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("ActivateDate");
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind("WBTable", "WBActDate");

                    oColumn = oColumns.Add("WBStartDat", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("TransportDate");
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind("WBTable", "WBStartDat");


                    oColumn = oColumns.Add("WBStartAdd", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("StartAddress");
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind("WBTable", "WBStartAdd");

                    oColumn = oColumns.Add("WBEndAdd", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("EndAddress");
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind("WBTable", "WBEndAdd");

                    oColumn = oColumns.Add("TYPE", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("Type");
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind("WBTable", "TYPE");

                    oColumn = oColumns.Add("WBCOMMENT", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("Comment");
                    oColumn.Width = 100;
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind("WBTable", "WBCOMMENT");

                    oColumn.ValidValues.Add("1", BDOSResources.getTranslate("Return"));//1
                    oColumn.ValidValues.Add("2", BDOSResources.getTranslate("Purchase"));//2

                    oColumn.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
                    oColumn.DisplayDesc = true;

                    oColumn = oColumns.Add("WBCheckb", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("saveWhs");
                    oColumn.Width = 100;
                    oColumn.Editable = true;
                    oColumn.DataBind.Bind("WBTable", "WBCheckb");


                    //WBWhs
                    FormsB1.addChooseFromList(oForm, false, "64", "WBWarehouseCFL");
                    oColumn = oColumns.Add("WBWhs", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("Warehouse");
                    oColumn.Width = 50;
                    oColumn.Editable = true;
                    oColumn.DataBind.Bind("WBTable", "WBWhs");

                    oLink = oColumn.ExtendedObject;
                    oLink.LinkedObjectType = "64"; //SAPbouiCOM.BoLinkedObject.

                    oColumn.ChooseFromListUID = "WBWarehouseCFL";
                    oColumn.ChooseFromListAlias = "WhsCode";

                    //WBProject
                    FormsB1.addChooseFromList(oForm, false, "63", "WBProjectCFL");
                    oColumn = oColumns.Add("WBProject", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("Project");
                    oColumn.Width = 50;
                    oColumn.Editable = true;
                    oColumn.DataBind.Bind("WBTable", "WBProject");

                    oLink = oColumn.ExtendedObject;
                    oLink.LinkedObjectType = "63";

                    oColumn.ChooseFromListUID = "WBProjectCFL";
                    oColumn.ChooseFromListAlias = "PrjCode";


                    //-----------
                    oMatrix.Clear();
                    oMatrix.LoadFromDataSource();
                    oMatrix.AutoResizeColumns();

                    //საქონლის ცხრილი
                    oDataTable = oForm.DataSources.DataTables.Add("WbGdsTable");
                    oDataTable.Columns.Add("#", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20); //0
                    oDataTable.Columns.Add("WBNo", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20);//1
                    oDataTable.Columns.Add("WBBarcode", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 50);//2
                    oDataTable.Columns.Add("WBItmName", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 150);//3
                    oDataTable.Columns.Add("WBUntCode", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);//4
                    //oDataTable.Columns.Add("WBUntName", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);//5
                    oDataTable.Columns.Add("ItemCode", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20);//6
                    oDataTable.Columns.Add("WBQty", SAPbouiCOM.BoFieldsType.ft_Quantity, 20);//7
                    oDataTable.Columns.Add("WBPrice", SAPbouiCOM.BoFieldsType.ft_Price, 20);//8
                    oDataTable.Columns.Add("WBLPrice", SAPbouiCOM.BoFieldsType.ft_Price, 20);//8
                    oDataTable.Columns.Add("WBSum", SAPbouiCOM.BoFieldsType.ft_Sum, 20);//9     
                    oDataTable.Columns.Add("WBUntCdRS", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20);//10
                    oDataTable.Columns.Add("WBPrjCode", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);//11
                    oDataTable.Columns.Add("RSVatCode", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20);//12
                    oDataTable.Columns.Add("ItemName", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);//13
                    oDataTable.Columns.Add("WbUntNmRS", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);//14
                    oDataTable.Columns.Add("DistNumber", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);//15

                    Dictionary<string, string> activeDimensionsList = CommonFunctions.getActiveDimensionsList(out errorText);

                    for (int i = 1; i <= activeDimensionsList.Count; i++)
                    {
                        oDataTable.Columns.Add("DistrRul" +  i.ToString(), SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 8);
                    }

                    oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("WBGdMatrix").Specific));
                    oColumns = oMatrix.Columns;

                    oColumn = oColumns.Add("#", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = "#";
                    oColumn.Width = 40;
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind("WbGdsTable", "#");

                    oColumn = oColumns.Add("WBNo", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("WaybillNumber");
                    oColumn.Width = 40;
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind("WbGdsTable", "WBNo");

                    oColumn = oColumns.Add("WBBarcode", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("Barcode");
                    oColumn.Width = 40;
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind("WbGdsTable", "WBBarcode");

                    oColumn = oColumns.Add("RSVatCode", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("VatCode");
                    oColumn.Width = 40;
                    oColumn.Editable = false;
                    oColumn.Visible = false;
                    oColumn.DataBind.Bind("WbGdsTable", "RSVatCode");

                    oColumn = oColumns.Add("WBItmName", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("ItemName");
                    oColumn.Width = 40;
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind("WbGdsTable", "WBItmName");

                    //UOM Code RS
                    oColumn = oColumns.Add("WBUntCdRS", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("UomRsCode");
                    oColumn.Width = 40;
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind("WbGdsTable", "WBUntCdRS");
                    //UOM Code RS

                    //UOM Name RS
                    oColumn = oColumns.Add("WbUntNmRS", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("UomNameRS");
                    oColumn.Width = 40;
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind("WbGdsTable", "WbUntNmRS");
                    //UOM Name RS

                    //UOM Code
                    oCFLCreationParams.MultiSelection = false;
                    oCFLCreationParams.ObjectType = "10000199";
                    oCFLCreationParams.UniqueID = "CFLUoMCdB";
                    oCFL = oCFLs.Add(oCFLCreationParams);

                    oColumn = oColumns.Add("WBUntCode", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("UomCode");
                    oColumn.Width = 40;
                    oColumn.Editable = true;
                    oColumn.DataBind.Bind("WbGdsTable", "WBUntCode");
                    oColumn.ChooseFromListUID = "CFLUoMCdB";
                    oColumn.ChooseFromListAlias = "UoMCode";
                    //UOM Code

                    //UOM Name
                    //oColumn = oColumns.Add("WBUntName", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    //oColumn.TitleObject.Caption = BDOSResources.getTranslate("UomName");
                    //oColumn.Width = 40;
                    ////oColumn.Editable = false;
                    //oColumn.Visible = false;
                    //oColumn.DataBind.Bind("WbGdsTable", "WBUntName");
                    //UOM Name

                    //item column         
                    oCFLCreationParams = Program.uiApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);

                    oCFLCreationParams.MultiSelection = false;
                    oCFLCreationParams.ObjectType = "4";
                    oCFLCreationParams.UniqueID = "CFLItmCd";
                    oCFL = oCFLs.Add(oCFLCreationParams);

                    oColumn = oColumns.Add("ItemCode", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("ItemCode");
                    oColumn.DataBind.Bind("WbGdsTable", "ItemCode");
                    oColumn.Width = 40;
                    oColumn.Editable = true;
                    oLink = oColumn.ExtendedObject;
                    oLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Items;
                    oColumn.ChooseFromListUID = "CFLItmCd";
                    oColumn.ChooseFromListAlias = "ItemCode";

                    oColumn = oColumns.Add("ItemName", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("ItemName");
                    oColumn.DataBind.Bind("WbGdsTable", "ItemName");
                    oColumn.Width = 40;
                    oColumn.Editable = false;


                    oColumn = oColumns.Add("DistNumber", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("BatchNumber");
                    oColumn.DataBind.Bind("WbGdsTable", "DistNumber");
                    oColumn.Width = 40;
                    oColumn.Editable = true;


                    //item column


                    oCFLCreationParams = Program.uiApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);

                    oCFLCreationParams.MultiSelection = false;
                    oCFLCreationParams.ObjectType = "63";
                    oCFLCreationParams.UniqueID = "WBProject_CFLA";
                    oCFL = oCFLs.Add(oCFLCreationParams);

                    oColumn = oColumns.Add("WBPrjCode", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("Project");
                    oColumn.DataBind.Bind("WbGdsTable", "WBPrjCode");
                    oColumn.Width = 40;
                    oColumn.Editable = true;
                    // oLink = oColumn.ExtendedObject;
                    // oLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Items;
                    oColumn.ChooseFromListUID = "WBProject_CFLA";
                    oColumn.ChooseFromListAlias = "PrjCode";

                    oColumn = oColumns.Add("WBQty", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("Quantity");
                    oColumn.Width = 40;
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind("WbGdsTable", "WBQty");

                    oColumn = oColumns.Add("WBPrice", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("Price");
                    oColumn.Width = 40;
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind("WbGdsTable", "WBPrice");

                    oColumn = oColumns.Add("WBLPrice", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("LastPrice");
                    oColumn.Width = 40;
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind("WbGdsTable", "WBLPrice");

                    oColumn = oColumns.Add("WBSum", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.TitleObject.Caption = BDOSResources.getTranslate("Amount");
                    oColumn.Width = 40;
                    oColumn.Editable = false;
                    oColumn.DataBind.Bind("WbGdsTable", "WBSum");

                    for (int i = 1; i <= activeDimensionsList.Count; i++)
                    {
                        objectType = "62";
                        string uniqueID_lf_DistrRule = "Rule_CFL" + i.ToString() + "A";
                        FormsB1.addChooseFromList(oForm, false, objectType, uniqueID_lf_DistrRule);


                        oColumn = oColumns.Add("DistrRul" + i.ToString(), SAPbouiCOM.BoFormItemTypes.it_EDIT);
                        oColumn.TitleObject.Caption = BDOSResources.getTranslate(activeDimensionsList[i.ToString()]);
                        oColumn.DataBind.Bind("WbGdsTable", "DistrRul" + i.ToString());
                        oColumn.Width = 100;
                        oColumn.Editable = true;
                        oColumn.ChooseFromListUID = uniqueID_lf_DistrRule;
                        oColumn.ChooseFromListAlias = "OcrCode";
                    }

                    fillWBGoods(oForm, 1, false, out errorText);

                    resizeItems(oForm, out errorText);

                }
                // oForm.Settings.MatrixUID = "WBMatrix";
                oForm.Visible = true;
                oForm.Select();

                FormsB1.WB_TAX_AuthorizationsItems(oForm);

            }
            GC.Collect();
        }
        public static void addMenus()
        {
            SAPbouiCOM.Menus moduleMenus;
            SAPbouiCOM.MenuItem menuItem;
            SAPbouiCOM.MenuItem fatherMenuItem;
            SAPbouiCOM.MenuCreationParams oCreationPackage;

            try
            {
                // Find the id of the menu into wich you want to add your menu item
                // ModuleMenuId = "43520"
                fatherMenuItem = Program.uiApp.Menus.Item("2304");

                // Get the menu collection of SAP Business One
                //moduleMenus = menuItem.SubMenus;

                // Add a pop-up menu item
                oCreationPackage = (SAPbouiCOM.MenuCreationParams)Program.uiApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                oCreationPackage.Checked = false;
                oCreationPackage.Enabled = true;
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "BDO_WBR";
                oCreationPackage.String = BDOSResources.getTranslate("WaybillReceived");
                oCreationPackage.Position = fatherMenuItem.SubMenus.Count - 1;

                menuItem = fatherMenuItem.SubMenus.AddEx(oCreationPackage);
            }
            catch
            {

            }

            try
            {
                // Find the id of the menu into wich you want to add your menu item
                // ModuleMenuId = "43520"
                menuItem = Program.uiApp.Menus.Item("43520");

                // Get the menu collection of SAP Business One
                moduleMenus = menuItem.SubMenus;
                fatherMenuItem = moduleMenus.Item("11520");

                // Add a pop-up menu item
                oCreationPackage = (SAPbouiCOM.MenuCreationParams)Program.uiApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                oCreationPackage.Checked = false;
                oCreationPackage.Enabled = true;
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "BDO_UoMRS";
                oCreationPackage.String = BDOSResources.getTranslate("UomRsCodes");
                oCreationPackage.Position = 4;

                menuItem = fatherMenuItem.SubMenus.AddEx(oCreationPackage);
            }
            catch
            {

            }
        }
        public static void chooseFromList(SAPbouiCOM.Form oForm, SAPbouiCOM.ChooseFromListEvent oCFLEvento, string ItemUID, bool BeforeAction, int row, out string errorText)
        {
            errorText = null;

            if (BeforeAction == false)
            {
                try
                {
                    if (oCFLEvento.SelectedObjects == null)
                    {
                        errorText = "noselectedobjects";
                        return;
                    }

                    if (oCFLEvento.ChooseFromListUID == "Whs_CFL")
                    {
                        SAPbouiCOM.DataTable oDataTableSelectedObjects = oCFLEvento.SelectedObjects;
                        string WhsCode = oDataTableSelectedObjects.GetValue("WhsCode", 0);
                        string projectCode = oDataTableSelectedObjects.GetValue("U_BDOSPrjCod", 0);
                        LanguageUtils.IgnoreErrors<string>(() => oForm.Items.Item("Whs").Specific.Value = WhsCode);
                        LanguageUtils.IgnoreErrors<string>(() => oForm.Items.Item("PrjCode").Specific.Value = projectCode);
                        SAPbouiCOM.EditText oEdit = oForm.Items.Item("Whs").Specific;
                        oEdit.Value = WhsCode;
                    }

                    else if (oCFLEvento.ChooseFromListUID.Length > 1 && oCFLEvento.ChooseFromListUID.Substring(0, oCFLEvento.ChooseFromListUID.Length - 2) == "Rule_CFL")
                    {
                        SAPbouiCOM.DataTable oDataTableSelectedObjects = oCFLEvento.SelectedObjects;
                        string ocrCode = oDataTableSelectedObjects.GetValue("OcrCode", 0);
                        string j = oCFLEvento.ChooseFromListUID.Substring(oCFLEvento.ChooseFromListUID.Length - 2,1);
                        SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("WBGdMatrix").Specific));
                        LanguageUtils.IgnoreErrors<string>(() =>oMatrix.Columns.Item("DistrRul"+j).Cells.Item(oCFLEvento.Row).Specific.Value = ocrCode);
                        saveWBGoods(oMatrix, 1, false);

                    }

                    else if (oCFLEvento.ChooseFromListUID == "CFLItmCd")
                    {
                        SAPbouiCOM.DataTable oDataTableSelectedObjects = oCFLEvento.SelectedObjects;
                        string ItemCode = oDataTableSelectedObjects.GetValue("ItemCode", 0);

                        SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("WBGdMatrix").Specific));
                        LanguageUtils.IgnoreErrors<string>(() => oMatrix.Columns.Item("ItemCode").Cells.Item(oCFLEvento.Row).Specific.Value = ItemCode);
                        LanguageUtils.IgnoreErrors<string>(() => oMatrix.Columns.Item("ItemName").Cells.Item(oCFLEvento.Row).Specific.Value = oDataTableSelectedObjects.GetValue("ItemName", 0));
                    }

                    else if (oCFLEvento.ChooseFromListUID == "CFLUoMCdB")
                    {
                        SAPbouiCOM.DataTable oDataTableSelectedObjects = oCFLEvento.SelectedObjects;
                        string UoMCode = oDataTableSelectedObjects.GetValue("UomCode", 0);

                        SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("WBGdMatrix").Specific));
                        LanguageUtils.IgnoreErrors<string>(() => oMatrix.Columns.Item("WBUntCode").Cells.Item(oCFLEvento.Row).Specific.Value = UoMCode);
                    }

                    else if (oCFLEvento.ChooseFromListUID == "Project_CFLA")
                    {
                        SAPbouiCOM.DataTable oDataTableSelectedObjects = oCFLEvento.SelectedObjects;
                        string PrjCode = oDataTableSelectedObjects.GetValue("PrjCode", 0);
                        LanguageUtils.IgnoreErrors<string>(() => oForm.Items.Item("PrjCode").Specific.Value = PrjCode);
                        oForm.Items.Item("Whs").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        oForm.Items.Item("PrjCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);

                    }

                    else if (oCFLEvento.ChooseFromListUID == "WBBlankAgr_CFLA")
                    {
                        SAPbouiCOM.DataTable oDataTableSelectedObjects = oCFLEvento.SelectedObjects;
                        string WBBPCode = oDataTableSelectedObjects.GetValue("AbsID", 0).ToString();
                        var WBPrjCode = oDataTableSelectedObjects.GetValue("Project", 0).ToString();
                        SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("WBMatrix").Specific));
                        LanguageUtils.IgnoreErrors<string>(() => oMatrix.Columns.Item("WBBlankAgr").Cells.Item(oCFLEvento.Row).Specific.Value = WBBPCode);
                        LanguageUtils.IgnoreErrors<string>(() => oMatrix.Columns.Item("WBProject").Cells.Item(oCFLEvento.Row).Specific.Value = WBPrjCode);

                        FillGoodsProject(oForm, WBPrjCode);
                    }

                    else if (oCFLEvento.ChooseFromListUID == "WBProject_CFLA")
                    {
                        SAPbouiCOM.DataTable oDataTableSelectedObjects = oCFLEvento.SelectedObjects;
                        string WBPrjCode = oDataTableSelectedObjects.GetValue("PrjCode", 0);

                        SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("WBGdMatrix").Specific));
                        LanguageUtils.IgnoreErrors<string>(() => oMatrix.Columns.Item("WBPrjCode").Cells.Item(oCFLEvento.Row).Specific.Value = WBPrjCode);
                        saveWBGoods(oMatrix, oCFLEvento.Row, true);
                    }

                    else if (oCFLEvento.ChooseFromListUID == "WBWarehouseCFL")
                    {
                        SAPbouiCOM.DataTable oDataTableSelectedObjects = oCFLEvento.SelectedObjects;
                        string WBWhsCode = oDataTableSelectedObjects.GetValue("WhsCode", 0);

                        SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("WBMatrix").Specific));
                        LanguageUtils.IgnoreErrors<string>(() => oMatrix.Columns.Item("WBWhs").Cells.Item(oCFLEvento.Row).Specific.Value = WBWhsCode);

                        string blAgreement = oMatrix.Columns.Item("WBBlankAgr").Cells.Item(oCFLEvento.Row).Specific.Value.ToString();
                        string prjCode = oMatrix.Columns.Item("WBProject").Cells.Item(oCFLEvento.Row).Specific.Value.ToString();

                        if (string.IsNullOrEmpty(prjCode))
                        {
                            var WBPrjCode = oDataTableSelectedObjects.GetValue("U_BDOSPrjCod", 0);
                            LanguageUtils.IgnoreErrors<string>(() => oMatrix.Columns.Item("WBProject").Cells.Item(oCFLEvento.Row).Specific.Value = WBPrjCode);

                            FillGoodsProject(oForm, WBPrjCode);
                        }

                        //if (string.IsNullOrEmpty(blAgreement))
                        //{
                        //    var WBPrjCode = oDataTableSelectedObjects.GetValue("U_BDOSPrjCod", 0);
                        //    LanguageUtils.IgnoreErrors<string>(() => oMatrix.Columns.Item("WBProject").Cells.Item(oCFLEvento.Row).Specific.Value = WBPrjCode);

                        //    FillGoodsProject(oForm, WBPrjCode);
                        //}
                    }

                    else if (oCFLEvento.ChooseFromListUID == "WBProjectCFL")
                    {
                        SAPbouiCOM.DataTable oDataTableSelectedObjects = oCFLEvento.SelectedObjects;
                        string WBPrjCode = oDataTableSelectedObjects.GetValue("PrjCode", 0);

                        SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("WBMatrix").Specific));
                        LanguageUtils.IgnoreErrors<string>(() => oMatrix.Columns.Item("WBProject").Cells.Item(oCFLEvento.Row).Specific.Value = WBPrjCode);

                        SAPbouiCOM.Matrix oMatrix1 = ((SAPbouiCOM.Matrix)(oForm.Items.Item("WBGdMatrix").Specific));
                        FillGoodsProject(oForm, WBPrjCode);
                        saveWBGoods(oMatrix1, 1, false);
                    }

                    //else if (oCFLEvento.ChooseFromListUID == "CFLUoMCdB")
                    //{
                    //    SAPbouiCOM.DataTable oDataTableSelectedObjects;
                    //    oDataTableSelectedObjects = oCFLEvento.SelectedObjects;
                    //    string UoMName = oDataTableSelectedObjects.GetValue("UomName", 0);

                    //    SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("WBGdMatrix").Specific));
                    //    LanguageUtils.IgnoreErrors<string>(() => oMatrix.Columns.Item("WBUntName").Cells.Item(oCFLEvento.Row).Specific.Value = UoMName);
                    //}
                }
                catch (Exception ex)
                {
                    throw new Exception(ex.Message);
                }
                finally
                {
                    GC.Collect();
                }
            }
            else
            {
                if (oCFLEvento.ChooseFromListUID == "WBBlankAgr_CFLA")
                {
                    SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("WBMatrix").Specific));

                    string TIN = oMatrix.GetCellSpecific("WBSupTIN", oCFLEvento.Row).Value;
                    string cardName = "";

                    SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID);
                    SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();

                    SAPbouiCOM.Condition oCon = oCons.Add();
                    oCon.Alias = "BpCode";
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCon.CondVal = BusinessPartners.GetCardCodeByTin(TIN, "S", out cardName);
                    oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                    oCon = oCons.Add();
                    oCon.Alias = "Method";
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCon.CondVal = "M";
                    oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                    oCon = oCons.Add();
                    oCon.Alias = "Status";
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCon.CondVal = "A";
                    oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_NONE;

                    oCFL.SetConditions(oCons);
                }

                else if (oCFLEvento.ChooseFromListUID == "CFLUoMCdB")
                {
                    string sCFL_ID = oCFLEvento.ChooseFromListUID;
                    SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(sCFL_ID);
                    SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();

                    SAPbouiCOM.Matrix oMatrixGoods = oForm.Items.Item("WBGdMatrix").Specific;
                    SAPbouiCOM.EditText oEditTextGoods = (SAPbouiCOM.EditText)oMatrixGoods.Columns.Item("ItemCode").Cells.Item(row).Specific;
                    string ItemCode = oEditTextGoods.Value;

                    SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    string query =
                    "SELECT \"UGP1\".\"UomEntry\" FROM \"OITM\"" +
                    "INNER JOIN \"UGP1\" ON \"OITM\".\"UgpEntry\" = \"UGP1\".\"UgpEntry\"" +
                    "WHERE \"OITM\".\"ItemCode\" = N'" + ItemCode + "'";

                    try
                    {
                        oRecordSet.DoQuery(query);
                        int recordCount = oRecordSet.RecordCount;
                        int i = 1;

                        while (!oRecordSet.EoF)
                        {
                            SAPbouiCOM.Condition oCon = oCons.Add();
                            oCon.Alias = "UomEntry";
                            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                            oCon.CondVal = oRecordSet.Fields.Item("UomEntry").Value.ToString();
                            oCon.Relationship = (i == recordCount) ? SAPbouiCOM.BoConditionRelationship.cr_NONE : SAPbouiCOM.BoConditionRelationship.cr_OR;

                            i = i + 1;
                            oRecordSet.MoveNext();
                        }
                        oCFL.SetConditions(oCons);
                    }
                    catch (Exception ex)
                    {
                        errorText = ex.Message;
                    }
                }

                else if (oCFLEvento.ChooseFromListUID == "CFLItmCd")
                {
                    string sCFL_ID = oCFLEvento.ChooseFromListUID;
                    SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(sCFL_ID);
                    SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();

                    //SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    //string query = "SELECT \"ItemCode\" FROM \"OITM\"" +
                    //               "WHERE (\"ItemType\" = 'I' AND \"frozenFor\"='N' AND \"PrchseItem\" = 'Y') OR (\"ItemType\" = 'F' AND \"frozenFor\"='N' AND  \"PrchseItem\" = 'Y')";
                    try
                    {
                        //    oRecordSet.DoQuery(query);
                        //    int recordCount = oRecordSet.RecordCount;
                        //    int i = 1;

                        //    while (!oRecordSet.EoF)
                        //    {
                        //        SAPbouiCOM.Condition oCon = oCons.Add();
                        //        oCon.Alias = "ItemCode";
                        //        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        //        oCon.CondVal = oRecordSet.Fields.Item("ItemCode").Value.ToString();
                        //        oCon.Relationship = (i == recordCount) ? SAPbouiCOM.BoConditionRelationship.cr_NONE : SAPbouiCOM.BoConditionRelationship.cr_OR;

                        //        i = i + 1;
                        //        oRecordSet.MoveNext();
                        //    }


                        SAPbouiCOM.Condition oCon = oCons.Add();
                        oCon.Alias = "frozenFor";
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = "N";
                        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                        oCon = oCons.Add();
                        oCon.Alias = "PrchseItem";
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = "Y";
                        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                        oCon = oCons.Add();
                        oCon.BracketOpenNum = 1;
                        oCon.Alias = "ItemType";
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = "I";
                        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR;

                        oCon = oCons.Add();
                        oCon.Alias = "ItemType";
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = "F";
                        oCon.BracketCloseNum = 1;

                        oCFL.SetConditions(oCons);
                    }
                    catch (Exception ex)
                    {
                        errorText = ex.Message;
                    }
                }

                else if (oCFLEvento.ChooseFromListUID.Length > 1 && oCFLEvento.ChooseFromListUID.Substring(0, oCFLEvento.ChooseFromListUID.Length - 2) == "Rule_CFL")
                {
                    oForm.Freeze(true);
                    string dimensionCode = oCFLEvento.ChooseFromListUID.Substring(oCFLEvento.ChooseFromListUID.Length - 2, 1);
                    SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();

                    SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    string query = @"SELECT
	                                     ""OCR1"".""OcrCode"",
	                                     ""OOCR"".""DimCode"" 
                                    FROM ""OCR1"" 
                                    LEFT JOIN ""OOCR"" ON ""OCR1"".""OcrCode"" = ""OOCR"".""OcrCode"" 
                                    WHERE ""OOCR"".""DimCode"" = '" + dimensionCode + "'";

                    try
                    {
                        oRecordSet.DoQuery(query);
                        int recordCount = oRecordSet.RecordCount;
                        int i = 1;

                        while (!oRecordSet.EoF)
                        {
                            SAPbouiCOM.Condition oCon = oCons.Add();
                            oCon.Alias = "OcrCode";
                            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                            oCon.CondVal = oRecordSet.Fields.Item("OcrCode").Value.ToString();
                            oCon.Relationship = (i == recordCount) ? SAPbouiCOM.BoConditionRelationship.cr_NONE : SAPbouiCOM.BoConditionRelationship.cr_OR;

                            i = i + 1;
                            oRecordSet.MoveNext();
                        }

                        //თუ არცერთი შეესაბამება ცარიელზე გავიდეს
                        if (oCons.Count == 0)
                        {
                            SAPbouiCOM.Condition oCon = oCons.Add();
                            oCon.Alias = "OcrCode";
                            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                            oCon.CondVal = "";
                        }

                        oForm.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID).SetConditions(oCons);
                    }

                    catch (Exception ex)
                    {
                        errorText = ex.Message;
                    }
                    oForm.Freeze(false);
                }
            }
        }

        private static void FillGoodsProject(Form oForm, string wBPrjCode)
        {
            Dictionary<string, string> activeDimensionsList = CommonFunctions.getActiveDimensionsList(out string errorText);
            var goodsMatrix = (Matrix)oForm.Items.Item("WBGdMatrix").Specific;

            string[][] wbTempTable = null;
            string[][] array_GOODS = null;
            if (wbTempLines.TryGetValue(goodsMatrix.GetCellSpecific("WBNo", 1).Value, out wbTempTable))
            {
                array_GOODS = wbTempLines[goodsMatrix.GetCellSpecific("WBNo", 1).Value];
            }

            for (var goodsRow = 1; goodsRow <= goodsMatrix.RowCount; goodsRow++)
            {
                if (array_GOODS != null) {
                    if(array_GOODS[goodsRow-1][14 + activeDimensionsList.Count] != "1") LanguageUtils.IgnoreErrors<string>(() => goodsMatrix.GetCellSpecific("WBPrjCode", goodsRow).Value = wBPrjCode);
                }
                else{
                    LanguageUtils.IgnoreErrors<string>(() => goodsMatrix.GetCellSpecific("WBPrjCode", goodsRow).Value = wBPrjCode);
                 }
            }
        }
        public static void findItemCode(SAPbouiCOM.Form oForm, string WBItmName, string WBBarcode)
        {
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("WBMatrix").Specific;
        }
        public static void fillWBGoods(SAPbouiCOM.Form oForm, int row, bool refresh, out string errorText)
        {
            errorText = null;

            var Stopwatch = new Stopwatch();

            //Stopwatch.Restart();

            try
            {
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("WBMatrix").Specific;

                if (oMatrix.Columns.Item(1).Title == "")
                {
                    if (oMatrix.GetCellSpecific("WBCheckbox", row).Checked)
                    {
                        if ((oMatrix.GetCellSpecific("APInvoice", row).Value != "") || (oMatrix.GetCellSpecific("GdsRcpPO", row).Value != "") || (oMatrix.GetCellSpecific("CredMemo", row).Value != ""))
                        {
                            SAPbouiCOM.CheckBox oCheckBox = (SAPbouiCOM.CheckBox)oMatrix.Columns.Item("WBCheckbox").Cells.Item(row).Specific;
                            oCheckBox.Checked = false;
                            Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("DocumentLinkedToWaybill"));
                        }
                    }
                }

                if (oMatrix.RowCount > 0)
                {
                    oForm.Freeze(false);
                    for (int i = 1; i <= oMatrix.RowCount; i++)
                    {
                        oMatrix.CommonSetting.SetRowBackColor(i, FormsB1.getLongIntRGB(231, 231, 231));
                    }

                    oMatrix.CommonSetting.SetRowBackColor(row, FormsB1.getLongIntRGB(255, 255, 153));
                    oForm.Freeze(true);
                }

                SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("WbGdsTable");
                SAPbouiCOM.Matrix oMatrixGoods = (SAPbouiCOM.Matrix)oForm.Items.Item("WBGdMatrix").Specific;

                string GoodsWB = "";
                string HeadWB = "";
                string cardName;

                if (oMatrixGoods.RowCount > 0)
                {
                    SAPbouiCOM.EditText oEditTextGoods = (SAPbouiCOM.EditText)oMatrixGoods.Columns.Item("WBNo").Cells.Item(1).Specific;
                    GoodsWB = oEditTextGoods.Value;
                    //oForm.Freeze(false);
                    //oMatrixGoods.Clear();
                    //oForm.Freeze(true);

                }

                if (oMatrix.RowCount > 0)
                {
                    SAPbouiCOM.EditText oEditTextHeader = (SAPbouiCOM.EditText)oMatrix.Columns.Item("WBNo").Cells.Item(row).Specific;
                    oForm.DataSources.UserDataSources.Item("CurrWBNo").Value = oEditTextHeader.Value;

                    HeadWB = oEditTextHeader.Value;

                    oEditTextHeader = (SAPbouiCOM.EditText)oMatrix.Columns.Item("WBSupTIN").Cells.Item(row).Specific;
                    oForm.DataSources.UserDataSources.Item("CurrBP").Value = BusinessPartners.GetCardCodeByTin(oEditTextHeader.Value, "S", out cardName);

                    oEditTextHeader = (SAPbouiCOM.EditText)oMatrix.Columns.Item("WBID").Cells.Item(row).Specific;
                    oForm.DataSources.UserDataSources.Item("CurrWBID").Value = oEditTextHeader.Value;

                    oForm.DataSources.UserDataSources.Item("CurrRow").Value = row.ToString();

                    oEditTextHeader = (SAPbouiCOM.EditText)oMatrix.Columns.Item("WBActDate").Cells.Item(row).Specific;
                    oForm.DataSources.UserDataSources.Item("CurrDate").Value = oEditTextHeader.Value;

                    //for (int i = 1; i <= oMatrix.RowCount; i++)
                    //{
                    //    oMatrix.CommonSetting.SetRowBackColor(i, FormsB1.getLongIntRGB(231, 231, 231));
                    //}

                    //oMatrix.CommonSetting.SetRowBackColor(row, FormsB1.getLongIntRGB(255, 255, 153));

                    SAPbouiCOM.ComboBox oCombobox = (SAPbouiCOM.ComboBox)oMatrix.Columns.Item("WBStat").Cells.Item(row).Specific;
                    oForm.DataSources.UserDataSources.Item("CurrWBSt").Value = oCombobox.Value;
                }

                if (GoodsWB == HeadWB)
                {
                    return;
                }

                //Diagnostics
                //Program.uiApp.StatusBar.SetSystemMessage("Time needed before Web service " + Stopwatch.ElapsedMilliseconds + " MiliSec");
                //Stopwatch.Restart();
                //Diagnostics

                oDataTable.Rows.Clear();

                oMatrixGoods.Clear();
                oMatrixGoods.LoadFromDataSource();
                oMatrixGoods.AutoResizeColumns();

                string WbNumber = oMatrix.GetCellSpecific("WBNo", row).String;
                int WbID = Int32.Parse(oMatrix.GetCellSpecific("WBID", row).String);

                Dictionary<string, string> rsSettings = CompanyDetails.getRSSettings(out errorText);
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

                string[] array_HEADER;
                string[][] array_GOODS, array_SUB_WAYBILLS;
                int returnCode = oWayBill.get_waybill(WbID, out array_HEADER, out array_GOODS, out array_SUB_WAYBILLS, out errorText);

                string[][] wbTempTable = null;

                if (wbTempLines.TryGetValue(WbNumber, out wbTempTable))
                {
                    array_GOODS = wbTempTable;
                }


                int headRow = oMatrix.GetNextSelectedRow();
                SAPbouiCOM.EditText BPIDEdit = oMatrix.Columns.Item("WBSupTIN").Cells.Item(row).Specific;
                string BPID = BPIDEdit.Value;
                oForm.DataSources.UserDataSources.Item("CurrBP").Value = BusinessPartners.GetCardCodeByTin(BPID, "S", out cardName);

                SAPbouiCOM.EditText EditCell = oMatrix.Columns.Item("WBNo").Cells.Item(row).Specific;
                oForm.DataSources.UserDataSources.Item("CurrWBNo").Value = EditCell.Value;

                EditCell = oMatrix.Columns.Item("WBID").Cells.Item(row).Specific;
                oForm.DataSources.UserDataSources.Item("CurrWBID").Value = EditCell.Value;

                oForm.DataSources.UserDataSources.Item("CurrRow").Value = row.ToString();

                SAPbouiCOM.ComboBox Combobox = oMatrix.Columns.Item("WBStat").Cells.Item(row).Specific;
                oForm.DataSources.UserDataSources.Item("CurrWBSt").Value = Combobox.Value;

                EditCell = oMatrix.Columns.Item("WBActDate").Cells.Item(row).Specific;
                oForm.DataSources.UserDataSources.Item("CurrDate").Value = EditCell.Value;

                //Diagnostics
                //Program.uiApp.StatusBar.SetSystemMessage("Time needed to prepare for loop " + Stopwatch.ElapsedMilliseconds + " MiliSec");
                //Stopwatch.Restart();
                //Diagnostics


                //Parallel.ForEach()         


                Object locker = new Object();

                string XML = "";
                XML = oDataTable.GetAsXML();
                XML = XML.Replace("<Rows/></DataTable>", "");

                StringBuilder Sbuilder = new StringBuilder();
                Sbuilder.Append(XML);
                Sbuilder.Append("<Rows>");

                string WBNo = "";
                string WBBarcode = "";
                string RSVatCode = "";
                string WBPrjCode = "";
                string WBItmName = "";
                string WBGUntName = "";
                string WBGUntCode = "";
                string WBUntCdRS = "";
                string WbUntNmRS = "";
                string Cardcode = null;
                string DistNumber = "";

                Cardcode = BusinessPartners.GetCardCodeByTin(BPID, "S", out cardName);

                if (Cardcode == null)
                {
                    Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("BPNotFound") + BDOSResources.getTranslate("BPTin") + " : " + BPID);
                    return;
                }




                SAPbobsCOM.Recordset CatalogEntry = null;
                SAPbobsCOM.Recordset oRecordsetbyRSCODE = null;

                SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                SAPbobsCOM.Recordset oRecordSetBN = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                SAPbobsCOM.Recordset oRecordPrevPrice = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                Dictionary<string, string> activeDimensionsList = CommonFunctions.getActiveDimensionsList(out errorText);

                string apInvoice = oMatrix.GetCellSpecific("APInvoice", row).Value;
                string queryBN = "";

                if (!string.IsNullOrEmpty(apInvoice))
                {
                    queryBN = @"Select Distinct OBTN.""DistNumber""

                                From PCH1 Inner Join

                                OPCH on OPCH.""DocEntry"" = PCH1.""DocEntry"" Inner Join

                                ITL1 on ITL1.""ItemCode"" = PCH1.""ItemCode"" Inner Join

                                OITL ON ITL1.""LogEntry"" = OITL.""LogEntry"" INNER JOIN

                                OBTN ON ITL1.""ItemCode"" = OBTN.""ItemCode"" AND ITL1.""SysNumber"" = OBTN.""SysNumber""

                                and OITL.""DocLine"" = PCH1.""LineNum"" AND OITL.""DocEntry"" = PCH1.""DocEntry""

                                WHERE OPCH.""DocEntry"" = '" + apInvoice + "'";

                    oRecordSetBN.DoQuery(queryBN);
                }

                string apCreditMemo= oMatrix.GetCellSpecific("CredMemo", row).Value;

                //foreach (string[] goodsRow in array_GOODS)
                for (int i = 0; i < array_GOODS.Length; i++)
                //Parallel.ForEach(array_GOODS, goodsRow =>
                {
                    string[] goodsRow = array_GOODS[i];

                    CatalogEntry = null;
                    oRecordsetbyRSCODE = null;

                    WBNo = WbNumber;
                    WBBarcode = goodsRow[6] == null ? "" : Regex.Replace(goodsRow[6], @"\t|\n|\r|'", "").Trim();
                    WBItmName = goodsRow[1];
                    WBGUntName = "";
                    WBGUntCode = "";
                    WBUntCdRS = goodsRow[2];
                    WbUntNmRS = string.IsNullOrEmpty(goodsRow[13]) ? oWayBill.get_waybill_unit_name_by_code(WBUntCdRS) : goodsRow[13];
                    RSVatCode = goodsRow[8];
                    WBPrjCode = goodsRow.Length > 12 ? goodsRow[12] : "";

                    string ItemCode = "";
                    string ItemName = "";

                    ItemCode = findItemByNameOITM(WBItmName, WBBarcode, Cardcode, out ItemName);
                    if (ItemName == null) ItemName = "";
                    CatalogEntry = BDO_BPCatalog.getCatalogEntryByBPBarcode(Cardcode, WBItmName, WBBarcode);

                    if (CatalogEntry != null)
                    {
                        SAPbobsCOM.Recordset orec = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        ItemCode = CatalogEntry.Fields.Item("ItemCode").Value;
                        string queryForName = "select \"ItemName\" from OITM " + "\n" + "where \"ItemCode\"= N'" + ItemCode + "'";
                        orec.DoQuery(queryForName);
                        if (!orec.EoF) ItemName = orec.Fields.Item("ItemName").Value;
                        WBGUntCode = CatalogEntry.Fields.Item("U_BDO_UoMCod").Value;
                    }

                    oRecordsetbyRSCODE = BDO_RSUoM.getUomByRSCode(ItemCode, WBUntCdRS, out errorText);

                    if (oRecordsetbyRSCODE != null)
                    {
                        if (WBGUntCode == "")
                        {
                            WBGUntCode = oRecordsetbyRSCODE.Fields.Item("UomCode").Value;
                        }
                    }
                    string query;

                    query = @"SELECT ""UomName"" FROM ""OUOM""WHERE ""UomCode"" = N'" + WBGUntCode + "'";

                    oRecordSet.DoQuery(query);

                    if (!oRecordSet.EoF)
                    {
                        WBGUntName = oRecordSet.Fields.Item("UomName").Value;
                    }

                    string strWBQty = goodsRow[3];
                    string strWBPrice = goodsRow[4];
                    string strWBSum = goodsRow[5];

                    decimal price = CommonFunctions.roundAmountByGeneralSettings(FormsB1.cleanStringOfNonDigits(strWBSum) / FormsB1.cleanStringOfNonDigits(strWBQty), "Price");
                    decimal prevPrice = 0;
                    if (WBItmName.Length > 254)
                    {
                        WBItmName = WBItmName.Substring(0, 254);
                    }

                    if (!oRecordSetBN.EoF)
                    {
                        DistNumber = oRecordSetBN.Fields.Item("DistNumber").Value;
                        oRecordSetBN.MoveNext();
                    }
                    String priceQuery = "select TOP 1 * from( " + "\n" + "select OPDN.\"CardCode\",PDN1.\"PriceAfVAT\" as \"Price\",PDN1.\"DocDate\" as \"DocDatee\" from OIVL " + "\n" + "left join OPDN on OIVL.\"CreatedBy\"=OPDN.\"DocEntry\" " + "\n"
                    + "join PDN1 on OPDN.\"DocEntry\"=PDN1.\"DocEntry\" " + "\n" + "where OIVL.\"TransType\"='20' and OPDN.\"CardCode\"= N'" + Cardcode + "' and PDN1.\"ItemCode\"= N'" + ItemCode + "' " +
                    "\n" + "union all " + "\n" + "select OPCH.\"CardCode\",PCH1.\"PriceAfVAT\"as \"Price\",PCH1.\"DocDate\" as \"DocDatee\" from OIVL " + "\n"
                    + "left join OPCH on OIVL.\"CreatedBy\"=OPCH.\"DocEntry\" " + "\n" + "join PCH1 on OPCH.\"DocEntry\"=PCH1.\"DocEntry\" " + "\n" + "where OIVL.\"TransType\"='18' and OPCH.\"CardCode\"= N'" + Cardcode + "' and PCH1.\"ItemCode\"= N'" +
                    ItemCode + "' " + "\n" + "order by \"DocDatee\" desc)";
                    oRecordPrevPrice.DoQuery(priceQuery);
                    if (!oRecordPrevPrice.EoF) prevPrice = Convert.ToDecimal(oRecordPrevPrice.Fields.Item("Price").Value);

                    Sbuilder.Append("<Row>");
                    Sbuilder.Append("<Cell> <ColumnUid>#</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, (i + 1).ToString());
                    Sbuilder.Append("</Value></Cell>");

                    Sbuilder.Append("<Cell> <ColumnUid>WBNo</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, WBNo);
                    Sbuilder.Append("</Value></Cell>");

                    Sbuilder.Append("<Cell> <ColumnUid>WBBarcode</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, WBBarcode);
                    Sbuilder.Append("</Value></Cell>");

                    Sbuilder.Append("<Cell> <ColumnUid>RSVatCode</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, RSVatCode);
                    Sbuilder.Append("</Value></Cell>");

                    Sbuilder.Append("<Cell> <ColumnUid>WBItmName</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, WBItmName);
                    Sbuilder.Append("</Value></Cell>");

                    Sbuilder.Append("<Cell> <ColumnUid>WBUntCode</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, WBGUntCode);
                    Sbuilder.Append("</Value></Cell>");

                    Sbuilder.Append("<Cell> <ColumnUid>WbUntNmRS</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, WbUntNmRS);
                    Sbuilder.Append("</Value></Cell>");

                    //Sbuilder.Append("<Cell> <ColumnUid>WBUntName</ColumnUid> <Value>");
                    //Sbuilder = CommonFunctions.AppendXML(Sbuilder, WBGUntName);
                    //Sbuilder.Append("</Value></Cell>");

                    Sbuilder.Append("<Cell> <ColumnUid>ItemCode</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, ItemCode.Substring(0, Math.Min(ItemCode.Length, 20)));
                    Sbuilder.Append("</Value></Cell>");

                    Sbuilder.Append("<Cell> <ColumnUid>ItemName</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, ItemName);
                    Sbuilder.Append("</Value></Cell>");

                    Sbuilder.Append("<Cell> <ColumnUid>DistNumber</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, DistNumber);
                    Sbuilder.Append("</Value></Cell>");

                    if (string.IsNullOrEmpty(apInvoice) && string.IsNullOrEmpty(apCreditMemo))
                    {
                        Sbuilder.Append("<Cell> <ColumnUid>WBPrjCode</ColumnUid> <Value>");
                        Sbuilder = CommonFunctions.AppendXML(Sbuilder, WBPrjCode);
                        Sbuilder.Append("</Value></Cell>");
                    }

                    Sbuilder.Append("<Cell> <ColumnUid>WBQty</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, strWBQty);
                    Sbuilder.Append("</Value></Cell>");

                    Sbuilder.Append("<Cell> <ColumnUid>WBPrice</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, price.ToString(CultureInfo.InvariantCulture));
                    Sbuilder.Append("</Value></Cell>");

                    Sbuilder.Append("<Cell> <ColumnUid>WBLPrice</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, prevPrice.ToString());
                    Sbuilder.Append("</Value></Cell>");

                    Sbuilder.Append("<Cell> <ColumnUid>WBSum</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, strWBSum);
                    Sbuilder.Append("</Value></Cell>");

                    Sbuilder.Append("<Cell> <ColumnUid>WBUntCdRS</ColumnUid> <Value>");
                    Sbuilder = CommonFunctions.AppendXML(Sbuilder, WBUntCdRS);
                    Sbuilder.Append("</Value></Cell>");

                    for (int j = 1; j <= activeDimensionsList.Count; j++)
                    {
                        string currDir = goodsRow[14+j-1] == null ? "" : goodsRow[14+j-1];
                        Sbuilder.Append("<Cell> <ColumnUid>DistrRul"+j.ToString()+"</ColumnUid> <Value>");
                        Sbuilder = CommonFunctions.AppendXML(Sbuilder, currDir);
                        Sbuilder.Append("</Value></Cell>");
                    }


                    Sbuilder.Append("</Row>");

                }

                Sbuilder.Append("</Rows>");
                Sbuilder.Append("</DataTable>");

                XML = Sbuilder.ToString();

                XDocument xdoc = XDocument.Parse(XML);
                XDocument xNewDoc = new XDocument();

                xNewDoc.Add(xdoc.Root);

                xNewDoc.Root.RemoveNodes();
                xNewDoc.Root.Add(xdoc.Root.Elements().OrderBy(e => e.Element("WBBarcode")));

                //xNewDoc.Element("Rows").get

                oDataTable.LoadFromXML(xNewDoc.ToString());

                //Diagnostics
                //Program.uiApp.StatusBar.SetSystemMessage("Time needed for loop " + Stopwatch.ElapsedMilliseconds + " MiliSec");
                //Stopwatch.Restart();
                //Diagnostics

                oMatrixGoods.Clear();
                oMatrixGoods.LoadFromDataSource();
                oMatrixGoods.AutoResizeColumns();

                for (int i = 1; i <= oMatrixGoods.RowCount; i++)
                {
                    SAPbouiCOM.EditText last = (SAPbouiCOM.EditText)oMatrixGoods.Columns.Item("WBLPrice").Cells.Item(i).Specific;
                    SAPbouiCOM.EditText that = (SAPbouiCOM.EditText)oMatrixGoods.Columns.Item("WBPrice").Cells.Item(i).Specific;
                    if (Decimal.Parse(last.Value.ToString()) >= Decimal.Parse(that.Value.ToString()) || Decimal.Parse(last.Value.ToString()) == 0)
                    {
                        oMatrixGoods.CommonSetting.SetCellFontColor(i, 13, FormsB1.getLongIntRGB(1, 110, 3));

                    }
                    else
                    {
                        oMatrixGoods.CommonSetting.SetCellFontColor(i, 13, FormsB1.getLongIntRGB(255, 0, 0));
                    }
                }

                WBGdMatrixRow = 0;

                //Diagnostics
                //Program.uiApp.StatusBar.SetSystemMessage("Time needed to display data in matrix " + Stopwatch.ElapsedMilliseconds + " MiliSec");
                //Stopwatch.Restart();
                //Diagnostics

            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
        }
        public static void resizeItems(SAPbouiCOM.Form oForm, out string errorText)
        {
            errorText = null;

            SAPbouiCOM.Item WBMatrix = oForm.Items.Item("WBMatrix");

            WBMatrix.Height = oForm.Height / 4 + 120;
            WBMatrix.Width = oForm.Width - 20;

            oForm.Items.Item("AddRow").Top = WBMatrix.Top + WBMatrix.Height + 5;

            SAPbouiCOM.Item WBGdMatrix = oForm.Items.Item("WBGdMatrix");
            WBGdMatrix.Top = WBMatrix.Top + WBMatrix.Height + 25;
            WBGdMatrix.Height = oForm.Height - 80 - WBGdMatrix.Top;
            WBGdMatrix.Width = oForm.Width - 20;

            oForm.Items.Item("2").Top = oForm.Height - 80;
            oForm.Items.Item("3").Top = oForm.Height - 80;

            try
            {
                oForm.Items.Item("CreateDocs").Top = oForm.Height - 80;
            }
            catch
            {
            }

            SAPbouiCOM.Matrix oMatrixGoods = (SAPbouiCOM.Matrix)oForm.Items.Item("WBGdMatrix").Specific;
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("WBMatrix").Specific;

            oMatrixGoods.AutoResizeColumns();
            oMatrix.AutoResizeColumns();
        }
        public static string getCardCodeByTIN(string BPID)
        {
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string query = @"SELECT * FROM ""OCRD"" WHERE ""LicTradNum"" = '" + BPID + @"' AND ""CardType"" = 'S'";
            oRecordSet.DoQuery(query);

            if (!oRecordSet.EoF)
            {
                return oRecordSet.Fields.Item("CardCode").Value;
            }
            else
            {
                return "";
            }
        }
        public static string findItemByNameOITM(string WBItmName, string RSBarCode, string CardCode, out string itemName)
        {
            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                SAPbobsCOM.BusinessPartners oBP = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);
                oBP.GetByKey(CardCode);
                string query;

                string searchingParam = oBP.UserFields.Fields.Item("U_BDO_ItmPrm").Value;
                if (string.IsNullOrEmpty(searchingParam) || searchingParam == "-1")
                {
                    Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("PleaseFillSearchingParameterOnThisBusinessPartner"), SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    throw new Exception(BDOSResources.getTranslate("FillItemSearchParameterOnTheBP"));
                }

                itemName = null;
                if (searchingParam == "1") //დასახელებით
                    query = @"SELECT ""ItemCode"", ""ItemName"" FROM ""OITM"" WHERE ""ItemName"" = N'" + WBItmName.Replace("'", "''") + "'";
                else //კოდით
                    query = @"SELECT ""ItemCode"", ""ItemName""  FROM ""OITM"" WHERE ""ItemCode"" = '" + RSBarCode + "'";

                oRecordSet.DoQuery(query);

                if (!oRecordSet.EoF)
                {
                    itemName = oRecordSet.Fields.Item("ItemName").Value;
                    return oRecordSet.Fields.Item("ItemCode").Value;
                }
                return "";
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /*
        public static string findItemByNameOITM(string WBItmName, string RSBarCode, string CardCode)
        {
            SAPbobsCOM.BusinessPartners oBP;
            oBP = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);
            oBP.GetByKey(CardCode);

            string searchingParam = oBP.UserFields.Fields.Item("U_BDO_ItmPrm").Value;

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string query;

            if (WBItmName.Length > 100)
            {
                WBItmName = WBItmName.Substring(0, 100);
            }

            if (searchingParam == "1") //დასახელებით
            {
                query = @"SELECT ""ItemCode"" FROM ""OITM"" WHERE ""ItemName"" = N'" + WBItmName.Replace("'", "''") + "'";
            }
            else //კოდით
            {
                query = @"SELECT ""ItemCode"" FROM ""OITM"" WHERE ""ItemCode"" = '" + RSBarCode + "'";
            }

            oRecordSet.DoQuery(query);

            if (!oRecordSet.EoF)
            {
                return oRecordSet.Fields.Item("ItemCode").Value;
            }
            else
            {
                return "";
            }
        }
        */
        public static void updateBPCatalog(SAPbouiCOM.Form oForm, int row)
        {
            try
            {
                //საქონლის ცხრილის მონაცემები
                SAPbouiCOM.Matrix oMatrixGoods = ((SAPbouiCOM.Matrix)(oForm.Items.Item("WBGdMatrix").Specific));

                SAPbouiCOM.EditText ItemCodeEdit = oMatrixGoods.Columns.Item("ItemCode").Cells.Item(row).Specific;
                string ItemCode = ItemCodeEdit.Value;

                SAPbouiCOM.EditText BarCodeEdit = oMatrixGoods.Columns.Item("WBBarcode").Cells.Item(row).Specific;
                string RSBarCode = BarCodeEdit.Value == null ? "" : Regex.Replace(BarCodeEdit.Value, @"\t|\n|\r|'", "").Trim();

                SAPbouiCOM.EditText ItmNameEdit = oMatrixGoods.Columns.Item("WBItmName").Cells.Item(row).Specific;
                string RSItmName = ItmNameEdit.Value;

                SAPbouiCOM.EditText UomCodeEdit = oMatrixGoods.Columns.Item("WBUntCode").Cells.Item(row).Specific;
                string UomCode = UomCodeEdit.Value;

                //მომწოდებელი
                SAPbouiCOM.Matrix oMatrix = ((SAPbouiCOM.Matrix)(oForm.Items.Item("WBMatrix").Specific));

                int CurrRow = Convert.ToInt32(oForm.DataSources.UserDataSources.Item("CurrRow").Value);
                SAPbouiCOM.EditText BPIDEdit = oMatrix.Columns.Item("WBSupTIN").Cells.Item(CurrRow).Specific;
                string BPID = BPIDEdit.Value;
                string cardName;
                string CardCode = BusinessPartners.GetCardCodeByTin(BPID, "S", out cardName);
                if (CardCode == null)
                {
                    throw new Exception(BDOSResources.getTranslate("BPNotFound") + BDOSResources.getTranslate("BPTin") + " : " + BPID);
                }
                SAPbobsCOM.BusinessPartners oBP;
                oBP = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);
                oBP.GetByKey(CardCode);

                string searchingParam = oBP.UserFields.Fields.Item("U_BDO_ItmPrm").Value;
                if (string.IsNullOrEmpty(searchingParam) || searchingParam == "-1")
                {
                    Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("PleaseFillSearchingParameterOnThisBusinessPartner"), SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    throw new Exception(BDOSResources.getTranslate("FillItemSearchParameterOnTheBP"));
                }

                SAPbobsCOM.AlternateCatNum oACN = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oAlternateCatNum);

                SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                string query;

                if (searchingParam == "1") //დასახელებით
                {
                    query = @"SELECT * FROM ""OSCN"" WHERE ""U_BDO_SubDsc"" = N'" + RSItmName.Replace("'", "''") + @"' AND ""CardCode"" = N'" + CardCode + "'";
                }
                else //კოდით
                {
                    query = @"SELECT * FROM ""OSCN"" WHERE ""Substitute"" = N'" + RSBarCode + @"' AND ""CardCode"" = N'" + CardCode + "'";
                }

                oRecordSet.DoQuery(query);

                while (!oRecordSet.EoF)
                {
                    oACN.GetByKey(oRecordSet.Fields.Item("ItemCode").Value, oRecordSet.Fields.Item("CardCode").Value, oRecordSet.Fields.Item("Substitute").Value);
                    oACN.Remove();

                    oRecordSet.MoveNext();
                }

                oACN.GetByKey(ItemCode, CardCode, RSBarCode.Replace("'", ""));

                string Operation = "update";

                if (oACN.ItemCode == "")
                {
                    Operation = "add";
                }
                oACN.CardCode = CardCode;
                oACN.ItemCode = ItemCode;
                oACN.Substitute = string.IsNullOrEmpty(RSBarCode) ? (RSItmName.Length > 50 ? RSItmName.Substring(0, 50) : RSItmName) : RSBarCode.Replace("'", "");

                if (RSItmName.Length > 254)
                {
                    RSItmName = RSItmName.Substring(0, 254);
                }

                oACN.UserFields.Fields.Item("U_BDO_SubDsc").Value = RSItmName;
                oACN.UserFields.Fields.Item("U_BDO_UoMCod").Value = UomCode;

                int errorCode;
                string errorDesc;
                if (Operation == "add")
                {
                    errorCode = oACN.Add();
                }
                else
                {
                    errorCode = oACN.Update();
                }

                if (errorCode != 0)
                {
                    Program.oCompany.GetLastError(out errorCode, out errorDesc);
                    Program.uiApp.StatusBar.SetText(errorDesc, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }

            }
            catch (Exception ex)
            {
                Program.uiApp.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        public static void uiApp_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string errorText = null;

            if (FormUID == "BDO_WaybillsReceivedNewRowForm")
            {
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && !pVal.BeforeAction)
                {
                    SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                    if (pVal.ItemUID == "1")
                    {
                        string newQty = oForm.Items.Item("newQty").Specific.Value;
                        WBGdMatrixNewQty = FormsB1.cleanStringOfNonDigits(newQty);

                        if (WBGdMatrixNewQty >= WBGdMatrixMaxQty)
                        {
                            Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("Error") + ", " + BDOSResources.getTranslate("QuantityShouldBeLessThan") + ": " + WBGdMatrixMaxQty);
                        }
                        else if (WBGdMatrixNewQty == 0)
                        {
                            oForm.Close();
                        }
                        else
                        {
                            oForm.Close();
                            addRow(out errorText);
                        }
                    }
                }
            }
            else
            {
                if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
                {
                    SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);

                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && !pVal.BeforeAction)
                    {
                        if (pVal.ItemUID == "10")
                            updateForm(oForm, out errorText);
                        else if (pVal.ItemUID == "stAddEm")
                        {
                            if (oForm.ActiveItem == "StartAdd")
                                oForm.Items.Item("CarNo").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            oForm.Items.Item("StartAdd").Enabled = ((SAPbouiCOM.CheckBox)oForm.Items.Item("stAddEm").Specific).Checked;
                        }
                        else if (pVal.ItemUID == "endAddEm")
                        {
                            if (oForm.ActiveItem == "EndAdd")
                                oForm.Items.Item("CarNo").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            oForm.Items.Item("EndAdd").Enabled = ((SAPbouiCOM.CheckBox)oForm.Items.Item("endAddEm").Specific).Checked;
                        }
                        else if (pVal.ItemUID == "AddRow")
                        {
                            if (WBGdMatrixRow > 0)
                            {
                                SAPbouiCOM.Form noForm = null;
                                createFormNewRow(noForm, out errorText);
                            }
                        }
                        else if (pVal.ItemUID == "3")
                        {
                            string DocType = oForm.DataSources.UserDataSources.Item("DocType").Value;
                            if (Program.oIncWaybDocFormAPInv != null || Program.oIncWaybDocFormCrMemo != null || Program.oIncWaybDocFormGdsRecpPO != null || Program.oIncWaybDocFormAPCorInv != null)
                            {
                                if (DocType == "1")
                                {
                                    BDO_WBReceivedDocs.attachWBToDoc(oForm, Program.oIncWaybDocFormAPInv);
                                }
                                else if (DocType == "2")
                                {
                                    BDO_WBReceivedDocs.attachWBToDoc(oForm, Program.oIncWaybDocFormCrMemo);
                                }
                                else if (DocType == "3")
                                {
                                    BDO_WBReceivedDocs.attachWBToDoc(oForm, Program.oIncWaybDocFormGdsRecpPO);
                                }
                                else if (DocType == "4")
                                {
                                    BDO_WBReceivedDocs.attachWBToDoc(oForm, Program.oIncWaybDocFormAPCorInv);
                                }
                            }
                            else
                            {
                                oForm.Close();
                            }
                        }

                        else if (pVal.ItemUID == "CreateDocs")
                        {
                            if(Program.WBAUT!="2")
                            {
                                Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("WBAutError"));
                                return;

                            }

                            int answer = 0;

                            answer = Program.uiApp.MessageBox(BDOSResources.getTranslate("CreatePurchaseDocuments") + "?", 1, BDOSResources.getTranslate("Yes"), BDOSResources.getTranslate("No"), "");

                            if (answer == 1)
                            {
                                createWaybillIncDocs(oForm, out errorText);
                            }
                        }
                    }

                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_RESIZE && !pVal.BeforeAction)
                    {
                        resizeItems(oForm, out errorText);
                    }

                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && !pVal.BeforeAction && pVal.ItemUID == "WBMatrix")
                    {
                        int row = pVal.Row;

                        oForm.Freeze(true);
                        fillWBGoods(oForm, row, false, out errorText);

                        if (pVal.ColUID == "WBCheckbox")
                        {
                            var wbMatrix = (Matrix)oForm.Items.Item("WBMatrix").Specific;
                            var isChecked = wbMatrix.GetCellSpecific("WBCheckbox", row).Checked;
                            var wbGDMatrix = (Matrix)oForm.Items.Item("WBGdMatrix").Specific;
                            if (isChecked)
                            {
                                var wbProject = wbMatrix.GetCellSpecific("WBProject", pVal.Row).Value;
                                var wbWarehouse = wbMatrix.GetCellSpecific("WBWhs", pVal.Row).Value;
                                var project = oForm.Items.Item("PrjCode").Specific.Value;
                                var warehouse = oForm.Items.Item("Whs").Specific.Value;

                                project = getProjectByWhs(warehouse);

                                string warehouseADD= BDOSWarehouseAddresses.GetWhsByAddress(wbMatrix.GetCellSpecific("WBEndAdd", pVal.Row).Value);   
                                if (string.IsNullOrEmpty(wbWarehouse) && !string.IsNullOrEmpty(warehouse))
                                {
                                    wbMatrix.GetCellSpecific("WBWhs", pVal.Row).Value = warehouse;
                                }
                                if (!string.IsNullOrEmpty(warehouseADD)) {                           
                                    wbMatrix.GetCellSpecific("WBWhs", pVal.Row).Value = warehouseADD;
                                    project = getProjectByWhs(warehouseADD);
                                }

                                if (!string.IsNullOrEmpty(project))
                                {
                                    wbMatrix.GetCellSpecific("WBProject", pVal.Row).Value = project;
                                }
                                 wbProject = wbMatrix.GetCellSpecific("WBProject", pVal.Row).Value;

                                if (!string.IsNullOrEmpty(wbProject))
                                {
                                    wbGDMatrix.Columns.Item("WBPrjCode").Cells.Item(1).Specific.Value = wbProject;
                                }
                                saveWBGoods(wbGDMatrix,1,false);
                            }
                        }

                        oForm.Freeze(false);
                    }

                    if (pVal.ItemUID == "WBGdMatrix")
                    {
                        if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                        {
                            SAPbouiCOM.ChooseFromListEvent oCFLEvento = (SAPbouiCOM.ChooseFromListEvent)pVal;
                            chooseFromList(oForm, oCFLEvento, pVal.ItemUID, pVal.BeforeAction, pVal.Row, out errorText);

                            if (!pVal.BeforeAction)
                            {
                                if (errorText != "noselectedobjects")
                                {
                                    updateBPCatalog(oForm, pVal.Row);

                                    if (pVal.ColUID =="ItemCode")
                                    {
                                        setUomCodeBtRSCode(oForm, pVal.Row, out errorText);
                                    }
                                    fillWBGoods(oForm, Convert.ToInt32(oForm.DataSources.UserDataSources.Item("CurrRow").Value), true, out errorText);
                                }
                            }
                        }
                        else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_GOT_FOCUS)
                        {
                            if (!pVal.BeforeAction)
                            {
                                if (pVal.ColUID == "ItemCode") //Item No.
                                {
                                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(pVal.ItemUID).Specific;
                                    itemCodeOld = oMatrix.GetCellSpecific(pVal.ColUID, pVal.Row).Value;
                                }
                            }
                        }
                        else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_LOST_FOCUS)
                        {
                            if (!pVal.BeforeAction)
                            {
                                if (pVal.ColUID == "ItemCode") //Item No.
                                {
                                    oForm.Freeze(true);
                                    try
                                    {
                                        SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(pVal.ItemUID).Specific;
                                        string itemCode = oMatrix.GetCellSpecific(pVal.ColUID, pVal.Row).Value;
                                        if (itemCode != itemCodeOld && string.IsNullOrEmpty(itemCode))
                                        {
                                            int rowIndex = pVal.Row;
                                            LanguageUtils.IgnoreErrors<string>(() => oMatrix.Columns.Item("WBUntCode").Cells.Item(rowIndex).Specific.Value = "");
                                            //LanguageUtils.IgnoreErrors<string>(() => oMatrix.Columns.Item("WbUntName").Cells.Item(rowIndex).Specific.Value = "");
                                            LanguageUtils.IgnoreErrors<string>(() => oMatrix.Columns.Item("ItemName").Cells.Item(rowIndex).Specific.Value = "");
                                            itemCodeOld = null;
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        itemCodeOld = null;
                                        throw new Exception(ex.Message);
                                    }
                                    finally
                                    {
                                        oForm.Freeze(false);
                                    }
                                }
                            }
                        }
                        else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && !pVal.BeforeAction)
                        {
                            int row = pVal.Row;

                            oForm.Freeze(true);

                            SAPbouiCOM.Matrix oMatrixGoods = (SAPbouiCOM.Matrix)oForm.Items.Item("WBGdMatrix").Specific;

                            WBGdMatrixRow = row;

                            if (oMatrixGoods.RowCount > 0)
                            {
                                oForm.Freeze(false);
                                for (int i = 1; i <= oMatrixGoods.RowCount; i++)
                                {
                                    oMatrixGoods.CommonSetting.SetRowBackColor(i, FormsB1.getLongIntRGB(231, 231, 231));
                                }

                                try
                                {
                                    oMatrixGoods.CommonSetting.SetRowBackColor(row, FormsB1.getLongIntRGB(255, 255, 153));

                                    WBGdMatrixMaxQty = FormsB1.cleanStringOfNonDigits(oMatrixGoods.Columns.Item("WBQty").Cells.Item(WBGdMatrixRow).Specific.Value);

                                }
                                catch
                                {
                                }
                                oForm.Freeze(true);
                            }

                            oForm.Freeze(false);
                        }
                    }

                    if (pVal.EventType == BoEventTypes.et_MATRIX_LINK_PRESSED && pVal.BeforeAction)
                    {
                        if (pVal.ItemUID == "WBMatrix" && (pVal.ColUID == "APInvoice" || pVal.ColUID == "GdsRcpPO"))
                            MatrixColumnSetArrow(oForm, pVal);
                    }

                    if ((pVal.ItemUID == "Whs" || pVal.ItemUID == "PrjCode" || (pVal.ItemUID == "WBMatrix" && (pVal.ColUID == "WBBlankAgr" || pVal.ColUID == "WBWhs" || pVal.ColUID == "WBProject"))) && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                    {
                        SAPbouiCOM.ChooseFromListEvent oCFLEvento = (SAPbouiCOM.ChooseFromListEvent)pVal;
                        chooseFromList(oForm, oCFLEvento, pVal.ItemUID, pVal.BeforeAction, pVal.Row, out errorText);
                    }

                    else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE && !pVal.BeforeAction)
                    {
                        oForm.Freeze(true);
                        SAPbouiCOM.OptionBtn optBtn = oForm.Items.Item("APInv").Specific;
                        optBtn.Selected = true;
                        oForm.Freeze(false);
                    }

                    else if (pVal.EventType == BoEventTypes.et_VALIDATE && !pVal.BeforeAction)
                    {
                        if (pVal.ItemUID == "WBMatrix" && pVal.ColUID == "WBProject" && pVal.ItemChanged)
                        {
                            var wbMatrix = (Matrix)oForm.Items.Item("WBMatrix").Specific;
                            var wbProject = wbMatrix.GetCellSpecific("WBProject", pVal.Row).Value;
                            FillGoodsProject(oForm, wbProject);
                        }
                    }


                    //if (pVal.BeforeAction == false && pVal.ItemUID == "DocAttch" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_COMBO_SELECT && pVal.ItemChanged)
                    //{
                    //    SAPbouiCOM.Form oForm = Program.uiApp.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount);
                    //    SAPbouiCOM.ComboBox oButtonCombo = (SAPbouiCOM.ComboBox)(oForm.Items.Item("DocAttch").Specific);
                    //    string chosen = oButtonCombo.Selected.Value;
                    //    if (chosen == "1")
                    //    {
                    //        string objectType = "18";
                    //        bool multiselection = false;
                    //        string uniqueID_dc_BusinessPartnerCFL = "Doc0_CFL";
                    //        FormsB1.addChooseFromList(oForm, multiselection, objectType, uniqueID_dc_BusinessPartnerCFL);
                    //        SAPbouiCOM.EditText doc = (SAPbouiCOM.EditText)(oForm.Items.Item("DocC").Specific);
                    //        doc.ChooseFromListAlias = "DocEntry";
                    //        doc.ChooseFromListUID = uniqueID_dc_BusinessPartnerCFL;
                    //    }
                    //    if (chosen == "2")
                    //    {
                    //        string objectType = "20";
                    //        bool multiselection = false;
                    //        string uniqueID_dc_BusinessPartnerCFL = "Doc1_CFL";
                    //        FormsB1.addChooseFromList(oForm, multiselection, objectType, uniqueID_dc_BusinessPartnerCFL);
                    //        SAPbouiCOM.EditText doc = (SAPbouiCOM.EditText)(oForm.Items.Item("DocC").Specific);
                    //        doc.ChooseFromListAlias = "DocEntry";
                    //        doc.ChooseFromListUID = uniqueID_dc_BusinessPartnerCFL;


                    //    }
                    //    if (chosen == "3" ||chosen=="0")
                    //    {
                    //        string objectType = "163";
                    //        bool multiselection = false;
                    //        string uniqueID_dc_BusinessPartnerCFL = "Doc2_CFL";
                    //        FormsB1.addChooseFromList(oForm, multiselection, objectType, uniqueID_dc_BusinessPartnerCFL);
                    //        SAPbouiCOM.EditText doc = (SAPbouiCOM.EditText)(oForm.Items.Item("DocC").Specific);
                    //        doc.ChooseFromListAlias = "DocEntry";
                    //        doc.ChooseFromListUID = uniqueID_dc_BusinessPartnerCFL;

                    //    }
                    //}
                }
            }
        }
        
        public static void saveWBGoods(SAPbouiCOM.Matrix oMatrix,int rowNumber,bool chosen)
        {

            string errorText;
            string[][] array_GOODS = null;
            Dictionary<string, string> activeDimensionsList = CommonFunctions.getActiveDimensionsList(out errorText);
            //oMatrix.Columns.Item("WBPrjCode").Cells.Item(oCFLEvento.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
            //oForm.Freeze(false);
            //SAPbouiCOM.EditText WBPrjCodeEdit = (SAPbouiCOM.EditText)oMatrix.Columns.Item("WBPrjCode").Cells.Item(oCFLEvento.Row).Specific;
            array_GOODS = new string[oMatrix.RowCount][];

            //DataRow wbLinesRow = null;

            //try
            //{
            //    WBPrjCodeEdit.Value = WBPrjCode;
            //}
            //catch
            //{
            for (int i = 1; i <= oMatrix.RowCount; i++)
            {
                //    wbLinesRow = wbLines.Rows.Add();

                //    wbLinesRow["WBPrjCode"] = oMatrix.Columns.Item("WBPrjCode").Cells.Item(i).Specific;

                array_GOODS[i - 1] = new string[14+activeDimensionsList.Count+1];
                //array_GOODS[i][0] = oMatrix.Columns.Item("WBNo").Cells.Item(i).Specific;

                array_GOODS[i - 1][1] = oMatrix.GetCellSpecific("WBItmName", i).Value;
                array_GOODS[i - 1][2] = oMatrix.GetCellSpecific("WBUntCdRS", i).Value;
                array_GOODS[i - 1][3] = oMatrix.GetCellSpecific("WBQty", i).Value;
                array_GOODS[i - 1][4] = oMatrix.GetCellSpecific("WBPrice", i).Value;
                array_GOODS[i - 1][5] = oMatrix.GetCellSpecific("WBSum", i).Value;
                array_GOODS[i - 1][6] = oMatrix.GetCellSpecific("WBBarcode", i).Value;
                array_GOODS[i - 1][8] = oMatrix.GetCellSpecific("RSVatCode", i).Value;
                array_GOODS[i - 1][12] = oMatrix.GetCellSpecific("WBPrjCode", i).Value;
                array_GOODS[i - 1][13] = oMatrix.GetCellSpecific("WbUntNmRS", i).Value;

                //    array_GOODS[i][0] = (itemNode.SelectSingleNode("ID") == null) ? "" : itemNode.SelectSingleNode("ID").InnerText;
                //    array_GOODS[i][1] = (itemNode.SelectSingleNode("W_NAME") == null) ? "" : itemNode.SelectSingleNode("W_NAME").InnerText;
                //    array_GOODS[i][2] = (itemNode.SelectSingleNode("UNIT_ID") == null) ? "" : itemNode.SelectSingleNode("UNIT_ID").InnerText;
                //    array_GOODS[i][3] = (itemNode.SelectSingleNode("QUANTITY") == null) ? "" : itemNode.SelectSingleNode("QUANTITY").InnerText;
                //    array_GOODS[i][4] = (itemNode.SelectSingleNode("PRICE") == null) ? "" : itemNode.SelectSingleNode("PRICE").InnerText;
                //    array_GOODS[i][5] = (itemNode.SelectSingleNode("AMOUNT") == null) ? "" : itemNode.SelectSingleNode("AMOUNT").InnerText;
                //    array_GOODS[i][6] = (itemNode.SelectSingleNode("BAR_CODE") == null) ? "" : itemNode.SelectSingleNode("BAR_CODE").InnerText;
                //    array_GOODS[i][7] = (itemNode.SelectSingleNode("A_ID") == null) ? "" : itemNode.SelectSingleNode("A_ID").InnerText;
                //    array_GOODS[i][8] = (itemNode.SelectSingleNode("VAT_TYPE") == null) ? "" : itemNode.SelectSingleNode("VAT_TYPE").InnerText;
                //    array_GOODS[i][9] = (itemNode.SelectSingleNode("QUANTITY_EXT") == null) ? "" : itemNode.SelectSingleNode("QUANTITY_EXT").InnerText;
                //    array_GOODS[i][10] = (itemNode.SelectSingleNode("STATUS") == null) ? "" : itemNode.SelectSingleNode("STATUS").InnerText;
                //    array_GOODS[i][11] = (itemNode.SelectSingleNode("QUANTITY_F") == null) ? "" : itemNode.SelectSingleNode("QUANTITY_F").InnerText;

                for (int k = 1; k <= activeDimensionsList.Count; k++)
                {
                    array_GOODS[i - 1][14 + Convert.ToInt32(k) - 1] = oMatrix.GetCellSpecific("DistrRul" + k, i).Value;
                }
                if(chosen&& i== rowNumber) array_GOODS[i - 1][14 + activeDimensionsList.Count] = "1";
            }

            if (oMatrix.RowCount > 0)
            {
                string[][] wbTempTable = null;

                if (wbTempLines.TryGetValue(oMatrix.GetCellSpecific("WBNo", 1).Value, out wbTempTable))
                {
                    wbTempLines[oMatrix.GetCellSpecific("WBNo", 1).Value] = array_GOODS;
                }
                else
                {
                    wbTempLines.Add(oMatrix.GetCellSpecific("WBNo", 1).Value, array_GOODS);
                }
            }
        }
        private static void MatrixColumnSetArrow(Form oForm, ItemEvent pVal)
        {
            try
            {
                var oMatrix = (Matrix)oForm.Items.Item("WBMatrix").Specific;
                string wbId = oMatrix.Columns.Item("WBID").Cells.Item(pVal.Row).Specific.Value;

                BDO_WBReceivedDocs.GetDraftByWB(wbId, out var linkedDocTypeDraft, out var linkedDocEntryDraft, out var linkedWhsDraft, out var linkedProjectDraft,
                    out var errorText);
                Column oColumn;

                if (pVal.ColUID == "GdsRcpPO")
                {
                    oColumn = oMatrix.Columns.Item(pVal.ColUID);
                    LinkedButton oLink = oColumn.ExtendedObject;

                    if (linkedDocTypeDraft == "GdsRcptDraft")
                    {
                        oLink.LinkedObjectType = "112"; //Draft
                    }

                    else
                    {
                        oLink.LinkedObjectType = "20"; //Goods Receipt
                    }
                }
                else if (pVal.ColUID == "APInvoice")
                {
                    oColumn = oMatrix.Columns.Item(pVal.ColUID);
                    LinkedButton oLink = oColumn.ExtendedObject;

                    if (linkedDocTypeDraft == "APInvoiceDraft")
                    {
                        oLink.LinkedObjectType = "112"; //Draft
                    }

                    else
                    {
                        oLink.LinkedObjectType = "18"; //AP Invoice
                    }
                }

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private static string getProjectByWhs(string warehouse)
        {
            SAPbobsCOM.Recordset oRecordSetWH = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string queryPr = @"SELECT ""U_BDOSPrjCod"" FROM ""OWHS"" WHERE ""WhsCode"" = '" + warehouse + "'";

            oRecordSetWH.DoQuery(queryPr);
            if (!string.IsNullOrEmpty(oRecordSetWH.Fields.Item("U_BDOSPrjCod").Value))
            {
                 return oRecordSetWH.Fields.Item("U_BDOSPrjCod").Value;
            }
            return string.Empty;

        }
    }
}
