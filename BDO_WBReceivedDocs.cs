using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Globalization;
using System.Text.RegularExpressions;

namespace BDO_Localisation_AddOn
{
    static partial class BDO_WBReceivedDocs
    {
        public static void cancel_wb(string WBID, SAPbouiCOM.Form oDocForm, string Type, out string errorText)
        {
            errorText = null;

            try
            {
                int docEntry = Convert.ToInt32(oDocForm.DataSources.DBDataSources.Item(0).GetValue("DocEntry", 0));

                SAPbobsCOM.Documents APInv;

                if (Type == "Invoice")
                {
                    APInv = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices);
                }
                else if (Type == "GoodsReceiptPO")
                {
                    APInv = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes);
                }
                else
                {
                    APInv = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes);
                }

                APInv.GetByKey(docEntry);

                //BDO_WBNo BDO_WBID actDate BDO_WBSt WBrec
                APInv.UserFields.Fields.Item("U_BDO_WBNo").Value = "";
                APInv.UserFields.Fields.Item("U_BDO_WBID").Value = "";
                APInv.UserFields.Fields.Item("U_actDate").Value = "";
                APInv.UserFields.Fields.Item("U_BDO_WBSt").Value = "-1";
                APInv.UserFields.Fields.Item("U_WBrec").Value = "N";

                APInv.Update();
            }
            catch
            {
                Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("UnableToCancelWaybilForThisDocument"));
            }
        }

        public static void confirm_wb(string WBID, SAPbouiCOM.Form oDocForm, out string errorText)
        {
            errorText = null;

            WBID = WBID.Trim();

            Dictionary<string, string> rsSettings = CompanyDetails.getRSSettings(out errorText);
            if (errorText != null)
            {
                oDocForm.Freeze(false);
                return;
            }

            string su = rsSettings["SU"];
            string sp = rsSettings["SP"];
            WayBill oWayBill = new WayBill(rsSettings["ProtocolType"]);

            bool chek_service_user = oWayBill.chek_service_user(su, sp, out errorText);
            if (chek_service_user == false)
            {
                errorText = BDOSResources.getTranslate("ServiceUserPasswordNotCorrect");
                oDocForm.Freeze(false);
                return;
            }

            bool confirmWaybill = oWayBill.confirm_waybill(Convert.ToInt32(WBID), out errorText);

            if (errorText == null)
            {
                Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("WaybillIsConfirmed"), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                SAPbobsCOM.Documents APInv = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices);
                APInv = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices);
                int docEntry = Convert.ToInt32(oDocForm.DataSources.DBDataSources.Item(0).GetValue("DocEntry", 0));

                if (docEntry != 0)
                {
                    int retvals;
                    SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    string query = "SELECT \"DocEntry\",\"U_BDO_WBID\" FROM \"OPCH\"" +
                                    "WHERE ISNULL(\"U_BDO_WBSt\", '') <> '3' AND \"U_BDO_WBID\" = '" + WBID + "'";
                    if (Program.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                    {
                        query = query.Replace("ISNULL", "IFNULL");
                    }
                    //ინვოისებზე სტატუსების განახლება
                    oRecordSet.DoQuery(query);
                    while (!oRecordSet.EoF)
                    {
                        APInv.GetByKey(Convert.ToInt32(oRecordSet.Fields.Item("DocEntry").Value));
                        APInv.UserFields.Fields.Item("U_WBrec").Value = "Y";

                        retvals = APInv.Update();

                        if (retvals != 0)
                        {
                            int errCode;
                            string errMSG;
                            Program.oCompany.GetLastError(out errCode, out errMSG);
                            Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("Error") + " " + errMSG);
                        }

                        oRecordSet.MoveNext();
                    }

                    //კორექტირებებზე სტატუსების განახლება
                    SAPbobsCOM.Documents CredMemo = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes);
                    CredMemo = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes);

                    query = "SELECT \"DocEntry\",\"U_BDO_WBID\" FROM \"ORPC\"" +
                                    " WHERE ISNULL(\"U_BDO_WBSt\", '') <> '3' AND \"U_BDO_WBID\" = '" + WBID + "'";
                    if (Program.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                    {
                        query = query.Replace("ISNULL", "IFNULL");
                    }
                    oRecordSet.DoQuery(query);

                    while (!oRecordSet.EoF)
                    {
                        CredMemo.GetByKey(Convert.ToInt32(oRecordSet.Fields.Item("DocEntry").Value));
                        CredMemo.UserFields.Fields.Item("U_WBrec").Value = "Y";

                        retvals = CredMemo.Update();

                        if (retvals != 0)
                        {
                            int errCode;
                            string errMSG;

                            Program.oCompany.GetLastError(out errCode, out errMSG);
                            Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("Error") + " " + errMSG);
                        }
                        oRecordSet.MoveNext();
                    }

                    //Goods Receipt PO სტატუსების განახლება
                    SAPbobsCOM.Documents GoodsReceiptPO = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes);
                    GoodsReceiptPO = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes);

                    query = "SELECT \"DocEntry\",\"U_BDO_WBID\" FROM \"OPDN\"" +
                                    " WHERE ISNULL(\"U_BDO_WBSt\", '') <> '3' AND \"U_BDO_WBID\" = '" + WBID + "'";
                    if (Program.oCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                    {
                        query = query.Replace("ISNULL", "IFNULL");
                    }
                    oRecordSet.DoQuery(query);

                    while (!oRecordSet.EoF)
                    {
                        GoodsReceiptPO.GetByKey(Convert.ToInt32(oRecordSet.Fields.Item("DocEntry").Value));
                        GoodsReceiptPO.UserFields.Fields.Item("U_WBrec").Value = "Y";

                        retvals = GoodsReceiptPO.Update();

                        if (retvals != 0)
                        {
                            int errCode;
                            string errMSG;

                            Program.oCompany.GetLastError(out errCode, out errMSG);
                            Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("Error") + " " + errMSG);
                        }
                        oRecordSet.MoveNext();
                    }
                    oDocForm.Freeze(false);
                }
            }
            else
            {
                Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("WaybillIsConfirmed"), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
        }

        public static string detectWBStatus(string StatusRS)
        {
            if (StatusRS == "1") return "2";
            else if (StatusRS == "0") return "1";
            else if (StatusRS == "2") return "3";
            else if (StatusRS == "-1") return "4";
            else if (StatusRS == "-2") return "5";

            return "0";
        }

        public static bool waybillsCompare(string WBID, SAPbouiCOM.Form oDocForm, string RSControlType, string DocType, out string errorText)
        {
            errorText = null;
            bool WBCompares = true;

            if (oDocForm.DataSources.DBDataSources.Item(0).GetValue("CANCELED", 0) == "C")
            {
                return true;
            }

            double B1Total = 0;
            double RSTotal = 0;

            int docEntry;
            try
            {
                docEntry = Convert.ToInt32(oDocForm.DataSources.DBDataSources.Item(0).GetValue("DocEntry", 0));
            }
            catch
            {
                docEntry = 0;
            }

            SAPbouiCOM.Matrix oMatrix = oDocForm.Items.Item("38").Specific;
            Dictionary<string, double> FormGoodsAmounts = new Dictionary<string, double>();

            for (int row = 1; row <= oMatrix.RowCount; row++)
            {
                // SAPbouiCOM.EditText Edtfieldtxt = oMatrix.Columns.Item("ItemCode").Cells.Item(row).Specific;
                string formItemCode = oMatrix.GetCellSpecific("1", row).Value;

                if (Items.isStockItem(formItemCode) == false)
                {
                    continue;
                }

                //Edtfieldtxt = oMatrix.Columns.Item(21).Cells.Item(row).Specific;
                string localAmounttxt = oMatrix.GetCellSpecific("288", row).Value;
                localAmounttxt = FormsB1.cleanStringOfNonDigits(localAmounttxt).ToString();

                double formItemAmount = 0;

                if (localAmounttxt != "")
                {
                    formItemAmount = Convert.ToDouble(localAmounttxt, CultureInfo.InvariantCulture);
                }

                if (FormGoodsAmounts.ContainsKey(formItemCode) == false)
                {
                    if (formItemCode != "")
                    {
                        if (DocType == "APInvoice" || DocType == "GoodsReceiptPO")
                        {
                            FormGoodsAmounts.Add(formItemCode, formItemAmount);
                            B1Total = B1Total + formItemAmount;
                        }
                        else if (DocType == "CredMemo")
                        {
                            FormGoodsAmounts.Add(formItemCode, -formItemAmount);
                            B1Total = B1Total - formItemAmount;
                        }
                    }
                }
                else
                {
                    if (DocType == "APInvoice" || DocType == "GoodsReceiptPO")
                    {
                        FormGoodsAmounts[formItemCode] = FormGoodsAmounts[formItemCode] + formItemAmount;
                        B1Total = B1Total + formItemAmount;
                    }
                    else if (DocType == "CredMemo")
                    {
                        FormGoodsAmounts[formItemCode] = FormGoodsAmounts[formItemCode] - formItemAmount;
                        B1Total = B1Total - formItemAmount;
                    }
                }
            }

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            //მოთხოვნა იღებს მოცემულ ზედნადებზე მიმაგრებული დოკუმენტების საქონელს და თანხებს, დაჯგუფებულს (დაბრუნების თანხები აღებულია მინუსით!!! -""RPC1"".""GTotal"")
            //მისამაგრებელ დოკუმენტს ბაზაში არასოდეს არ ვუყურებთ, მონაცემები აღებულია ფორმიდან.
            string Query = @"SELECT ""ItemCode"", SUM(""GTotal"") AS ""GTotal""
            FROM 
            (SELECT ""OPCH"".""DocStatus"",""OPCH"".""DocEntry"",'APInvoice' AS ""Type"", ""OPCH"".""CANCELED"",""OPCH"".""DocNum"",""OPCH"".""DocDate"", 
            ""OPCH"".""DocEntry"" AS ""Entry"",  
            ""PCH1"".""ItemCode"", ""PCH1"".""GTotal""  
            FROM ""OPCH""
            INNER JOIN ""PCH1"" ON ""OPCH"".""DocEntry"" = ""PCH1"".""DocEntry""  
            WHERE (""U_BDO_WBID""='" + WBID + @"')
            AND (""OPCH"".""CANCELED""='N') " + ((DocType == "APInvoice") ? @" AND (""OPCH"".""DocEntry""<>" + docEntry + @")" : "") +

            @"UNION ALL SELECT ""OPDN"".""DocStatus"",""OPDN"".""DocEntry"",'APInvoice' AS ""Type"", ""OPDN"".""CANCELED"",""OPDN"".""DocNum"",""OPDN"".""DocDate"", 
                ""OPDN"".""DocEntry"" AS ""Entry"",  
                ""PDN1"".""ItemCode"", ""PDN1"".""GTotal""  
                FROM ""OPDN""
                INNER JOIN ""PDN1"" ON ""OPDN"".""DocEntry"" = ""PDN1"".""DocEntry""  
                WHERE (""U_BDO_WBID""='" + WBID + @"')
                AND (""OPDN"".""CANCELED""='N') " + ((DocType == "GoodsReceiptPO") ? @" AND (""OPDN"".""DocEntry""<>" + docEntry + @")" : "") +

            @"UNION ALL 
            SELECT ""ORPC"".""DocStatus"", 
               ""ORPC"".""DocEntry"", 
               'CredMemo' AS ""Type"", 
               ""ORPC"".""CANCELED"", 
               ""ORPC"".""DocNum"", 
               ""ORPC"".""DocDate"", 
               ""ORPC"".""DocEntry"" AS ""Entry"", 
               ""RPC1"".""ItemCode"",-""RPC1"".""GTotal"" 
           FROM ""ORPC"" 
             INNER JOIN ""RPC1"" ON ""ORPC"".""DocEntry"" = ""RPC1"".""DocEntry"" WHERE (""U_BDO_WBID""='" + WBID + @"')
             AND (""ORPC"".""CANCELED""='N')" + ((DocType == "CredMemo") ? @" AND (""ORPC"".""DocEntry""<>" + docEntry + "))" : ")") +
           @" GROUP BY ""ItemCode""";

            oRecordSet.DoQuery(Query);

            string ItemCode = "";
            string WBGUntCode;
            double ItemAmnt;

            Dictionary<string, string> rsSettings = CompanyDetails.getRSSettings(out errorText);
            if (errorText != null)
            {
                return false;
            }

            string su = rsSettings["SU"];
            string sp = rsSettings["SP"];
            WayBill oWayBill = new WayBill(rsSettings["ProtocolType"]);

            bool chek_service_user = oWayBill.chek_service_user(su, sp, out errorText);
            if (chek_service_user == false)
            {
                errorText = BDOSResources.getTranslate("ServiceUserPasswordNotCorrect");
                return false;
            }

            string[] array_HEADER;
            string[][] array_GOODS, array_SUB_WAYBILLS;
            int returnCode = oWayBill.get_waybill(Convert.ToInt32(WBID), out array_HEADER, out array_GOODS, out array_SUB_WAYBILLS, out errorText);

            SAPbouiCOM.DBDataSource DocDBSourceOCRD = oDocForm.DataSources.DBDataSources.Item(1);
            string CardCode = DocDBSourceOCRD.GetValue("CardCode", 0);

            Dictionary<string, double> RSGoodsAmounts = new Dictionary<string, double>();

            string TYPE = "";

            try
            {
                TYPE = array_HEADER[1];
            }
            catch
            {
                Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("WaybillTyoeUnknown"));
                return false;
            }

            foreach (string[] goodsRow in array_GOODS)
            {
                string WBBarcode = goodsRow[6] == null ? "" : Regex.Replace(goodsRow[6], @"\t|\n|\r|'", "").Trim();
                string WBItmName = goodsRow[1];

                SAPbobsCOM.Recordset CatalogEntry = BDO_BPCatalog.getCatalogEntryByBPBarcode(CardCode, WBItmName, WBBarcode, out errorText);

                if (CatalogEntry != null)
                {
                    ItemCode = CatalogEntry.Fields.Item("ItemCode").Value;
                    WBGUntCode = CatalogEntry.Fields.Item("U_BDO_UoMCod").Value;
                }

                //if (Items.isStockItem( ItemCode) == false)
                //{
                //    continue;
                //}

                double WBSum = Convert.ToDouble(goodsRow[5], CultureInfo.InvariantCulture);

                //"5" დაბრუნებაა
                if (TYPE == "დაბრუნება")
                {
                    WBSum = WBSum * (-1);
                }

                RSTotal = RSTotal + WBSum;

                if (RSGoodsAmounts.ContainsKey(ItemCode) == false)
                {
                    if (ItemCode != "")
                    {
                        RSGoodsAmounts.Add(ItemCode, WBSum);
                    }
                }
                else
                {
                    RSGoodsAmounts[ItemCode] = RSGoodsAmounts[ItemCode] + WBSum;
                }
            }

            while (!oRecordSet.EoF)
            {
                ItemCode = oRecordSet.Fields.Item("ItemCode").Value;
                ItemAmnt = oRecordSet.Fields.Item("GTotal").Value;

                if (Items.isStockItem(ItemCode) == false)
                {
                    oRecordSet.MoveNext();
                    continue;
                }

                B1Total = B1Total + ItemAmnt;

                //მიმდინარე ფორმაზე მიმდინარე საქონლის თანხა დავუმატოთ ბაზაში დაფიქსირებულს
                if (FormGoodsAmounts.ContainsKey(ItemCode))
                {
                    ItemAmnt = ItemAmnt + FormGoodsAmounts[ItemCode];
                }

                if (RSControlType == "3")
                {
                    if (RSGoodsAmounts.ContainsKey(ItemCode))
                    {
                        if (ItemAmnt != RSGoodsAmounts[ItemCode])
                        {
                            WBCompares = false;
                            Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("ItemCode") + " " + ItemCode + "თანხები არ შეესაბამება");
                        }
                    }
                    else
                    {
                        WBCompares = false;
                        Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("ItemCode") + " " + ItemCode + "თანხები არ შეესაბამება");
                    }
                }
                oRecordSet.MoveNext();
            }

            if (RSTotal != B1Total)
            {
                WBCompares = false;
                Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("ItemsTotalAmountsNotMatch"));
            }

            errorText = null;
            return WBCompares;
        }

        public static void getInvoiceByWB(string WBID, out string DocType, out int DocEntry, out string whs, out string project, out string errorText)
        {
            errorText = null;
            DocType = null;
            DocEntry = 0;
            whs = "";
            project = "";

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oRecordSet.DoQuery("SELECT TOP 1 \"DocStatus\",'APInvoice' AS \"Type\",\"CANCELED\",\"DocNum\",\"OPCH\".\"DocDate\",\"OPCH\".\"DocEntry\" AS \"Entry\",\"PCH1\".\"WhsCode\", \"OPCH\".\"Project\"" +
                                  "FROM \"OPCH\" JOIN \"PCH1\" ON \"OPCH\".\"DocEntry\" = \"PCH1\".\"DocEntry\""
                                  + "WHERE  (\"U_BDO_WBID\"='" + WBID + "') AND (\"CANCELED\"='N')" +
                                  "ORDER BY \"OPCH\".\"DocDate\", \"Entry\" DESC");

                if (!oRecordSet.EoF)
                {
                    DocType = oRecordSet.Fields.Item("Type").Value;
                    DocEntry = oRecordSet.Fields.Item("Entry").Value;
                    whs = oRecordSet.Fields.Item("WhsCode").Value;
                    project = oRecordSet.Fields.Item("Project").Value;
                }
                else
                {
                    DocType = "";
                    DocEntry = 0;
                }
            }
            catch (Exception ex)
            {
                int errCode;
                string errMsg;

                Program.oCompany.GetLastError(out errCode, out errMsg);
                errorText = BDOSResources.getTranslate("ErrorDescription") + " : " + errMsg + "! " + BDOSResources.getTranslate("Code") + " : " + errCode + "! " + BDOSResources.getTranslate("OtherInfo") + " : " + ex.Message;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                GC.Collect();
            }
        }

        public static void getGoodsReceipePOByWB(string WBID, out string DocType, out int DocEntry, out string whs, out string project, out string errorText)
        {
            errorText = null;
            DocType = null;
            DocEntry = 0;
            whs = "";
            project = "";

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                string query = @"SELECT
                                    TOP 1
	                                ""DocStatus"",
	                                'APInvoice' AS ""Type"",
	                                ""CANCELED"",
	                                ""DocNum"",
	                                ""OPDN"".""DocDate"",
	                                ""OPDN"".""DocEntry"" AS ""Entry"",
                                    ""PDN1"".""WhsCode"",
                                    ""OPDN"".""Project""
                                FROM ""OPDN"" JOIN ""PDN1"" ON ""OPDN"".""DocEntry"" = ""PDN1"".""DocEntry""
                                WHERE (""U_BDO_WBID""='" + WBID + @"') 
                                AND (""CANCELED""='N') 
                                ORDER BY ""OPDN"".""DocDate"",
	                                 ""Entry"" DESC";
                oRecordSet.DoQuery(query);

                if (!oRecordSet.EoF)
                {
                    DocType = oRecordSet.Fields.Item("Type").Value;
                    DocEntry = oRecordSet.Fields.Item("Entry").Value;
                    whs = oRecordSet.Fields.Item("WhsCode").Value;
                    project = oRecordSet.Fields.Item("Project").Value;
                }
                else
                {
                    DocType = "";
                    DocEntry = 0;
                }
            }
            catch (Exception ex)
            {
                int errCode;
                string errMsg;

                Program.oCompany.GetLastError(out errCode, out errMsg);
                errorText = BDOSResources.getTranslate("ErrorDescription") + " : " + errMsg + "! " + BDOSResources.getTranslate("Code") + " : " + errCode + "! " + BDOSResources.getTranslate("OtherInfo") + " : " + ex.Message;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                GC.Collect();
            }
        }

        public static void GetDraftByWB(string wbId, out string docType, out int docEntry, out string whs, out string project, out string errorText)
        {
            errorText = null;
            docType = null;
            docEntry = 0;
            whs = "";
            project = "";

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                string query = @"SELECT
                                    TOP 1
	                                ""DocStatus"",
	                                CASE WHEN ""ODRF"".""ObjType"" = '18' THEN 'APInvoiceDraft' WHEN ""ODRF"".""ObjType"" = '20' THEN 'GdsRcptDraft' END AS ""Type"",
	                                ""CANCELED"",
	                                ""DocNum"",
	                                ""ODRF"".""DocDate"",
	                                ""ODRF"".""DocEntry"" AS ""Entry"",
                                    ""DRF1"".""WhsCode"",
                                    ""ODRF"".""Project""
                                FROM ""ODRF"" JOIN ""DRF1"" ON ""ODRF"".""DocEntry"" = ""DRF1"".""DocEntry""
                                WHERE (""U_BDO_WBID""='" + wbId + @"') 
                                AND (""CANCELED""='N') 
                                ORDER BY ""ODRF"".""DocDate"",
	                                 ""Entry"" DESC";
                oRecordSet.DoQuery(query);

                if (!oRecordSet.EoF)
                {
                    docType = oRecordSet.Fields.Item("Type").Value;
                    docEntry = oRecordSet.Fields.Item("Entry").Value;
                    whs = oRecordSet.Fields.Item("WhsCode").Value;
                    project = oRecordSet.Fields.Item("Project").Value;
                }
                else
                {
                    docType = "";
                    docEntry = 0;
                }
            }
            catch (Exception ex)
            {
                int errCode;
                string errMsg;

                Program.oCompany.GetLastError(out errCode, out errMsg);
                errorText = BDOSResources.getTranslate("ErrorDescription") + " : " + errMsg + "! " + BDOSResources.getTranslate("Code") + " : " + errCode + "! " + BDOSResources.getTranslate("OtherInfo") + " : " + ex.Message;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                GC.Collect();
            }
        }

        public static void getMemoByWB(string WBID, out string DocType, out int DocEntry, out string whs, out string project, out string errorText)
        {
            errorText = null;
            DocType = null;
            DocEntry = 0;
            whs = "";
            project = "";

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oRecordSet.DoQuery("SELECT TOP 1 \"DocStatus\",'CredMemo' AS \"Type\",\"CANCELED\",\"DocNum\", \"ORPC\".\"DocDate\", \"ORPC\".\"DocEntry\" AS \"Entry\",\"RPC1\".\"WhsCode\", \"ORPC\".\"Project\"" +
                                  "FROM \"ORPC\" JOIN \"RPC1\" ON \"ORPC\".\"DocEntry\" = \"RPC1\".\"DocEntry\" " + "WHERE  (\"U_BDO_WBID\"='" + WBID + "') AND (\"CANCELED\"='N')" +
                                  "ORDER BY \"ORPC\".\"DocDate\", \"Entry\" DESC");
                if (!oRecordSet.EoF)
                {
                    DocType = oRecordSet.Fields.Item("Type").Value;
                    DocEntry = oRecordSet.Fields.Item("Entry").Value;
                    whs = oRecordSet.Fields.Item("WhsCode").Value;
                    project = oRecordSet.Fields.Item("Project").Value;
                }
                else
                {
                    DocType = "";
                    DocEntry = 0;
                }
            }
            catch (Exception ex)
            {
                int errCode;
                string errMsg;

                Program.oCompany.GetLastError(out errCode, out errMsg);
                errorText = BDOSResources.getTranslate("ErrorDescription") + " : " + errMsg + "! " + BDOSResources.getTranslate("Code") + " : " + errCode + "! " + BDOSResources.getTranslate("OtherInfo") + " : " + ex.Message;
            }
            finally
            {
                Marshal.FinalReleaseComObject(oRecordSet);
                GC.Collect();
            }
        }

        public static void attachWBToDoc(SAPbouiCOM.Form oForm, SAPbouiCOM.Form oIncWaybDocForm, out string errorText)
        {
            errorText = null;

            try
            {
                oIncWaybDocForm.Select();
                oIncWaybDocForm.Freeze(true);

                SAPbouiCOM.EditText oEditText = (SAPbouiCOM.EditText)oIncWaybDocForm.Items.Item("BDO_WBNo").Specific;
                oEditText.Value = oForm.DataSources.UserDataSources.Item("CurrWBNo").Value;

                oEditText = (SAPbouiCOM.EditText)oIncWaybDocForm.Items.Item("BDO_WBID").Specific;
                oEditText.Value = oForm.DataSources.UserDataSources.Item("CurrWBID").Value;

                oEditText = (SAPbouiCOM.EditText)oIncWaybDocForm.Items.Item("actDate").Specific;
                //oEditText.Value = Convert.ToDateTime(oForm.DataSources.UserDataSources.Item("CurrDate").Value).ToString("yyyyMMdd");
                oEditText.Value = oForm.DataSources.UserDataSources.Item("CurrDate").Value;

                try
                {
                    oEditText = (SAPbouiCOM.EditText)oIncWaybDocForm.Items.Item("10").Specific;
                    //oEditText.Value = Convert.ToDateTime(oForm.DataSources.UserDataSources.Item("CurrDate").Value).ToString("yyyyMMdd");
                    oEditText.Value = oForm.DataSources.UserDataSources.Item("CurrDate").Value;

                    oEditText = (SAPbouiCOM.EditText)oIncWaybDocForm.Items.Item("4").Specific;
                    oEditText.Value = oForm.DataSources.UserDataSources.Item("CurrBP").Value;
                }
                catch
                {

                }

                SAPbouiCOM.ComboBox oComboBox = (SAPbouiCOM.ComboBox)oIncWaybDocForm.Items.Item("BDO_WBSt").Specific;
                string status = oForm.DataSources.UserDataSources.Item("CurrWBSt").Value;

                setwaybillText(oIncWaybDocForm);
                oForm.Close();

                oIncWaybDocForm.Freeze(false);

                Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("WaybillIsLinkedToDocument"), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                //FormsB1.SimulateRefresh();
                if (oIncWaybDocForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oIncWaybDocForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                }
                oIncWaybDocForm = null;
            }
            catch (Exception ex)
            {
                int errCode;
                string errMsg;

                Program.oCompany.GetLastError(out errCode, out errMsg);
                errorText = BDOSResources.getTranslate("ErrorDescription") + " : " + errMsg + "! " + BDOSResources.getTranslate("Code") + " : " + errCode + "! " + BDOSResources.getTranslate("OtherInfo") + " : " + ex.Message;
            }
            finally
            {
                if (oIncWaybDocForm != null)
                {
                    oIncWaybDocForm.Freeze(false);
                }
                GC.Collect();
            }

            if (errorText != null)
            {
                Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("WaybiliiNotLinkedToDocument") + " " + errorText, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        public static void createUserFields(string DocType, out string errorText)
        {
            errorText = null;
            Dictionary<string, object> fieldskeysMap;

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDO_WBNo");
            fieldskeysMap.Add("TableName", DocType);
            fieldskeysMap.Add("Description", "Waybill Number");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDO_WBSt");
            fieldskeysMap.Add("TableName", DocType);
            fieldskeysMap.Add("Description", "Waybill Status");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            List<string> statusValues = new List<string>();
            statusValues.Add(""); //-1
            statusValues.Add("Saved"); //1
            statusValues.Add("Active"); //2
            statusValues.Add("Finished"); //3
            statusValues.Add("Deleted"); //4
            statusValues.Add("Cancelled"); //5
            statusValues.Add("Sent To Transporter"); //6
            fieldskeysMap.Add("ValidValues", statusValues);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDO_WBID");
            fieldskeysMap.Add("TableName", DocType);
            fieldskeysMap.Add("Description", "Waybill ID");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "actDate");
            fieldskeysMap.Add("TableName", DocType);
            fieldskeysMap.Add("Description", "Activate Date");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Date);
            fieldskeysMap.Add("Visible", false);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "WBrec");
            fieldskeysMap.Add("TableName", DocType);
            fieldskeysMap.Add("Description", "Waybill Received");
            fieldskeysMap.Add("EditSize", 1);
            fieldskeysMap.Add("DefaultValue", "N");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);

            UDO.addUserTableFields(fieldskeysMap, out errorText);

            GC.Collect();
        }

        public static void createFormItems(SAPbouiCOM.Form oForm, string Doctype, out string errorText)
        {
            oForm.Freeze(true);
            errorText = null;
            Dictionary<string, object> formItems = null;

            string itemName = "";

            SAPbouiCOM.Item cancelButton = oForm.Items.Item("2");

            double left = cancelButton.Left + cancelButton.Width + 10;
            double Top = cancelButton.Top;
            double Width = cancelButton.Width;
            double Height = cancelButton.Height;

            formItems = new Dictionary<string, object>();
            List<string> listValidValues = null;

            listValidValues = new List<string>();
            listValidValues.Add(BDOSResources.getTranslate("AttachToWaybill"));
            listValidValues.Add(BDOSResources.getTranslate("ReceiveWaybill"));
            listValidValues.Add(BDOSResources.getTranslate("Cancel"));

            formItems = new Dictionary<string, object>();
            itemName = "WBOper";
            formItems.Add("Caption", BDOSResources.getTranslate("RSOperations"));
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_BUTTON_COMBO);
            formItems.Add("Left", left);
            formItems.Add("Width", 100);
            formItems.Add("Top", Top);
            formItems.Add("Height", Height);
            formItems.Add("UID", itemName);
            formItems.Add("ValidValues", listValidValues);
            formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
            formItems.Add("AffectsFormMode", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            //WB NO
            SAPbouiCOM.Item itemDocDateSt = oForm.Items.Item("86");
            left = itemDocDateSt.Left;
            Height = itemDocDateSt.Height;
            Top = itemDocDateSt.Top + Height + 1;
            Width = itemDocDateSt.Width;

            formItems = new Dictionary<string, object>();
            itemName = "BDO_WBNoST";
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left);
            formItems.Add("Width", Width);
            formItems.Add("Top", Top);
            formItems.Add("Height", Height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", "№");
            formItems.Add("Visible", false);

            //FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            //WB NO
            itemDocDateSt = oForm.Items.Item("86");

            left = itemDocDateSt.Left;
            Top = itemDocDateSt.Top + Height * 1.5 + 1;
            Width = itemDocDateSt.Width;
            Height = itemDocDateSt.Height;

            formItems = new Dictionary<string, object>();
            itemName = "WBInfoST";
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left);
            formItems.Add("Width", 250);
            formItems.Add("Top", Top);
            formItems.Add("Height", Height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("NotLinked"));
            formItems.Add("TextStyle", 4);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }


            //WB Status           
            Top = Top + Height + 2;

            formItems = new Dictionary<string, object>();
            itemName = "BDO_WBStST";
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left);
            formItems.Add("Width", Width);
            formItems.Add("Top", Top);
            formItems.Add("Height", Height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("Status"));
            formItems.Add("Visible", false);

            //FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            //ID
            Top = Top + Height + 2;

            formItems = new Dictionary<string, object>();
            itemName = "BDO_WBIDT";
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left);
            formItems.Add("Width", Width);
            formItems.Add("Top", Top);
            formItems.Add("Height", Height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", "ID");
            formItems.Add("Visible", false);

            //FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            //Activation Date
            Top = Top + Height + 2;

            formItems = new Dictionary<string, object>();
            itemName = "actDateT";
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            formItems.Add("Left", left);
            formItems.Add("Width", Width);
            formItems.Add("Top", Top);
            formItems.Add("Height", Height);
            formItems.Add("UID", itemName);
            formItems.Add("Caption", BDOSResources.getTranslate("ActivationDate"));
            formItems.Add("Visible", false);

            //FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            SAPbouiCOM.Item itemDocDate = oForm.Items.Item("46");
            left = itemDocDate.Left;
            //Top = itemDocDate.Top + itemDocDate.Height + 2;
            Top = itemDocDate.Top;
            Width = itemDocDate.Width;
            Height = itemDocDate.Height;

            formItems = new Dictionary<string, object>();
            itemName = "BDO_WBNo";
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", Doctype);
            formItems.Add("Alias", "U_BDO_WBNo");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left);
            formItems.Add("Width", Width);
            formItems.Add("Top", Top);
            formItems.Add("Height", Height);
            formItems.Add("UID", itemName);
            formItems.Add("Visible", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            //Top = Top + Height + 2;

            List<string> statusValues = new List<string>();
            statusValues.Add(""); //-1
            statusValues.Add(BDOSResources.getTranslate("Saved")); //1
            statusValues.Add(BDOSResources.getTranslate("Active")); //2
            statusValues.Add(BDOSResources.getTranslate("finished")); //3
            statusValues.Add(BDOSResources.getTranslate("deleted")); //4
            statusValues.Add(BDOSResources.getTranslate("cancelled")); //5
            statusValues.Add(BDOSResources.getTranslate("SentToTransporter")); //6

            formItems = new Dictionary<string, object>();
            itemName = "BDO_WBSt";
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", Doctype);
            formItems.Add("Alias", "U_BDO_WBSt");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            formItems.Add("Left", left);
            formItems.Add("Width", Width);
            formItems.Add("Top", Top);
            formItems.Add("Height", Height);
            formItems.Add("UID", itemName);
            //formItems.Add("Enabled", false);
            formItems.Add("ExpandType", SAPbouiCOM.BoExpandType.et_DescriptionOnly);
            formItems.Add("DisplayDesc", true);
            formItems.Add("Visible", false);
            formItems.Add("ValidValues", statusValues);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            //Top = Top + Height + 2;

            formItems = new Dictionary<string, object>();
            itemName = "BDO_WBID";
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", Doctype);
            formItems.Add("Alias", "U_BDO_WBID");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left);
            formItems.Add("Width", Width);
            formItems.Add("Top", Top);
            formItems.Add("Height", Height);
            formItems.Add("UID", itemName);
            formItems.Add("Visible", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            if (errorText != null)
            {
                return;
            }

            //Top = Top + Height + 2;

            formItems = new Dictionary<string, object>();
            itemName = "actDate";
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", Doctype);
            formItems.Add("Alias", "U_actDate");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            formItems.Add("Left", left);
            formItems.Add("Width", Width);
            formItems.Add("Top", Top);
            formItems.Add("Height", Height);
            formItems.Add("UID", itemName);
            formItems.Add("Visible", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);

            oForm.Freeze(false);

            if (errorText != null)
            {
                return;
            }

            //Top = Top + Height + 2;

            formItems = new Dictionary<string, object>();
            itemName = "WBrec";
            formItems.Add("Caption", BDOSResources.getTranslate("Received"));
            formItems.Add("isDataSource", true);
            formItems.Add("DataSource", "DBDataSources");
            formItems.Add("TableName", Doctype);
            formItems.Add("Alias", "U_WBrec");
            formItems.Add("ValOff", "N");
            formItems.Add("ValOn", "Y");
            formItems.Add("Bound", true);
            formItems.Add("Type", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            formItems.Add("Left", left);
            formItems.Add("Width", Width);
            formItems.Add("Top", Top);
            formItems.Add("Height", Height);
            formItems.Add("UID", itemName);
            formItems.Add("Visible", false);

            FormsB1.createFormItem(oForm, formItems, out errorText);
            oForm.Freeze(false);

            if (errorText != null)
            {
                return;
            }
        }

        public static void comboSelect(SAPbouiCOM.Form oForm, SAPbouiCOM.Form oIncWaybForm, string Type, out string errorText)
        {
            errorText = null;

            string operationRS = "";
            SAPbouiCOM.ButtonCombo oButtonCombo = (SAPbouiCOM.ButtonCombo)oForm.Items.Item("WBOper").Specific;
            if (oButtonCombo.Selected != null)
            {
                operationRS = oButtonCombo.Selected.Value;
            }
            oForm.Freeze(false);
            oButtonCombo.Caption = BDOSResources.getTranslate("Operations");
            string WBID = oForm.DataSources.DBDataSources.Item(0).GetValue("U_BDO_WBID", 0);

            if (operationRS == "0")
            {
                SAPbobsCOM.BusinessPartners oBP;
                oBP = Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);
                oBP.GetByKey(oForm.DataSources.DBDataSources.Item(0).GetValue("CardCode", 0).Trim());
                if (oBP.UserFields.Fields.Item("LicTradNum").Value.Trim() == "")
                {
                    Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("CheckBPTin"));
                    return;
                }

                if (Type == "Invoice")
                {
                    if (oForm.DataSources.DBDataSources.Item("PCH1").GetValue("BaseType", 0) == "20")
                    {
                        Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("APInvoiceIsCreatedFromGoodsReceiptPO"));
                        return;
                    }
                }
                BDO_WaybillsJournalReceived.createForm(oIncWaybForm, out errorText);
            }
            else if (operationRS == "1")
            {
                if (WBID == "")
                {
                    Program.uiApp.StatusBar.SetSystemMessage(BDOSResources.getTranslate("LinkWaybilToDocument"));
                    return;
                }
                confirm_wb(WBID, oForm, out errorText);
                FormsB1.SimulateRefresh();
            }
            else if (operationRS == "2")
            {
                cancel_wb(WBID, oForm, Type, out errorText);
                FormsB1.SimulateRefresh();
            }
        }

        public static void setwaybillText(SAPbouiCOM.Form oForm)
        {
            string waybillId = oForm.DataSources.DBDataSources.Item(0).GetValue("U_BDO_WBID", 0).Trim();
            string waybillNo = oForm.DataSources.DBDataSources.Item(0).GetValue("U_BDO_WBNo", 0).Trim();
            string waybillStatus = oForm.Items.Item("BDO_WBSt").Specific.Selected?.Description;
            string received = oForm.DataSources.DBDataSources.Item(0).GetValue("U_WBrec", 0).Trim() == "Y" ? BDOSResources.getTranslate("Received") : "";
            string caption = BDOSResources.getTranslate("NotLinked");

            if (!string.IsNullOrEmpty(waybillId))
                caption = BDOSResources.getTranslate("Wb") + ": " + waybillStatus + " ID " + waybillId + (!string.IsNullOrEmpty(waybillNo) ? " № " + waybillNo : "") + " " + received;

            oForm.Items.Item("WBInfoST").Specific.Caption = caption;
        }

        public static void ClearWaybillItemsValues(SAPbouiCOM.Form oForm)
        {
            try
            {
                oForm.Freeze(true);

                oForm.Items.Item("BDO_WBNo").Visible = true;
                oForm.Items.Item("BDO_WBID").Visible = true;
                oForm.Items.Item("actDate").Visible = true;
                oForm.Items.Item("BDO_WBSt").Visible = true;
                oForm.Items.Item("WBrec").Visible = true;

                oForm.Items.Item("BDO_WBNo").Specific.Value = string.Empty;
                oForm.Items.Item("BDO_WBID").Specific.Value = string.Empty;
                oForm.Items.Item("actDate").Specific.Value = "00010101";
                oForm.Items.Item("BDO_WBSt").Specific.Select("-1");
                oForm.Items.Item("WBrec").Specific.Checked = false;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                oForm.ActiveItem = "16";
                oForm.Items.Item("BDO_WBNo").Visible = false;
                oForm.Items.Item("BDO_WBID").Visible = false;
                oForm.Items.Item("actDate").Visible = false;
                oForm.Items.Item("BDO_WBSt").Visible = false;
                oForm.Items.Item("WBrec").Visible = false;

                setwaybillText(oForm);

                oForm.Freeze(false);
            }
        }
    }
}