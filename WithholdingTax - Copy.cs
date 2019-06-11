using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Globalization;
using System.Data;

namespace BDO_Localisation_AddOn
{
    static partial class WithholdingTax
    {
        public static CultureInfo cultureInfo = null;

        public static void JrnEntryAPInvoiceCredidNote(SAPbobsCOM.Company oCompany, SAPbouiCOM.Form oForm, string DocType, string DocEntry, string DocNum, DateTime DocDate, out string errorText)
        {
            errorText = null;

            try
            {
               SAPbouiCOM.DBDataSource DocDBSourceWT = null;

               if( DocType =="204")
               {
                   DocDBSourceWT = oForm.DataSources.DBDataSources.Item("DPO1");
               }
               else if (DocType == "18")
               { 
                   DocDBSourceWT = oForm.DataSources.DBDataSources.Item("PCH1");
               }
               else
               { 
                   DocDBSourceWT = oForm.DataSources.DBDataSources.Item("RPC1");
               }



                if (DocDBSourceWT.Size == 0)
                {
                    return;
                }

                SAPbouiCOM.DBDataSource DocDBSourceOCRD = oForm.DataSources.DBDataSources.Item("OCRD");
                string ECVatGroup = DocDBSourceOCRD.GetValue("ECVatGroup", 0);

                DataTable jeLines = JournalEntry.JournalEntryTable();
                DataRow jeLinesRow = null;


                SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                string query = "SELECT " +
                                "\"OVTG\".\"U_BDOSAccF\" AS \"DebitAccount\", " +
                                "\"OVTG\".\"Account\" AS \"CreditAccount\", " +
                                "SUM(\"PCH1\".\"VatSum\") as \"TaxAmount\"  " +
                                "FROM \"" + oCompany.CompanyDB + (DocType == "18" ? "\".\"OPCH\" " : DocType == "204" ? "\".\"ODPO\" " : "\".\"ORPC\" ") + " AS \"OPCH\"" +
                                "LEFT JOIN \"" + oCompany.CompanyDB + "\".\"OVTG\" ON  \"PCH1\".\"VatGroup\" = \"OVTG\".\"Code\"  " +
                                "WHERE \"PCH1\".\"DocEntry\" = " + DocEntry + " "+ 
                                "GROUP BY \"OVTG\".\"U_BDOSAccF\", \"OVTG\".\"Account\"";

                oRecordSet.DoQuery(query);

                while (!oRecordSet.EoF)
                {
                    jeLinesRow = jeLines.Rows.Add(0);
                    jeLinesRow["AccountCode"] = oRecordSet.Fields.Item("CreditAccount").Value;
                    jeLinesRow["ShortName"] = oRecordSet.Fields.Item("CreditAccount").Value;
                    jeLinesRow["ContraAccount"] = oRecordSet.Fields.Item("DebitAccount").Value;
                    //jeLinesRow["VatGroup"] = oRecordSet.Fields.Item("U_BDOSVatGrp").Value;
                    jeLinesRow["Credit"] = oRecordSet.Fields.Item("TaxAmount").Value;
                    jeLinesRow["Debit"] = 0;

                    jeLinesRow = jeLines.Rows.Add(1);
                    jeLinesRow["AccountCode"] = oRecordSet.Fields.Item("DebitAccount").Value;
                    jeLinesRow["ShortName"] = oRecordSet.Fields.Item("DebitAccount").Value;
                    jeLinesRow["ContraAccount"] = oRecordSet.Fields.Item("CreditAccount").Value;
                    //jeLinesRow["VatGroup"] = oRecordSet.Fields.Item("U_BDOSVatGrp").Value;
                    jeLinesRow["Credit"] = 0;
                    jeLinesRow["Debit"] = oRecordSet.Fields.Item("TaxAmount").Value;

                    oRecordSet.MoveNext();
                }

                JournalEntry.JrnEntry(oCompany, DocEntry, DocType, (DocType == "18" ? "AP Invoicess: " : "AP Credit note: ") + DocNum, DocDate, jeLines, out errorText);
            }
            catch (Exception ex)
            {
                errorText = ex.Message;
            }
        }

        public static void JrnEntryAPInvoiceCredidNoteCheck(SAPbobsCOM.Company oCompany, SAPbouiCOM.Form oForm, string DocType, out string errorText)
        {
            errorText = null;

            try
            {
                SAPbouiCOM.DBDataSource DocDBSource = null;

                if (DocType == "204")
                {
                    DocDBSource = oForm.DataSources.DBDataSources.Item("DPO1");
                }
                else if (DocType == "18")
                {
                    DocDBSource = oForm.DataSources.DBDataSources.Item("PCH1");
                }
                else
                {
                    DocDBSource = oForm.DataSources.DBDataSources.Item("RPC1");
                }

                if (DocDBSource.Size == 0)
                {
                    return;
                }

                for (int i = 0; i < DocDBSource.Size; i++)
                {
                    string VatGroup = DocDBSource.GetValue("VatGroup", i);

                    SAPbobsCOM.VatGroups oVG;
                    oVG = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVatGroups);
                    oVG.GetByKey(VatGroup);

                    string BDOSAccF = oVG.UserFields.Fields.Item("U_BDOSAccF").Value;
                    string TaxAccount = oVG.TaxAccount;

                    if (TaxAccount == "")
                    {
                        errorText = BDOSResources.getTranslate("CheckVatGroupAccounts");
                    }
                    if (BDOSAccF == "")
                    {
                        errorText = BDOSResources.getTranslate("CheckVatGroupAccounts");
                    }
                }

                    
            }
            catch (Exception ex)
                    {
                errorText = ex.Message;
                    }
                }

        public static void createUserFields(SAPbobsCOM.Company oCompany, out string errorText)
        {
            errorText = null;
            Dictionary<string, object> fieldskeysMap;


            fieldskeysMap = new Dictionary<string, object>();
            fieldskeysMap.Add("Name", "BDOSVatGrp");
            fieldskeysMap.Add("TableName", "OWHT");
            fieldskeysMap.Add("Description", "Vat Group");
            fieldskeysMap.Add("Type", SAPbobsCOM.BoFieldTypes.db_Alpha);
            fieldskeysMap.Add("EditSize", 50);

            UDO.addUserTableFields(oCompany, fieldskeysMap, out errorText);

            GC.Collect();
        }
    }
}
